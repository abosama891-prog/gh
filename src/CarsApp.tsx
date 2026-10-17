import React, { FormEvent, useEffect, useState } from 'react';
import { ArrowRight, Eye, ShieldCheck } from 'lucide-react';
import CarFuelPage from './pages/fleet/CarFuelPage';
import { getAdminAuthEmail, getCloudFleetUsers, readAdmins, recordLoginAttempt, saveAdminPasswordHash, saveAdmins, verifyAdminPassword, verifyGuestPassword } from './services/fleetUserStore';
import type { FleetGuest, FleetUser } from './services/fleetUserStore';
import { firebaseAuth, firebaseReady, isFirebaseConfigured } from './services/firebase';
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from 'firebase/auth';

const GUEST_SESSION_KEY = 'cars-site-guest-session';
const ADMIN_SESSION_KEY = 'cars-site-admin-session';
const LOGIN_LOCK_KEY = 'cars-site-login-attempts-v1';
const MAX_LOGIN_ATTEMPTS = 5;
const LOGIN_ATTEMPT_WINDOW = 5 * 60 * 1000;
const LOGIN_LOCK_DURATION = 5 * 60 * 1000;

type LoginAttemptState = { failedAt: number[]; lockUntil: number };

const readLoginAttemptState = (): LoginAttemptState => {
  try {
    const value = JSON.parse(localStorage.getItem(LOGIN_LOCK_KEY) || '{}') as Partial<LoginAttemptState>;
    return {
      failedAt: Array.isArray(value.failedAt)
        ? value.failedAt.filter((timestamp): timestamp is number => Number.isFinite(timestamp))
        : [],
      lockUntil: typeof value.lockUntil === 'number' && Number.isFinite(value.lockUntil) ? value.lockUntil : 0,
    };
  } catch (storageError) {
    console.error('Could not read login rate limit state', storageError);
    return { failedAt: [], lockUntil: 0 };
  }
};

const saveLoginAttemptState = (state: LoginAttemptState): void => {
  try {
    localStorage.setItem(LOGIN_LOCK_KEY, JSON.stringify(state));
  } catch (storageError) {
    console.error('Could not save login rate limit state', storageError);
  }
};

const getLoginLockRemainingSeconds = (): number => {
  const remaining = readLoginAttemptState().lockUntil - Date.now();
  return remaining > 0 ? Math.ceil(remaining / 1000) : 0;
};

const registerFailedLoginAttempt = (): number => {
  const now = Date.now();
  const state = readLoginAttemptState();
  if (state.lockUntil > now) return Math.ceil((state.lockUntil - now) / 1000);
  const failedAt = state.failedAt.filter((timestamp) => now - timestamp < LOGIN_ATTEMPT_WINDOW);
  failedAt.push(now);
  if (failedAt.length >= MAX_LOGIN_ATTEMPTS) {
    const lockUntil = now + LOGIN_LOCK_DURATION;
    saveLoginAttemptState({ failedAt: [], lockUntil });
    return Math.ceil((lockUntil - now) / 1000);
  }
  saveLoginAttemptState({ failedAt, lockUntil: 0 });
  return 0;
};

const clearLoginAttemptState = (): void => saveLoginAttemptState({ failedAt: [], lockUntil: 0 });

const logLoginAttemptSafely = (username: string, success: boolean): void => {
  void recordLoginAttempt(username, success).catch((logError: unknown) => {
    console.warn('Could not record login attempt in Firestore', logError);
  });
};

const createAdminProfile = (username: string): FleetUser => ({
  id: `admin-${username}`,
  username,
  email: getAdminAuthEmail(username),
  password: '',
  role: 'admin',
  active: true,
  createdAt: new Date().toISOString(),
});

const CarsApp: React.FC = () => {
  const [currentScreen, setCurrentScreen] = useState<'landing' | 'admin-login' | 'guest-login'>('landing');
  const [authenticated, setAuthenticated] = useState(false);
  const [authReady, setAuthReady] = useState(false);
  const [currentUser, setCurrentUser] = useState(() => {
    const storedAdmin = localStorage.getItem(ADMIN_SESSION_KEY);
    const adminUsername = storedAdmin || readAdmins()[0]?.username || 'admin';
    return createAdminProfile(adminUsername);
  });
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [guestPassword, setGuestPassword] = useState('');

  useEffect(() => {
    let disposed = false;
    let unsubscribe = () => undefined;
    const restoreAdminSession = () => {
      const storedAdmin = localStorage.getItem(ADMIN_SESSION_KEY);
      if (!storedAdmin) return false;
      const admin = readAdmins().find((candidate) => candidate.username.trim().toLowerCase() === storedAdmin.trim().toLowerCase());
      if (!admin) {
        localStorage.removeItem(ADMIN_SESSION_KEY);
        return false;
      }
      setCurrentUser(createAdminProfile(admin.username));
      setAuthenticated(true);
      setAuthReady(true);
      return true;
    };

    const restoreGuestSession = async () => {
      try {
        const guestId = localStorage.getItem(GUEST_SESSION_KEY);
        if (!guestId) {
          setAuthenticated(false);
          return;
        }
        const guests = await getCloudFleetUsers();
        if (disposed) return;
        const guest = guests.find((candidate) => candidate.id === guestId && candidate.active);
        if (!guest) {
          localStorage.removeItem(GUEST_SESSION_KEY);
          setAuthenticated(false);
          return;
        }
        setCurrentUser({
          id: guest.id,
          username: guest.employeeName,
          password: '',
          role: 'viewer',
          active: true,
          isGuest: true,
          guestName: guest.employeeName,
          createdAt: guest.createdAt,
        });
        setAuthenticated(true);
      } catch (restoreError) {
        console.error('Guest session restoration failed', restoreError);
        if (!disposed) {
          setError('تعذر استعادة جلسة الزائر. تحقق من الاتصال بالإنترنت ثم أعد تحميل الصفحة.');
          setAuthenticated(false);
        }
      }
    };

    if (restoreAdminSession()) {
      return () => { disposed = true; };
    }

    if (!firebaseAuth) {
      void restoreGuestSession().finally(() => {
        if (!disposed) setAuthReady(true);
      });
      return () => { disposed = true; };
    }

    void firebaseReady.then(() => {
      if (disposed) return;
      unsubscribe = onAuthStateChanged(firebaseAuth, (user) => {
        if (!user?.email) {
          void restoreGuestSession().finally(() => {
            if (!disposed) setAuthReady(true);
          });
          return;
        }
        const email = user.email.toLowerCase();
        const admin = readAdmins().find((candidate) => getAdminAuthEmail(candidate.username) === email);
        if (!admin) {
          setAuthenticated(false);
          setAuthReady(true);
          void signOut(firebaseAuth).catch((authError: unknown) => console.error('Failed to sign out an unrecognized account', authError));
          return;
        }
        localStorage.setItem(ADMIN_SESSION_KEY, admin.username);
        setCurrentUser(createAdminProfile(admin.username));
        setAuthenticated(true);
        setAuthReady(true);
      });
    }).catch((authError: unknown) => {
      console.error('Failed to initialize Firebase authentication', authError);
      setAuthReady(true);
    });
    return () => {
      disposed = true;
      unsubscribe();
    };
  }, []);

  const handleAdminLogin = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError('');
    setLoading(true);
    const attemptedUsername = username.trim().toLowerCase() || 'unknown';
    let loginSucceeded = false;
    try {
      const lockedForSeconds = getLoginLockRemainingSeconds();
      if (lockedForSeconds > 0) {
        setError(`تم قفل الحساب مؤقتاً. حاول بعد ${lockedForSeconds} ثانية`);
        return;
      }
      const normalizedUsername = username.trim().toLowerCase();
      const admins = readAdmins();
      const admin = admins.find((candidate) =>
        candidate.username.trim().toLowerCase() === normalizedUsername,
      );
      if (!admin || !password) throw new Error('بيانات غير صحيحة');
      saveAdmins(admins);

      const localLoginMatches = await verifyAdminPassword(admin.username, password);
      if (localLoginMatches) {
        localStorage.setItem(ADMIN_SESSION_KEY, admin.username);
        clearLoginAttemptState();
        loginSucceeded = true;
        setCurrentUser(createAdminProfile(admin.username));
        setAuthenticated(true);
        if (isFirebaseConfigured && firebaseAuth) {
          void signInWithEmailAndPassword(firebaseAuth, getAdminAuthEmail(admin.username), password)
            .then(() => {
              saveAdminPasswordHash(admin.username, password);
              localStorage.setItem(ADMIN_SESSION_KEY, admin.username);
            })
            .catch((firebaseError: unknown) => {
              console.warn('Background Firebase admin login failed; local session remains active.', firebaseError);
            });
        }
        return;
      }

      if (!isFirebaseConfigured || !firebaseAuth) {
        throw new Error('بيانات غير صحيحة');
      }

      void signInWithEmailAndPassword(firebaseAuth, getAdminAuthEmail(admin.username), password)
        .then(() => {
          saveAdminPasswordHash(admin.username, password);
          clearLoginAttemptState();
          localStorage.setItem(ADMIN_SESSION_KEY, admin.username);
          setCurrentUser(createAdminProfile(admin.username));
          setAuthenticated(true);
          loginSucceeded = true;
        })
        .catch((firebaseError: unknown) => {
          const lockedForSeconds = registerFailedLoginAttempt();
          console.error('Admin login failed', firebaseError);
          setError(lockedForSeconds > 0
            ? `تم قفل الحساب مؤقتاً. حاول بعد ${lockedForSeconds} ثانية`
            : 'بيانات غير صحيحة');
        });
      return;
    } catch (loginError) {
      const lockedForSeconds = registerFailedLoginAttempt();
      if (loginError instanceof Error && loginError.message === 'بيانات غير صحيحة') {
        setError(lockedForSeconds > 0
          ? `تم قفل الحساب مؤقتاً. حاول بعد ${lockedForSeconds} ثانية`
          : loginError.message);
      } else {
        console.error('Admin login failed', loginError);
        setError(lockedForSeconds > 0
          ? `تم قفل الحساب مؤقتاً. حاول بعد ${lockedForSeconds} ثانية`
          : loginError instanceof Error && loginError.message.includes('Firebase')
            ? loginError.message
            : 'بيانات غير صحيحة');
      }
    } finally {
      logLoginAttemptSafely(attemptedUsername, loginSucceeded);
      setLoading(false);
    }
  };

  const handleGuestLogin = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError('');
    setLoading(true);
    let attemptedUsername = 'guest';
    let loginSucceeded = false;
    try {
      const lockedForSeconds = getLoginLockRemainingSeconds();
      if (lockedForSeconds > 0) {
        setError(`تم قفل الحساب مؤقتاً. حاول بعد ${lockedForSeconds} ثانية`);
        return;
      }
      if (!guestPassword.trim()) {
        const failedAttemptsLockSeconds = registerFailedLoginAttempt();
        setError(failedAttemptsLockSeconds > 0
          ? `تم قفل الحساب مؤقتاً. حاول بعد ${failedAttemptsLockSeconds} ثانية`
          : 'ادخل كلمة المرور');
        return;
      }
      const guests = await getCloudFleetUsers();
      let guest: FleetGuest | undefined;
      for (const candidate of guests) {
        if (candidate.active && candidate.passwordHash && await verifyGuestPassword(guestPassword, candidate.passwordHash)) {
          guest = candidate;
          attemptedUsername = candidate.employeeName;
          break;
        }
      }
      if (!guest) {
        const failedAttemptsLockSeconds = registerFailedLoginAttempt();
        setError(failedAttemptsLockSeconds > 0
          ? `تم قفل الحساب مؤقتاً. حاول بعد ${failedAttemptsLockSeconds} ثانية`
          : 'كلمة المرور غير صحيحة');
        return;
      }
      try {
        localStorage.setItem(GUEST_SESSION_KEY, guest.id);
      } catch (storageError) {
        console.error('Could not persist guest session', storageError);
        setError('تعذر حفظ جلسة الزائر على هذا الجهاز');
        return;
      }
      setCurrentUser({
        id: guest.id,
        username: guest.employeeName,
        password: '',
        role: 'viewer',
        active: true,
        isGuest: true,
        guestName: guest.employeeName,
        createdAt: guest.createdAt,
      });
      setAuthenticated(true);
      loginSucceeded = true;
      clearLoginAttemptState();
      setGuestPassword('');
    } catch (guestLoginError) {
      console.error('Guest login failed', guestLoginError);
      setError('تعذر الاتصال بقائمة الزوار. حاول مرة أخرى');
    } finally {
      logLoginAttemptSafely(attemptedUsername, loginSucceeded);
      setLoading(false);
    }
  };

  const logout = () => {
    setAuthenticated(false);
    localStorage.removeItem(GUEST_SESSION_KEY);
    localStorage.removeItem(ADMIN_SESSION_KEY);
    if (currentUser.role === 'admin' && firebaseAuth) {
      void signOut(firebaseAuth).catch((authError: unknown) => console.error('Failed to sign out', authError));
    }
  };

  if (firebaseAuth && !authReady) {
    return <main className="cars-login" dir="rtl"><p className="cars-login-description">جار التحقق من جلسة الدخول...</p></main>;
  }

  if (authenticated) {
    return <CarFuelPage currentUser={currentUser} onLogout={logout} />;
  }

  const goToScreen = (screen: 'landing' | 'admin-login' | 'guest-login') => {
    setError('');
    setCurrentScreen(screen);
  };

  return (
    <main className={`cars-portal cars-portal-${currentScreen}`} dir="rtl">
      {currentScreen === 'landing' ? <section className="cars-landing-panel">
        <div className="cars-landing-background" aria-hidden="true">
          <img className="cars-landing-background-front" src="/fleet-truck-front.png" alt="" />
          <img className="cars-landing-background-rear" src="/fleet-truck-rear.png" alt="" />
        </div>
        <div className="cars-landing-copy">
          <div className="cars-landing-brand">
            <span className="cars-brand-mark" aria-hidden="true"><span className="cars-brand-ring" /></span>
            <span className="cars-brand-copy"><strong>Abu Ghali</strong><small>MURAQABAT</small></span>
          </div>
          <h1>نظام إدارة <span>السيارات</span></h1>
          <p className="cars-landing-description">متابعة سيارات الشركة وبياناته في مكان واحد.</p>
          <div className="cars-landing-actions">
            <button className="cars-portal-action cars-portal-admin-action" type="button" onClick={() => goToScreen('admin-login')}>
              <ShieldCheck aria-hidden="true" />
                <span><strong>دخول كمسؤول</strong></span>
              <ArrowRight aria-hidden="true" />
            </button>
            <button className="cars-portal-action cars-portal-guest-action" type="button" onClick={() => goToScreen('guest-login')}>
              <Eye aria-hidden="true" />
                <span><strong>دخول كزائر</strong></span>
              <ArrowRight aria-hidden="true" />
            </button>
          </div>
          <footer className="cars-landing-footer">جميع الحقوق محفوظة © شركة أبو غالي</footer>
        </div>
      </section> : <>
        <div className="cars-auth-background" aria-hidden="true">
          <img className="cars-auth-background-front" src="/fleet-truck-front.png" alt="" />
          <img className="cars-auth-background-rear" src="/fleet-truck-rear.png" alt="" />
        </div>
        <section className="cars-auth-panel">
        <button className="cars-portal-back" type="button" onClick={() => goToScreen('landing')}>
          <ArrowRight aria-hidden="true" /> رجوع
        </button>
        <div className="cars-auth-brand">
          <span className="cars-brand-mark cars-brand-mark-auth" aria-hidden="true"><span className="cars-brand-ring" /></span>
          <span className="cars-brand-copy"><strong>Abu Ghali</strong><small>MURAQABAT</small></span>
        </div>
        {currentScreen === 'admin-login' ? <>
          <div className="cars-auth-icon cars-auth-admin-icon"><ShieldCheck aria-hidden="true" /></div>
          <h1>دخول المسؤول</h1>
          <p className="cars-auth-description">أدخل بيانات المسؤول للوصول إلى لوحة التحكم.</p>
          <form className="cars-auth-form" onSubmit={handleAdminLogin}>
            <label htmlFor="admin-username">اسم المستخدم</label>
            <input id="admin-username" type="text" value={username} onChange={(event) => setUsername(event.target.value)} placeholder="اسم المستخدم" autoComplete="username" required disabled={loading} />
            <label htmlFor="admin-password">كلمة المرور</label>
            <input id="admin-password" type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete="current-password" required disabled={loading} />
            {error && <p className="cars-login-error" role="alert">{error}</p>}
            <button className="cars-auth-submit cars-auth-submit-admin" type="submit" disabled={loading}>{loading ? 'جارٍ تسجيل الدخول...' : 'دخول'}</button>
          </form>
        </> : <>
          <div className="cars-auth-icon cars-auth-guest-icon"><Eye aria-hidden="true" /></div>
          <h1>دخول الزائر</h1>
          <p className="cars-auth-description">اكتب كلمة المرور التي حصلت عليها من المسؤول.</p>
          <form className="cars-auth-form" onSubmit={(event) => void handleGuestLogin(event)}>
            <label htmlFor="guest-password">كلمة المرور</label>
            <input id="guest-password" type="password" value={guestPassword} onChange={(event) => setGuestPassword(event.target.value)} placeholder="ادخل كلمة المرور" autoComplete="current-password" required disabled={loading} />
            {error && <p className="cars-login-error" role="alert">{error}</p>}
            <button className="cars-auth-submit cars-auth-submit-guest" type="submit" disabled={loading}>{loading ? 'جارٍ التحقق...' : 'دخول'}</button>
          </form>
        </>}
        </section>
      </>}
    </main>
  );
};

export default CarsApp;
