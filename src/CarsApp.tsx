import React, { FormEvent, useEffect, useState } from 'react';
import bcrypt from 'bcryptjs';
import { ArrowRight, Eye, ShieldCheck } from 'lucide-react';
import CarFuelPage from './pages/fleet/CarFuelPage';
import { getAdminAuthEmail, getCloudFleetUsers, readAdmins, saveAdmins } from './services/fleetUserStore';
import type { FleetGuest, FleetUser } from './services/fleetUserStore';
import { firebaseAuth, firebaseReady, isFirebaseConfigured } from './services/firebase';
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from 'firebase/auth';

const GUEST_SESSION_KEY = 'cars-site-guest-session';

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
  const [currentUser, setCurrentUser] = useState(() => createAdminProfile(readAdmins()[0].username));
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [guestPassword, setGuestPassword] = useState('');

  useEffect(() => {
    let disposed = false;
    let unsubscribe = () => undefined;
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
    try {
      const normalizedUsername = username.trim().toLowerCase();
      const admins = readAdmins();
      const admin = admins.find((candidate) =>
        candidate.username.trim().toLowerCase() === normalizedUsername,
      );
      if (!admin || !password) throw new Error('بيانات غير صحيحة');
      if (!isFirebaseConfigured || !firebaseAuth) {
        throw new Error('تعذر الاتصال بخدمة تسجيل دخول المسؤول');
      }
      saveAdmins(admins);
      await firebaseReady;
      await signInWithEmailAndPassword(firebaseAuth, getAdminAuthEmail(admin.username), password);
      setCurrentUser(createAdminProfile(admin.username));
      setAuthenticated(true);
    } catch (loginError) {
      if (loginError instanceof Error && loginError.message === 'بيانات غير صحيحة') {
        setError(loginError.message);
      } else {
        console.error('Admin login failed', loginError);
        setError(loginError instanceof Error && loginError.message.includes('Firebase')
          ? loginError.message
          : 'بيانات غير صحيحة');
      }
    } finally {
      setLoading(false);
    }
  };

  const handleGuestLogin = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError('');
    setLoading(true);
    try {
      if (!guestPassword.trim()) {
        setError('ادخل كلمة المرور');
        return;
      }
      const guests = await getCloudFleetUsers();
      let guest: FleetGuest | undefined;
      for (const candidate of guests) {
        if (candidate.active && candidate.passwordHash && await bcrypt.compare(guestPassword, candidate.passwordHash)) {
          guest = candidate;
          break;
        }
      }
      if (!guest) {
        setError('كلمة المرور غير صحيحة');
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
      setGuestPassword('');
    } catch (guestLoginError) {
      console.error('Guest login failed', guestLoginError);
      setError('تعذر الاتصال بقائمة الزوار. حاول مرة أخرى');
    } finally {
      setLoading(false);
    }
  };

  const logout = () => {
    setAuthenticated(false);
    localStorage.removeItem(GUEST_SESSION_KEY);
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
