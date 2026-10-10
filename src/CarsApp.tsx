import React, { FormEvent, useEffect, useState } from 'react';
import CarFuelPage from './pages/fleet/CarFuelPage';
import { readFleetUsers } from './services/fleetUserStore';
import { firebaseAuth, firebaseReady } from './services/firebase';
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from 'firebase/auth';

const SESSION_KEY = 'cars-site-authenticated';

const CarsApp: React.FC = () => {
  const [authenticated, setAuthenticated] = useState(false);
  const [authReady, setAuthReady] = useState(!firebaseAuth);
  const [currentUser, setCurrentUser] = useState(() => readFleetUsers()[0]);
  const [email, setEmail] = useState('abosama891@gmail.com');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

  useEffect(() => {
    const captureInstall = (event: Event) => {
      event.preventDefault();
      window.deferredInstallPrompt = event as BeforeInstallPromptEvent;
    };
    window.addEventListener('beforeinstallprompt', captureInstall);
    return () => window.removeEventListener('beforeinstallprompt', captureInstall);
  }, []);

  useEffect(() => {
    if (!firebaseAuth) return undefined;
    let unsubscribe = () => undefined;
    void firebaseReady.then(() => {
      unsubscribe = onAuthStateChanged(firebaseAuth, (user) => {
        setAuthenticated(Boolean(user));
        setAuthReady(true);
      });
    });
    return () => unsubscribe();
  }, []);

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (firebaseAuth) {
      try {
        await firebaseReady;
        await signInWithEmailAndPassword(firebaseAuth, email.trim(), password);
        setCurrentUser(readFleetUsers()[0]);
        setError('');
        return;
      } catch {
        setError('البريد الإلكتروني أو كلمة المرور غير صحيحة');
        return;
      }
    }
    const account = readFleetUsers().find((user) => user.username === 'admin' && user.password === password && user.active);
    if (account) {
      localStorage.setItem(SESSION_KEY, 'true');
      setCurrentUser(account);
      setAuthenticated(true);
      setError('');
      return;
    }
    setError('اسم المستخدم أو كلمة المرور غير صحيحة');
  };

  if (firebaseAuth && !authReady) {
    return <main className="cars-login" dir="rtl"><p className="cars-login-description">جار التحقق من جلسة الدخول...</p></main>;
  }

  if (authenticated) {
    return <CarFuelPage currentUser={currentUser} onLogout={() => { if (firebaseAuth) void signOut(firebaseAuth); localStorage.removeItem(SESSION_KEY); setAuthenticated(false); }} />;
  }

  return (
    <main className="cars-login" dir="rtl">
      <form className="cars-login-form" onSubmit={handleSubmit}>
        <div className="cars-login-brand"><img src="/car-mark.svg" alt="" /><div><p>مراقبة سيارات أبو غالي</p><h1>تسجيل الدخول</h1></div></div>
        <p className="cars-login-description">أدخل بيانات المستخدم للوصول إلى لوحة السيارات.</p>
        <label>البريد الإلكتروني<input type="email" value={email} onChange={(event) => setEmail(event.target.value)} autoComplete="email" required /></label>
        <label>كلمة المرور<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete="current-password" required /></label>
        {error && <p className="cars-login-error" role="alert">{error}</p>}
        <button className="cars-login-submit" type="submit">دخول</button>
      </form>
    </main>
  );
};

export default CarsApp;