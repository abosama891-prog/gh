import { jsx as _jsx, jsxs as _jsxs } from "react/jsx-runtime";
import React, { useState } from 'react';
import CarFuelPage from './pages/fleet/CarFuelPage';
import { readFleetUsers } from './services/fleetUserStore';
const SESSION_KEY = 'cars-site-authenticated';
const USERNAME_KEY = 'cars-site-username';
const CarsApp = () => {
    const [authenticated, setAuthenticated] = useState(() => localStorage.getItem(SESSION_KEY) === 'true');
    const [currentUser, setCurrentUser] = useState(() => {
        const username = localStorage.getItem(USERNAME_KEY) || 'admin';
        return readFleetUsers().find((user) => user.username === username) || readFleetUsers()[0];
    });
    const [username, setUsername] = useState('admin');
    const [password, setPassword] = useState('');
    const [error, setError] = useState('');
    const handleSubmit = (event) => {
        event.preventDefault();
        const account = readFleetUsers().find((user) => user.username.trim().toLowerCase() === username.trim().toLowerCase() && user.password === password && user.active);
        if (account) {
            localStorage.setItem(SESSION_KEY, 'true');
            localStorage.setItem(USERNAME_KEY, account.username);
            setCurrentUser(account);
            setAuthenticated(true);
            setError('');
            return;
        }
        setError('اسم المستخدم أو كلمة المرور غير صحيحة');
    };
    if (authenticated) {
        return _jsx(CarFuelPage, { currentUser: currentUser, onLogout: () => { localStorage.removeItem(SESSION_KEY); localStorage.removeItem(USERNAME_KEY); setAuthenticated(false); } });
    }
    return (_jsx("main", { className: "cars-login", dir: "rtl", children: _jsxs("form", { className: "cars-login-form", onSubmit: handleSubmit, children: [_jsxs("div", { className: "cars-login-brand", children: [_jsx("img", { src: "/car-mark.svg", alt: "" }), _jsxs("div", { children: [_jsx("p", { children: "\u0645\u0631\u0627\u0642\u0628\u0629 \u0633\u064A\u0627\u0631\u0627\u062A \u0623\u0628\u0648 \u063A\u0627\u0644\u064A" }), _jsx("h1", { children: "\u062A\u0633\u062C\u064A\u0644 \u0627\u0644\u062F\u062E\u0648\u0644" })] })] }), _jsx("p", { className: "cars-login-description", children: "\u0623\u062F\u062E\u0644 \u0628\u064A\u0627\u0646\u0627\u062A \u0627\u0644\u0645\u0633\u062A\u062E\u062F\u0645 \u0644\u0644\u0648\u0635\u0648\u0644 \u0625\u0644\u0649 \u0644\u0648\u062D\u0629 \u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A." }), _jsxs("label", { children: ["\u0627\u0633\u0645 \u0627\u0644\u0645\u0633\u062A\u062E\u062F\u0645", _jsx("input", { value: username, onChange: (event) => setUsername(event.target.value), autoComplete: "username", required: true })] }), _jsxs("label", { children: ["\u0643\u0644\u0645\u0629 \u0627\u0644\u0645\u0631\u0648\u0631", _jsx("input", { type: "password", value: password, onChange: (event) => setPassword(event.target.value), autoComplete: "current-password", required: true })] }), error && _jsx("p", { className: "cars-login-error", role: "alert", children: error }), _jsx("button", { className: "cars-login-submit", type: "submit", children: "\u062F\u062E\u0648\u0644" }), _jsx("small", { className: "cars-login-hint", children: "\u0628\u064A\u0627\u0646\u0627\u062A \u0627\u0644\u062A\u062C\u0631\u0628\u0629: admin / admin123" })] }) }));
};
export default CarsApp;
