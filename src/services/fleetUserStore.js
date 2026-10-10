const USERS_KEY = 'cars-site-users-v1';
const defaultUsers = [{ id: 'u-admin', username: 'admin', password: 'admin123', role: 'admin', active: true, createdAt: new Date().toISOString() }];
export const readFleetUsers = () => {
    try {
        const stored = localStorage.getItem(USERS_KEY);
        if (!stored) {
            localStorage.setItem(USERS_KEY, JSON.stringify(defaultUsers));
            return [...defaultUsers];
        }
        const users = JSON.parse(stored);
        return users.length ? users : [...defaultUsers];
    }
    catch {
        return [...defaultUsers];
    }
};
export const saveFleetUsers = (users) => {
    localStorage.setItem(USERS_KEY, JSON.stringify(users));
};
