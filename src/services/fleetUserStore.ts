export type FleetRole = 'admin' | 'manager' | 'viewer' | 'accountant' | 'warehouse';

export type FleetUser = {
  id: string;
  username: string;
  email?: string;
  employeeName?: string;
  password: string;
  role: FleetRole;
  active: boolean;
  createdAt: string;
};

const USERS_KEY = 'cars-site-users-v1';
const defaultUsers: FleetUser[] = [{ id: 'u-admin', username: 'admin', email: 'abosama891@gmail.com', password: 'admin123', role: 'admin', active: true, createdAt: new Date().toISOString() }];

export const readFleetUsers = (): FleetUser[] => {
  try {
    const stored = localStorage.getItem(USERS_KEY);
    if (!stored) {
      localStorage.setItem(USERS_KEY, JSON.stringify(defaultUsers));
      return [...defaultUsers];
    }
    const users = JSON.parse(stored) as FleetUser[];
    return users.length ? users : [...defaultUsers];
  } catch {
    return [...defaultUsers];
  }
};

export const saveFleetUsers = (users: FleetUser[]) => {
  localStorage.setItem(USERS_KEY, JSON.stringify(users));
};