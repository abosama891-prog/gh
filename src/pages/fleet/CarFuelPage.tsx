import React, { useEffect, useMemo, useRef, useState } from 'react';
import type ExcelJS from 'exceljs';
import { ArrowRight, CarFront, Download, Fuel, LayoutDashboard, RefreshCw, UsersRound } from 'lucide-react';
import { readFleetUsers, saveFleetUsers } from '../../services/fleetUserStore';
import type { FleetRole, FleetUser } from '../../services/fleetUserStore';
import { isFirebaseConfigured } from '../../services/firebase';
import { readCloudFleetData, saveCloudFleetData } from '../../services/fleetCloudStore';
import { firebaseAuth } from '../../services/firebase';
import { sendPasswordResetEmail, updatePassword } from 'firebase/auth';
import { applyUpdate } from '../../services/serviceWorker';
import './car-fuel.css';

type Vehicle = {
  id: string;
  name: string;
  plate: string;
  driver: string;
  openingOdometer: number;
  kmPerLiter: number;
  fuelLitersPer105Km?: number;
  fuelDistanceBaseKm?: number;
};

type FuelEntry = {
  id: string;
  vehicleId: string;
  date: string;
  startOdometer: number;
  currentOdometer: number;
  liters: number;
  pricePerLiter: number;
  fuelActualValue?: number;
  fuelActualExpression?: string;
  driver?: string;
  note: string;
  route?: string;
  allowanceValue?: number;
  expenses?: Partial<ExpenseBreakdown>;
  expenseExpressions?: Partial<Record<keyof ExpenseBreakdown, string>>;
};

type ExpenseBreakdown = {
  mealAllowance: number;
  tolls: number;
  traffic: number;
  violations: number;
  maintenance: number;
  extra: number;
  carWash: number;
  tireExpenses: number;
  oilChange: number;
  transport: number;
  other: number;
};

type ExpenseInputs = Record<keyof ExpenseBreakdown, string>;

type EntryForm = {
  vehicleId: string;
  driver: string;
  startOdometer: string;
  currentOdometer: string;
  liters: string;
  pricePerLiter: string;
  fuelActualValue: string;
  allowanceValue: string;
  date: string;
  route: string;
  note: string;
  expenses: ExpenseInputs;
};

type FleetData = { vehicles: Vehicle[]; entries: FuelEntry[]; allowancePricePerLiter?: number };
type InstallPrompt = Event & { prompt: () => Promise<void>; userChoice: Promise<{ outcome: 'accepted' | 'dismissed' }> };

const STORAGE_KEY = 'company-fleet-fuel-v1';
const DEFAULT_ALLOWANCE_PRICE = 15.75;
const themePresets = {
  emerald: { label: 'أخضر مؤسسي', accent: '#087f72', heading: '#075e59', ink: '#183a37', line: '#cbded8', surface: '#ffffff', soft: '#edf7f4', background: '#f1f6f4' },
  ocean: { label: 'أزرق رسمي', accent: '#1465a0', heading: '#123e68', ink: '#18324b', line: '#c8d9e8', surface: '#ffffff', soft: '#edf5fc', background: '#f0f5f9' },
  ruby: { label: 'أحمر وذهبي', accent: '#b22d43', heading: '#7e2033', ink: '#402630', line: '#e7cbd0', surface: '#ffffff', soft: '#fff0f2', background: '#fbf5f5' },
  gold: { label: 'ذهبي دافئ', accent: '#a56b08', heading: '#714b08', ink: '#3f351f', line: '#e4d4aa', surface: '#fffefa', soft: '#fff8e7', background: '#faf7ee' },
  slate: { label: 'رمادي حديث', accent: '#3c6688', heading: '#28445d', ink: '#263744', line: '#cbd8e2', surface: '#ffffff', soft: '#eef4f8', background: '#f1f4f7' },
  turquoise: { label: 'تركوازي ناصع', accent: '#008c95', heading: '#005c68', ink: '#18383d', line: '#bfe0e2', surface: '#ffffff', soft: '#e9f8f8', background: '#eef8f8' },
  indigo: { label: 'نيلي ملكي', accent: '#4d46a8', heading: '#302878', ink: '#292744', line: '#d2d0ed', surface: '#ffffff', soft: '#f0effc', background: '#f5f4fc' },
  orange: { label: 'برتقالي حيوي', accent: '#d45a12', heading: '#94370b', ink: '#442a1d', line: '#f0d1bd', surface: '#ffffff', soft: '#fff3e9', background: '#fff8f2' },
  violet: { label: 'بنفسجي عصري', accent: '#7938a8', heading: '#512477', ink: '#35233e', line: '#dfcceb', surface: '#ffffff', soft: '#f7effc', background: '#faf6fc' },
  dark: { label: 'داكن', accent: '#64748b', heading: '#e2e8f0', ink: '#e5e7eb', line: '#475569', surface: '#1e293b', soft: '#293548', background: '#0f172a' },
  charcoal: { label: 'فحمي وليموني', accent: '#42515b', heading: '#26343d', ink: '#222b30', line: '#ccd6d8', surface: '#ffffff', soft: '#f0f5f2', background: '#eef2f0' },
  coral: { label: 'مرجاني مشرق', accent: '#d94f62', heading: '#9c2941', ink: '#422932', line: '#efcbd1', surface: '#ffffff', soft: '#fff0f2', background: '#fff7f7' },
} as const;
type ThemeKey = keyof typeof themePresets;
const fontSizePresets = { small: 'صغير', medium: 'متوسط', large: 'كبير' } as const;
type FontSizeKey = keyof typeof fontSizePresets;
const today = () => new Date().toISOString().slice(0, 10);
const vehicleToneStyle = (vehicleId: string, vehicles: Vehicle[]): React.CSSProperties => {
  const index = vehicles.findIndex((vehicle) => vehicle.id === vehicleId);
  const palette = [
    ['#19776c', '#edf6f3', '#cfe4dc'],
    ['#426f91', '#eef4f8', '#d3e0e9'],
    ['#a66e2d', '#faf4e9', '#e8dbc2'],
    ['#a85851', '#faf0ef', '#ead7d4'],
    ['#6d7d45', '#f3f5ec', '#dfe4d0'],
  ];
  const [accent, tint, line] = index < palette.length
    ? palette[Math.max(0, index)]
    : (() => {
      const hue = (index * 137.508 + 29) % 360;
      return [`hsl(${hue} 36% 38%)`, `hsl(${hue} 32% 96%)`, `hsl(${hue} 25% 84%)`];
    })();
  return { '--vehicle-accent': accent, '--vehicle-tint': tint, '--vehicle-line': line } as React.CSSProperties;
};

const initialData: FleetData = {
  vehicles: [
    { id: 'v1', name: 'تويوتا هايلوكس', plate: 'ر س د 4821', driver: 'أحمد العتيبي', openingOdometer: 42800, kmPerLiter: 8.5, fuelLitersPer105Km: 20 },
    { id: 'v2', name: 'نيسان باترول', plate: 'ر س د 1936', driver: 'خالد الحربي', openingOdometer: 71650, kmPerLiter: 6.8, fuelLitersPer105Km: 20 },
    { id: 'v3', name: 'فورد ترانزيت', plate: 'ر س د 7054', driver: 'سالم القحطاني', openingOdometer: 25400, kmPerLiter: 9.2, fuelLitersPer105Km: 20 },
  ],
  entries: [
    { id: 'e1', vehicleId: 'v1', date: '2026-09-24', startOdometer: 42800, currentOdometer: 43120, liters: 43, pricePerLiter: 2.18, note: 'محطة طريق المطار' },
    { id: 'e2', vehicleId: 'v2', date: '2026-09-25', startOdometer: 71650, currentOdometer: 72058, liters: 68, pricePerLiter: 2.18, note: 'مهمة ميدانية' },
    { id: 'e3', vehicleId: 'v3', date: '2026-09-26', startOdometer: 25400, currentOdometer: 25720, liters: 32, pricePerLiter: 2.18, note: '' },
  ],
};

const readData = (): FleetData => {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    return stored ? JSON.parse(stored) as FleetData : initialData;
  } catch {
    return initialData;
  }
};

const numberFormat = new Intl.NumberFormat('en-EG', { maximumFractionDigits: 1 });
const currencyFormat = new Intl.NumberFormat('en-EG', { style: 'currency', currency: 'EGP', maximumFractionDigits: 0 });
const rateFormat = new Intl.NumberFormat('en-EG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const dateFormat = new Intl.DateTimeFormat('ar-EG-u-nu-latn', { year: 'numeric', month: '2-digit', day: '2-digit' });
const formatDateArabic = (value: string) => dateFormat.format(new Date(`${value}T00:00:00`));
const fuelAllowanceRate = (data: FleetData) => data.allowancePricePerLiter ?? DEFAULT_ALLOWANCE_PRICE;
const vehicleFuelRate = (vehicle: Vehicle) => vehicle.fuelLitersPer105Km ?? 20;
const vehicleDistanceBase = (vehicle: Vehicle) => vehicle.fuelDistanceBaseKm ?? 105;
const litersForDistance = (vehicle: Vehicle, distance: number) => Math.max(0, distance) / vehicleDistanceBase(vehicle) * vehicleFuelRate(vehicle);
const litersForEntry = (entry: FuelEntry, vehicle?: Vehicle) => vehicle ? litersForDistance(vehicle, entry.currentOdometer - entry.startOdometer) : entry.liters;
const expenseFields: Array<{ key: keyof ExpenseBreakdown; label: string }> = [
  { key: 'mealAllowance', label: 'بدل وجبات' },
  { key: 'extra', label: 'إضافي' },
  { key: 'maintenance', label: 'صيانة' },
  { key: 'tolls', label: 'كارتة' },
  { key: 'traffic', label: 'مرور' },
  { key: 'violations', label: 'مخالفات' },
  { key: 'carWash', label: 'غسيل سيارة' },
  { key: 'tireExpenses', label: 'مصاريف كاوتش' },
  { key: 'oilChange', label: 'غيار زيت' },
  { key: 'transport', label: 'نقل وانتقال' },
  { key: 'other', label: 'أخرى' },
];
const emptyExpenses = (): ExpenseBreakdown => ({ mealAllowance: 0, tolls: 0, traffic: 0, violations: 0, maintenance: 0, extra: 0, carWash: 0, tireExpenses: 0, oilChange: 0, transport: 0, other: 0 });
const emptyExpenseInputs = (): ExpenseInputs => Object.fromEntries(expenseFields.map(({ key }) => [key, ''])) as ExpenseInputs;
const sumExpenses = (expenses?: Partial<ExpenseBreakdown>) => Math.round(expenseFields.reduce((sum, field) => sum + (Number(expenses?.[field.key]) || 0), 0));
const entryFuelCost = (entry: FuelEntry) => Math.round(entry.fuelActualValue ?? entry.liters * entry.pricePerLiter);
const allowanceCost = (liters: number, rate: number) => Math.round(Math.max(0, Number(liters) || 0) * rate);
const entryDeviation = (entry: FuelEntry, rate: number, liters = entry.liters) => entryFuelCost(entry) - allowanceCost(liters, rate);
const entryTotalCost = (entry: FuelEntry) => Math.round(entryFuelCost(entry) + sumExpenses(entry.expenses));

const parseCsvRows = (content: string): string[][] => {
  const text = content.replace(/^\uFEFF/, '');
  const firstLine = text.split(/\r?\n/, 1)[0] || '';
  let quoted = false;
  let commas = 0;
  let semicolons = 0;
  for (let index = 0; index < firstLine.length; index += 1) {
    if (firstLine[index] === '"') {
      if (quoted && firstLine[index + 1] === '"') index += 1;
      else quoted = !quoted;
    } else if (!quoted && firstLine[index] === ',') commas += 1;
    else if (!quoted && firstLine[index] === ';') semicolons += 1;
  }
  const delimiter = semicolons > commas ? ';' : ',';
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') quoted = false;
      else field += character;
    } else if (character === '"' && field.length === 0) quoted = true;
    else if (character === delimiter) {
      row.push(field);
      field = '';
    } else if (character === '\n' || character === '\r') {
      row.push(field);
      if (row.some((value) => value.trim() !== '')) rows.push(row);
      row = [];
      field = '';
      if (character === '\r' && text[index + 1] === '\n') index += 1;
    } else field += character;
  }

  if (quoted) throw new Error('يوجد حقل نصي بعلامة اقتباس غير مكتملة في الملف.');
  row.push(field);
  if (row.some((value) => value.trim() !== '')) rows.push(row);
  return rows;
};

const parseTemplateNumber = (value: string) => {
  const westernDigits = value
    .replace(/[٠-٩]/g, (digit) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit)))
    .replace(/[٬,،\s]/g, '')
    .replace(/٫/g, '.');
  return Number(westernDigits);
};

const parseAmountExpression = (value: string): number | null => {
  const normalized = value
    .replace(/[٠-٩]/g, (digit) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit)))
    .replace(/٫/g, '.')
    .replace(/[٬,،\s]/g, '');
  if (!normalized) return 0;
  const amounts = normalized.split('+');
  if (amounts.some((amount) => !/^(?:\d+(?:\.\d*)?|\.\d+)$/.test(amount))) return null;
  const total = amounts.reduce((sum, amount) => sum + Number(amount), 0);
  return Number.isFinite(total) ? Math.round(total) : null;
};
const parseSingleAmount = (value: string) => value.includes('+') ? null : parseAmountExpression(value);
const splitAmountExpression = (value: string) => value.split('+').map((amount) => amount.trim()).filter(Boolean);

const excelCellToString = (value: ExcelJS.CellValue): string => {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'object') {
    if ('result' in value) return excelCellToString(value.result as ExcelJS.CellValue);
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
    if ('text' in value) return String(value.text ?? '');
    return '';
  }
  return String(value);
};

const worksheetToRows = (worksheet: ExcelJS.Worksheet) => {
  const rows: string[][] = [];
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    rows.push(Array.from({ length: worksheet.columnCount }, (_, index) => excelCellToString(row.getCell(index + 1).value)));
  });
  return rows;
};

const fleetSections = [
  { key: 'overview', label: 'نظرة عامة', description: 'ملخص الأداء والحركات', Icon: LayoutDashboard },
  { key: 'vehicles', label: 'السيارات', description: 'بيانات ومصروفات المركبات', Icon: CarFront },
  { key: 'entries', label: 'سجل التعبئة', description: 'حركات الوقود المسجلة', Icon: Fuel },
] as const;
const fleetHomeActions = [
  { key: 'users', label: 'إدارة المستخدمين والصلاحيات', description: 'إدارة الحسابات والأدوار', Icon: UsersRound },
  { key: 'update', label: 'تحديث النظام', description: 'التحقق من أحدث نسخة', Icon: RefreshCw },
  { key: 'install', label: 'تثبيت التطبيق', description: 'إضافة التطبيق إلى جهازك', Icon: Download },
] as const;
const fleetRoleLabels: Record<FleetRole, string> = { admin: 'مدير النظام', manager: 'مدير فرع', viewer: 'مشاهد', accountant: 'محاسب', warehouse: 'مسؤول مخزن' };

type CarFuelPageProps = { onLogout?: () => void; currentUser?: FleetUser };
type LedgerPeriod = 'all' | 'month' | 'custom';
type FleetRoute = { tab: 'home' | 'overview' | 'vehicles' | 'entries' | 'users'; selectedVehicleId: string | null };
type FleetHistoryState = { fleetRoute?: FleetRoute };

const dateInputValue = (date: Date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
const monthNames = Array.from({ length: 12 }, (_, index) => new Intl.DateTimeFormat('ar-EG', { month: 'long' }).format(new Date(2020, index, 1)));
const CarFuelPage: React.FC<CarFuelPageProps> = ({ onLogout, currentUser = readFleetUsers()[0] }) => {
  const [data, setData] = useState<FleetData>(readData);
  const [cloudLoaded, setCloudLoaded] = useState(!isFirebaseConfigured);
  const [tab, setTab] = useState<'home' | 'overview' | 'vehicles' | 'entries' | 'users'>('home');
  const routeHistoryRef = useRef<FleetRoute | null>(null);
  const skipRoutePushRef = useRef(false);
  const [users, setUsers] = useState<FleetUser[]>(readFleetUsers);
  const [newUser, setNewUser] = useState({ username: '', employeeName: '', password: '', role: 'viewer' as FleetRole });
  const [editingUserId, setEditingUserId] = useState<string | null>(null);
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const themeStorageKey = `cars-site-theme-${currentUser.email || currentUser.username}`;
  const [themeKey, setThemeKey] = useState<ThemeKey>(() => (localStorage.getItem(themeStorageKey) as ThemeKey) || 'ruby');
  const [fontSizeKey, setFontSizeKey] = useState<FontSizeKey>(() => (localStorage.getItem(`${themeStorageKey}-font`) as FontSizeKey) || 'medium');
  const [openingOdometerDrafts, setOpeningOdometerDrafts] = useState<Record<string, string>>({});
  const [isThemePickerOpen, setIsThemePickerOpen] = useState(false);
  const theme = themePresets[themeKey];
  const [selectedVehicleId, setSelectedVehicleId] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [ledgerFilters, setLedgerFilters] = useState({ vehicleId: 'all', from: '', to: '', search: '' });
  const [ledgerPeriod, setLedgerPeriod] = useState<LedgerPeriod>('all');
  const [ledgerMonthYear, setLedgerMonthYear] = useState('');
  const [isLedgerMonthPickerOpen, setIsLedgerMonthPickerOpen] = useState(false);
  const [ledgerPickerYear, setLedgerPickerYear] = useState(new Date().getFullYear());
  const [vehicleDialog, setVehicleDialog] = useState(false);
  const [editingVehicleId, setEditingVehicleId] = useState<string | null>(null);
  const vehicleImportRef = useRef<HTMLInputElement>(null);
  const [isImportingVehicles, setIsImportingVehicles] = useState(false);
  const vehicleHistoryImportRef = useRef<HTMLInputElement>(null);
  const [isImportingVehicleHistory, setIsImportingVehicleHistory] = useState(false);
  const [entryDialog, setEntryDialog] = useState(false);
  const [editingEntryId, setEditingEntryId] = useState<string | null>(null);
  const [allowanceRateDraft, setAllowanceRateDraft] = useState(String(DEFAULT_ALLOWANCE_PRICE));
  const [installPrompt, setInstallPrompt] = useState<InstallPrompt | null>(null);
  const [notice, setNotice] = useState('');
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [pendingDeleteEntry, setPendingDeleteEntry] = useState<FuelEntry | null>(null);
  const [deleteConfirmationText, setDeleteConfirmationText] = useState('');
  const [isRateEditorOpen, setIsRateEditorOpen] = useState(false);
  const [vehicleForm, setVehicleForm] = useState({ name: '', plate: '', driver: '', openingOdometer: '', kmPerLiter: '20', distanceBaseKm: '105' });
  const [entryForm, setEntryForm] = useState<EntryForm>({ vehicleId: '', driver: '', startOdometer: '', currentOdometer: '', liters: '', pricePerLiter: '2.18', fuelActualValue: '', allowanceValue: '', date: today(), route: '', note: '', expenses: emptyExpenseInputs() });
  const [entryPlateSearch, setEntryPlateSearch] = useState('');
  const [fuelAmountDraft, setFuelAmountDraft] = useState('');
  const [expenseAmountDrafts, setExpenseAmountDrafts] = useState<ExpenseInputs>(emptyExpenseInputs);
  const canManageUsers = currentUser.role === 'admin';
  const canEditFleet = currentUser.role !== 'viewer';

  useEffect(() => {
    const handleUpdateAvailable = () => setUpdateAvailable(true);
    window.addEventListener('sw-update-available', handleUpdateAvailable);
    return () => window.removeEventListener('sw-update-available', handleUpdateAvailable);
  }, []);

  useEffect(() => {
    const currentRoute: FleetRoute = { tab, selectedVehicleId };
    const historyState = window.history.state as FleetHistoryState | null;
    if (!routeHistoryRef.current) {
      window.history.replaceState({ ...historyState, fleetRoute: currentRoute }, '');
      routeHistoryRef.current = currentRoute;
      return;
    }

    const previousRoute = routeHistoryRef.current;
    routeHistoryRef.current = currentRoute;
    if (!previousRoute || (previousRoute.tab === tab && previousRoute.selectedVehicleId === selectedVehicleId)) return;
    if (skipRoutePushRef.current) {
      skipRoutePushRef.current = false;
      return;
    }
    window.history.pushState({ fleetRoute: currentRoute }, '');
  }, [selectedVehicleId, tab]);

  useEffect(() => {
    const handlePopState = (event: PopStateEvent) => {
      const historyState = event.state as FleetHistoryState | null;
      const previousRoute = historyState?.fleetRoute;
      if (!previousRoute) return;

      skipRoutePushRef.current = true;
      setTab(previousRoute.tab);
      setSelectedVehicleId(previousRoute.selectedVehicleId);
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []);

  useEffect(() => {
    const storedTheme = localStorage.getItem(themeStorageKey) as ThemeKey | null;
    setThemeKey(storedTheme && storedTheme in themePresets ? storedTheme : 'ruby');
    const storedFontSize = localStorage.getItem(`${themeStorageKey}-font`) as FontSizeKey | null;
    setFontSizeKey(storedFontSize && storedFontSize in fontSizePresets ? storedFontSize : 'medium');
  }, [themeStorageKey]);

  const chooseTheme = (nextTheme: ThemeKey) => {
    setThemeKey(nextTheme);
    localStorage.setItem(themeStorageKey, nextTheme);
  };

  const chooseFontSize = (nextFontSize: FontSizeKey) => {
    setFontSizeKey(nextFontSize);
    localStorage.setItem(`${themeStorageKey}-font`, nextFontSize);
  };

  const saveOpeningOdometer = (vehicle: Vehicle, entriesCount: number) => {
    const value = Number(openingOdometerDrafts[vehicle.id] ?? vehicle.openingOdometer);
    if (!Number.isFinite(value) || value < 0) {
      setNotice('أدخل قراءة سابقة صحيحة');
      return;
    }
    setData((current) => ({ ...current, vehicles: current.vehicles.map((item) => item.id === vehicle.id ? { ...item, openingOdometer: value } : item) }));
    setOpeningOdometerDrafts((current) => ({ ...current, [vehicle.id]: String(value) }));
    setNotice(entriesCount ? `تم حفظ القراءة السابقة لسيارة ${vehicle.name} وتطبيقها كقراءة بداية لاحقة` : `تم حفظ القراءة السابقة وتطبيقها على سيارة ${vehicle.name}`);
  };

  const persistUsers = (nextUsers: FleetUser[]) => {
    saveFleetUsers(nextUsers);
    setUsers(nextUsers);
  };

  const addUser = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const username = newUser.username.trim();
    const editingExistingUser = Boolean(editingUserId && editingUserId !== 'new');
    if (!canManageUsers || !username || (!editingExistingUser && !newUser.password.trim())) {
      setNotice('أدخل اسم المستخدم وكلمة المرور');
      return;
    }
    if (users.some((user) => user.id !== editingUserId && user.username.trim().toLowerCase() === username.toLowerCase())) {
      setNotice('اسم المستخدم موجود بالفعل');
      return;
    }
    if (editingExistingUser) {
      persistUsers(users.map((user) => user.id === editingUserId ? { ...user, username, employeeName: newUser.employeeName, role: newUser.role, ...(newUser.password.trim() ? { password: newUser.password } : {}) } : user));
      setEditingUserId(null);
      setNotice('تم تعديل بيانات المستخدم');
    } else {
      persistUsers([...users, { id: `u-${Date.now()}`, username, employeeName: newUser.employeeName, password: newUser.password, role: newUser.role, active: true, createdAt: new Date().toISOString() }]);
      setNotice('تمت إضافة المستخدم');
    }
    setNewUser({ username: '', employeeName: '', password: '', role: 'viewer' });
  };

  const editUser = (user: FleetUser) => {
    setEditingUserId(user.id);
    setNewUser({ username: user.username, employeeName: user.employeeName || '', password: '', role: user.role });
  };

  const startNewUser = () => {
    setEditingUserId('new');
    setNewUser({ username: '', employeeName: '', password: '', role: 'viewer' });
  };

  const closeUserEditor = () => {
    setEditingUserId(null);
    setNewUser({ username: '', employeeName: '', password: '', role: 'viewer' });
  };

  const setUserNotice = (message: string) => setNotice(message);

  const updateUser = (userId: string, changes: Partial<Pick<FleetUser, 'role' | 'active'>>) => {
    if (!canManageUsers) return;
    if (userId === currentUser.id) {
      setNotice('لا يمكن تغيير صلاحية حسابك الحالي');
      return;
    }
    const nextUsers = users.map((user) => user.id === userId ? { ...user, ...changes } : user);
    if (!nextUsers.some((user) => user.active && user.role === 'admin')) {
      setNotice('يجب الاحتفاظ بمدير نظام نشط واحد على الأقل');
      return;
    }
    persistUsers(nextUsers);
    setNotice('تم تحديث صلاحية المستخدم');
  };

  const changeCurrentPassword = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!firebaseAuth?.currentUser) {
      setNotice('لا يوجد حساب Firebase مسجل حاليا');
      return;
    }
    if (newPassword.length < 6) {
      setNotice('يجب أن تتكون كلمة المرور من 6 أحرف أو أرقام على الأقل');
      return;
    }
    if (newPassword !== confirmPassword) {
      setNotice('كلمتا المرور غير متطابقتين');
      return;
    }
    try {
      await updatePassword(firebaseAuth.currentUser, newPassword);
      setNewPassword('');
      setConfirmPassword('');
      setNotice('تم تغيير كلمة المرور السحابية بنجاح');
    } catch (error) {
      setNotice((error as { code?: string }).code === 'auth/requires-recent-login'
        ? 'سجّل الخروج ثم ادخل مرة أخرى قبل تغيير كلمة المرور'
        : 'تعذر تغيير كلمة المرور. تحقق من اتصال الإنترنت');
    }
  };

  const sendUserPasswordReset = async (user: FleetUser) => {
    const email = user.email || (user.username.includes('@') ? user.username : '');
    if (!email) {
      setNotice('أضف البريد الإلكتروني لهذا المستخدم في Firebase لإرسال رابط تغيير كلمة المرور');
      return;
    }
    try {
      await sendPasswordResetEmail(firebaseAuth!, email);
      setNotice(`تم إرسال رابط تغيير كلمة المرور إلى ${email}`);
    } catch {
      setNotice('تعذر إرسال رابط تغيير كلمة المرور. تحقق من تفعيل Email/Password');
    }
  };

  const deleteUser = (userId: string) => {
    if (!canManageUsers || userId === currentUser.id) {
      setNotice('لا يمكن حذف حسابك الحالي');
      return;
    }
    const nextUsers = users.filter((user) => user.id !== userId);
    if (!nextUsers.some((user) => user.active && user.role === 'admin')) {
      setNotice('يجب الاحتفاظ بمدير نظام نشط واحد على الأقل');
      return;
    }
    persistUsers(nextUsers);
    setNotice('تم حذف المستخدم');
  };

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    if (cloudLoaded) void saveCloudFleetData(data);
  }, [cloudLoaded, data]);

  useEffect(() => {
    if (!isFirebaseConfigured) return undefined;
    let cancelled = false;
    readCloudFleetData<FleetData>().then((cloudData) => {
      if (cancelled) return;
      if (cloudData) setData(cloudData);
      setCloudLoaded(true);
    });
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    setAllowanceRateDraft(String(fuelAllowanceRate(data)));
  }, [data.allowancePricePerLiter]);

  useEffect(() => {
    const captureInstall = (event: Event) => {
      event.preventDefault();
      setInstallPrompt(event as InstallPrompt);
    };
    window.addEventListener('beforeinstallprompt', captureInstall);
    return () => window.removeEventListener('beforeinstallprompt', captureInstall);
  }, []);

  const summaries = useMemo(() => data.vehicles.map((vehicle) => {
    const entries = data.entries.filter((entry) => entry.vehicleId === vehicle.id);
    const distance = entries.reduce((total, entry) => total + Math.max(0, entry.currentOdometer - entry.startOdometer), 0);
    const liters = entries.reduce((total, entry) => total + litersForEntry(entry, vehicle), 0);
    const expected = litersForDistance(vehicle, distance);
    const cost = entries.reduce((total, entry) => total + entryTotalCost(entry), 0);
    const deviation = entries.reduce((total, entry) => total + entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)), 0);
    const lastOdometer = entries.reduce((latest, entry) => Math.max(latest, entry.currentOdometer), vehicle.openingOdometer);
    return { vehicle, entries, distance, liters, expected, deviation, cost, lastOdometer };
  }), [data]);

  const calculateFuelLiters = (vehicleId: string, startOdometer: number, currentOdometer: number) => {
    const vehicle = data.vehicles.find((item) => item.id === vehicleId);
    if (!vehicle) return 0;
    return litersForDistance(vehicle, currentOdometer - startOdometer);
  };
  const calculateAllowanceValue = (vehicleId: string, liters: number) => {
    return allowanceCost(liters, fuelAllowanceRate(data));
  };
  const allowanceValueForEntry = (entry: FuelEntry) => calculateAllowanceValue(entry.vehicleId, entry.liters);

  const totals = summaries.reduce((total, summary) => ({
    distance: total.distance + summary.distance,
    liters: total.liters + summary.liters,
    expected: total.expected + summary.expected,
    deviation: total.deviation + summary.deviation,
    cost: total.cost + summary.cost,
  }), { distance: 0, liters: 0, expected: 0, deviation: 0, cost: 0 });

  const selectedSummary = summaries.find((summary) => summary.vehicle.id === selectedVehicleId);
  const selectedEntries = data.entries
    .filter((entry) => entry.vehicleId === selectedVehicleId)
    .sort((a, b) => b.date.localeCompare(a.date));
  const selectedEntryDetails = selectedEntries.slice().reverse().reduce<Array<{
    entry: FuelEntry;
    distance: number;
    expectedLiters: number;
    actualFuelValue: number;
    deviation: number;
    accumulated: number;
    totalCost: number;
  }>>((rows, entry) => {
    const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
    const expectedLiters = selectedSummary ? litersForDistance(selectedSummary.vehicle, distance) : 0;
    const actualFuelValue = entryFuelCost(entry);
    const deviation = entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, selectedSummary?.vehicle));
    const accumulated = (rows[rows.length - 1]?.accumulated || 0) + deviation;
    rows.push({ entry, distance, expectedLiters, actualFuelValue, deviation, accumulated, totalCost: entryTotalCost(entry) });
    return rows;
  }, []).reverse();

  const filteredEntries = data.entries
    .slice()
    .sort((a, b) => b.date.localeCompare(a.date))
    .filter((entry) => {
      const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
      return `${vehicle?.name ?? ''} ${vehicle?.plate ?? ''} ${vehicle?.driver ?? ''} ${entry.note}`.toLowerCase().includes(search.toLowerCase());
    });

  const selectLedgerPeriod = (period: LedgerPeriod) => {
    setLedgerPeriod(period);
    if (period === 'month') return;
    setIsLedgerMonthPickerOpen(false);
    setLedgerMonthYear('');
    if (period === 'custom') {
      setLedgerFilters((current) => ({ ...current, from: '', to: '' }));
      return;
    }
    setLedgerFilters((current) => ({ ...current, from: '', to: '' }));
  };

  const selectLedgerMonthYear = (monthYear: string) => {
    const [year, month] = monthYear.split('-').map(Number);
    const monthIndex = month - 1;
    setLedgerMonthYear(monthYear);
    setLedgerPeriod('month');
    setIsLedgerMonthPickerOpen(false);
    setLedgerFilters((current) => ({
      ...current,
      from: dateInputValue(new Date(year, monthIndex, 1)),
      to: dateInputValue(new Date(year, monthIndex + 1, 0)),
    }));
  };

  const resetLedgerFilters = () => {
    setLedgerFilters({ vehicleId: 'all', from: '', to: '', search: '' });
    setLedgerPeriod('all');
    setLedgerMonthYear('');
    setIsLedgerMonthPickerOpen(false);
  };

  const ledgerEntries = useMemo(() => data.entries
    .slice()
    .sort((a, b) => b.date.localeCompare(a.date))
    .filter((entry) => {
      const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
      const searchable = `${vehicle?.name ?? ''} ${vehicle?.plate ?? ''} ${vehicle?.driver ?? ''} ${entry.driver ?? ''} ${entry.note}`.toLowerCase();
      return (ledgerFilters.vehicleId === 'all' || entry.vehicleId === ledgerFilters.vehicleId)
        && (!ledgerFilters.from || entry.date >= ledgerFilters.from)
        && (!ledgerFilters.to || entry.date <= ledgerFilters.to)
        && (!ledgerFilters.search.trim() || searchable.includes(ledgerFilters.search.trim().toLowerCase()));
    }), [data.entries, data.vehicles, ledgerFilters]);

  const ledgerTotals = ledgerEntries.reduce((total, entry) => {
    const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
    const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
    return {
      distance: total.distance + distance,
      liters: total.liters + litersForEntry(entry, vehicle),
      expected: total.expected + (vehicle ? litersForDistance(vehicle, distance) : 0),
      deviation: total.deviation + entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)),
      cost: total.cost + entryTotalCost(entry),
    };
  }, { distance: 0, liters: 0, expected: 0, deviation: 0, cost: 0 });
  const ledgerYears = Array.from(new Set([
    ...data.entries.map((entry) => Number(entry.date.slice(0, 4))),
    new Date().getFullYear(),
  ])).filter(Number.isInteger).sort((a, b) => b - a);

  const startEntry = (vehicleId = data.vehicles[0]?.id) => {
    const summary = summaries.find((item) => item.vehicle.id === vehicleId);
    const start = summary?.lastOdometer ?? 0;
    setEditingEntryId(null);
    setFuelAmountDraft('');
    setEntryPlateSearch('');
    setExpenseAmountDrafts(emptyExpenseInputs());
    setEntryForm({ vehicleId: vehicleId ?? '', driver: summaries.find((item) => item.vehicle.id === vehicleId)?.vehicle.driver || '', startOdometer: String(start), currentOdometer: '', liters: '0.0', pricePerLiter: '0', fuelActualValue: '', allowanceValue: '', date: today(), route: '', note: '', expenses: emptyExpenseInputs() });
    setEntryDialog(true);
  };

  const editEntry = (entry: FuelEntry) => {
    const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
    setEditingEntryId(entry.id);
    setFuelAmountDraft('');
    setExpenseAmountDrafts(emptyExpenseInputs());
    setEntryForm({
      vehicleId: entry.vehicleId,
      driver: entry.driver || vehicle?.driver || '',
      startOdometer: String(entry.startOdometer),
      currentOdometer: String(entry.currentOdometer),
      liters: calculateFuelLiters(entry.vehicleId, entry.startOdometer, entry.currentOdometer).toFixed(1),
      pricePerLiter: String(entry.pricePerLiter),
      fuelActualValue: entry.fuelActualExpression ?? String(entryFuelCost(entry)),
      allowanceValue: String(allowanceValueForEntry(entry)),
      date: entry.date,
      route: entry.route || '',
      note: entry.note || '',
      expenses: Object.fromEntries(expenseFields.map(({ key }) => [key, entry.expenseExpressions?.[key] ?? (entry.expenses?.[key] ? String(entry.expenses[key]) : '')])) as ExpenseInputs,
    });
    setEntryDialog(true);
  };

  const addFuelAmount = () => {
    const amount = parseSingleAmount(fuelAmountDraft.trim());
    if (amount === null || !fuelAmountDraft.trim()) {
      setNotice('أدخل مبلغًا واحدًا صحيحًا ثم اضغط إضافة');
      return;
    }
    setEntryForm((current) => ({ ...current, fuelActualValue: [...splitAmountExpression(current.fuelActualValue), String(amount)].join('+') }));
    setFuelAmountDraft('');
    setNotice('');
  };

  const removeFuelAmount = (indexToRemove: number) => {
    setEntryForm((current) => ({ ...current, fuelActualValue: splitAmountExpression(current.fuelActualValue).filter((_, index) => index !== indexToRemove).join('+') }));
  };

  const addExpenseAmount = (key: keyof ExpenseBreakdown) => {
    const draft = expenseAmountDrafts[key].trim();
    const amount = parseSingleAmount(draft);
    if (amount === null || !draft) {
      setNotice('أدخل مبلغًا واحدًا صحيحًا ثم اضغط إضافة');
      return;
    }
    setEntryForm((current) => ({
      ...current,
      expenses: { ...current.expenses, [key]: [...splitAmountExpression(current.expenses[key]), String(amount)].join('+') },
    }));
    setExpenseAmountDrafts((current) => ({ ...current, [key]: '' }));
    setNotice('');
  };

  const removeExpenseAmount = (key: keyof ExpenseBreakdown, indexToRemove: number) => {
    setEntryForm((current) => ({
      ...current,
      expenses: { ...current.expenses, [key]: splitAmountExpression(current.expenses[key]).filter((_, index) => index !== indexToRemove).join('+') },
    }));
  };

  const saveVehicle = (event: React.FormEvent) => {
    event.preventDefault();
    const vehicle: Vehicle = {
      id: editingVehicleId || `v${Date.now()}`,
      name: vehicleForm.name.trim(),
      plate: vehicleForm.plate.trim(),
      driver: vehicleForm.driver.trim(),
      openingOdometer: Number(vehicleForm.openingOdometer),
      kmPerLiter: Number(vehicleForm.kmPerLiter) > 0 ? 105 / Number(vehicleForm.kmPerLiter) : 5.25,
      fuelLitersPer105Km: Number(vehicleForm.kmPerLiter),
      fuelDistanceBaseKm: Number(vehicleForm.distanceBaseKm) > 0 ? Number(vehicleForm.distanceBaseKm) : 105,
    };
    setData((current) => ({ ...current, vehicles: editingVehicleId ? current.vehicles.map((item) => item.id === editingVehicleId ? vehicle : item) : [...current.vehicles, vehicle] }));
    setVehicleDialog(false);
    setEditingVehicleId(null);
    setVehicleForm({ name: '', plate: '', driver: '', openingOdometer: '', kmPerLiter: '20', distanceBaseKm: '105' });
    setNotice(editingVehicleId ? 'تم تحديث بيانات السيارة' : 'تمت إضافة السيارة');
  };

  const openVehicleEditor = (vehicle: Vehicle) => {
    setEditingVehicleId(vehicle.id);
    setVehicleForm({ name: vehicle.name, plate: vehicle.plate, driver: vehicle.driver, openingOdometer: String(vehicle.openingOdometer), kmPerLiter: String(vehicleFuelRate(vehicle)), distanceBaseKm: String(vehicleDistanceBase(vehicle)) });
    setVehicleDialog(true);
  };

  const downloadVehicleTemplate = () => {
    const rows = [
      ['رقم اللوحة', 'اسم السيارة', 'اسم السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)'],
      ...data.vehicles.map((vehicle) => [vehicle.plate, vehicle.name, vehicle.driver, vehicle.openingOdometer, vehicle.kmPerLiter]),
    ];
    const csv = `\uFEFF${rows.map((row) => row.map((value) => `"${String(value ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n')}`;
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = 'قالب-سيارات-الشركة.csv';
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const importVehicleTemplate = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const input = event.currentTarget;
    const file = input.files?.[0];
    if (!file) return;

    setIsImportingVehicles(true);
    try {
      const rows = parseCsvRows(await file.text());
      if (rows.length < 2) throw new Error('القالب لا يحتوي على صفوف سيارات.');

      const headers = rows[0].map((value) => value.trim());
      const requiredHeaders = ['رقم اللوحة', 'اسم السيارة', 'اسم السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)'];
      const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
      if (missingHeaders.length) throw new Error(`القالب غير مطابق. الأعمدة الناقصة: ${missingHeaders.join('، ')}`);

      const column = (header: string) => headers.indexOf(header);
      const imported: Array<Omit<Vehicle, 'id'>> = [];
      const seenPlates = new Set<string>();
      const rowErrors: string[] = [];
      rows.slice(1).forEach((row, index) => {
        const plate = (row[column('رقم اللوحة')] || '').trim();
        const name = (row[column('اسم السيارة')] || '').trim();
        const driver = (row[column('اسم السائق')] || '').trim();
        const openingOdometer = parseTemplateNumber(row[column('بداية العداد')] || '');
        const kmPerLiter = parseTemplateNumber(row[column('معدل الاستهلاك (كم/لتر)')] || '');
        if (!plate && !name && !row.some((value) => value.trim())) return;

        if (!plate || !name || !Number.isFinite(openingOdometer) || openingOdometer < 0 || !Number.isFinite(kmPerLiter) || kmPerLiter <= 0) {
          rowErrors.push(String(index + 2));
          return;
        }
        const normalizedPlate = plate.toLocaleLowerCase('ar-EG');
        if (seenPlates.has(normalizedPlate)) {
          rowErrors.push(String(index + 2));
          return;
        }
        seenPlates.add(normalizedPlate);
        imported.push({ plate, name, driver, openingOdometer, kmPerLiter });
      });

      if (rowErrors.length) throw new Error(`راجع بيانات الصفوف التالية (رقم اللوحة والاسم والعداد ومعدل الاستهلاك): ${rowErrors.join('، ')}`);
      if (!imported.length) throw new Error('لا توجد سيارات صالحة للاستيراد.');

      const existingPlates = new Set(data.vehicles.map((vehicle) => vehicle.plate.trim().toLocaleLowerCase('ar-EG')));
      const addedCount = imported.filter((vehicle) => !existingPlates.has(vehicle.plate.trim().toLocaleLowerCase('ar-EG'))).length;
      const updatedCount = imported.length - addedCount;
      setData((current) => {
        const vehicles = [...current.vehicles];
        imported.forEach((vehicle, index) => {
          const matchIndex = vehicles.findIndex((item) => item.plate.trim().toLocaleLowerCase('ar-EG') === vehicle.plate.trim().toLocaleLowerCase('ar-EG'));
          if (matchIndex >= 0) vehicles[matchIndex] = { ...vehicles[matchIndex], ...vehicle };
          else vehicles.push({ id: `v${Date.now()}-${index}`, ...vehicle });
        });
        return { ...current, vehicles };
      });
      setNotice(`تم استيراد ${imported.length} سيارة: إضافة ${addedCount} وتحديث ${updatedCount}`);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'تعذر استيراد الملف.');
    } finally {
      setIsImportingVehicles(false);
      input.value = '';
    }
  };

  const downloadSelectedVehicleTemplate = async () => {
    if (!selectedSummary) return;
    const ExcelJS = (await import('exceljs')).default;
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'مراقبة سيارات شركة أبو غالي';
    workbook.subject = `قالب مصروفات السيارة ${selectedSummary.vehicle.plate}`;
    workbook.calcProperties.fullCalcOnLoad = true;
    const info = workbook.addWorksheet('بيانات السيارة', { views: [{ rightToLeft: true, showGridLines: false }] });
    info.columns = Array.from({ length: 6 }, () => ({ width: 22 }));
    info.mergeCells('A1:F1');
    info.getCell('A1').value = `قالب مصروفات السيارة - ${selectedSummary.vehicle.plate}`;
    info.getCell('A1').font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    info.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF176B64' } };
    info.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    info.getRow(1).height = 34;
    const infoRows: Array<[string, string | number, string, string | number, string, string | number]> = [
      ['رقم اللوحة', selectedSummary.vehicle.plate, 'اسم السيارة', selectedSummary.vehicle.name, 'اسم السائق', selectedSummary.vehicle.driver],
      ['بداية العداد', selectedSummary.vehicle.openingOdometer, 'معدل الاستهلاك (كم/لتر)', selectedSummary.vehicle.kmPerLiter, 'سعر التموين (جنيه/لتر)', fuelAllowanceRate(data)],
    ];
    infoRows.forEach((values, rowIndex) => {
      const row = info.getRow(rowIndex + 3);
      values.forEach((value, columnIndex) => {
        const cell = row.getCell(columnIndex + 1);
        cell.value = value;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin', color: { argb: 'FFBAC8C3' } }, bottom: { style: 'thin', color: { argb: 'FFBAC8C3' } }, left: { style: 'thin', color: { argb: 'FFBAC8C3' } }, right: { style: 'thin', color: { argb: 'FFBAC8C3' } } };
        if (columnIndex % 2 === 0) {
          cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF304A41' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5EEE8' } };
        } else {
          cell.font = { name: 'Arial', size: 11, color: { argb: 'FF172C32' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F9F6' } };
        }
      });
      row.height = 25;
    });
    info.mergeCells('A6:F6');
    info.getCell('A6').value = 'أدخل البيانات في الخانات الملونة بصفحة سجل السيارة. الأعمدة المحسوبة تتحدث تلقائيًا عند فتح الملف في Excel.';
    info.getCell('A6').font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF64766D' } };
    info.getCell('A6').alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
    info.getRow(6).height = 30;

    const sheet = workbook.addWorksheet('سجل السيارة', { views: [{ rightToLeft: true, showGridLines: false, state: 'frozen', ySplit: 4, xSplit: 3 }] });
    const headers = [
      'معرف الحركة', 'رقم اللوحة', 'السيارة', 'التاريخ', 'اسم السائق', 'خط السير',
      'القراءة السابقة', 'القراءة الحالية', 'الفجوة (كم)', 'التموين الفعلي (لتر - محسوب)',
      'قيمة التموين المقررة (جنيه - محسوبة)', 'مصروف الوقود الفعلي (جنيه)',
      'الانحراف (جنيه - محسوب)', 'التراكم (جنيه - محسوب)',
      ...expenseFields.map((field) => `${field.label} (جنيه)`), 'التفصيل', 'إجمالي مصاريف السيارة (جنيه - محسوب)',
    ];
    sheet.columns = [12, 17, 22, 15, 20, 25, 17, 17, 15, 20, 24, 24, 20, 19, ...expenseFields.map(() => 17), 26, 26].map((width) => ({ width }));
    sheet.mergeCells(1, 1, 1, headers.length);
    sheet.getCell(1, 1).value = `كشف الحركة والمصاريف - ${selectedSummary.vehicle.name} (${selectedSummary.vehicle.plate})`;
    sheet.getCell(1, 1).font = { name: 'Arial', size: 15, bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getCell(1, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF176B64' } };
    sheet.getCell(1, 1).alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(1).height = 34;
    sheet.mergeCells(2, 1, 2, headers.length);
    sheet.getCell(2, 1).value = 'املأ الخانات ذات الخلفية الفاتحة. الخانات الرمادية محسوبة تلقائيًا. احتفظ بمعرف الحركة لتحديث الحركة القديمة عند إعادة الاستيراد.';
    sheet.getCell(2, 1).font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF52675E' } };
    sheet.getCell(2, 1).alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
    sheet.getRow(2).height = 28;
    sheet.addRow([]);
    const headerRow = sheet.addRow(headers);
    headerRow.height = 42;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF247C73' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { top: { style: 'thin', color: { argb: 'FFB7C7C1' } }, bottom: { style: 'thin', color: { argb: 'FFB7C7C1' } }, left: { style: 'thin', color: { argb: 'FFB7C7C1' } }, right: { style: 'thin', color: { argb: 'FFB7C7C1' } } };
    });

    const movementRows = selectedEntries.slice().reverse();
    const rowCount = Math.max(movementRows.length, 40);
    for (let index = 0; index < rowCount; index += 1) {
      const entry = movementRows[index];
      const rowNumber = index + 5;
      const expenses = entry?.expenses || {};
      const values: ExcelJS.CellValue[] = [
        entry?.id || '', selectedSummary.vehicle.plate, selectedSummary.vehicle.name,
        entry ? new Date(`${entry.date}T00:00:00`) : null,
        entry?.driver || selectedSummary.vehicle.driver, entry?.route || '',
        entry?.startOdometer ?? (index === 0 ? selectedSummary.lastOdometer : null), entry?.currentOdometer ?? null,
        { formula: `IF(OR(G${rowNumber}="",H${rowNumber}=""),"",MAX(0,H${rowNumber}-G${rowNumber}))` },
        { formula: `IF(OR(I${rowNumber}="",'بيانات السيارة'!$D$4<=0),"",I${rowNumber}/'بيانات السيارة'!$D$4)` },
        { formula: `IF(J${rowNumber}="","",ROUND(J${rowNumber}*'بيانات السيارة'!$F$4,0))` },
        entry ? entryFuelCost(entry) : null,
        { formula: `IF(L${rowNumber}="","",L${rowNumber}-K${rowNumber})` },
        { formula: `IF(M${rowNumber}="","",SUM($M$5:M${rowNumber}))` },
        ...expenseFields.map((field) => entry ? Math.round(expenses[field.key] || 0) : null),
        entry?.note || '',
        { formula: `IF(COUNTA(L${rowNumber},O${rowNumber}:Z${rowNumber})=0,"",ROUND(SUM(L${rowNumber},O${rowNumber}:Z${rowNumber}),0))` },
      ];
      const row = sheet.addRow(values);
      row.height = 25;
      row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
        cell.font = { name: 'Arial', size: 10, color: { argb: 'FF263A34' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { bottom: { style: 'hair', color: { argb: 'FFD7E0DA' } }, left: { style: 'hair', color: { argb: 'FFE0E6E1' } }, right: { style: 'hair', color: { argb: 'FFE0E6E1' } } };
        const calculated = [9, 10, 11, 13, 14, headers.length].includes(columnNumber);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: calculated ? 'FFE7EEE9' : index % 2 ? 'FFFAFCFA' : 'FFF0F6F2' } };
        if (columnNumber === 4) cell.numFmt = 'yyyy-mm-dd';
        if ([11, 12, 13, 14, headers.length].includes(columnNumber)) cell.numFmt = '#,##0 "جنيه"';
        if ([7, 8, 9].includes(columnNumber)) cell.numFmt = '#,##0';
        if (columnNumber === 10) cell.numFmt = '0.0';
      });
    }
    sheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4, column: headers.length } };
    sheet.getColumn(4).numFmt = 'yyyy-mm-dd';

    const buffer = await workbook.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buffer as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = `قالب-${selectedSummary.vehicle.plate.replace(/[\\/:*?"<>|\s]+/g, '-')}.xlsx`;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const importSelectedVehicleTemplate = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const input = event.currentTarget;
    const file = input.files?.[0];
    if (!file || !selectedSummary) return;
    const vehicle = selectedSummary.vehicle;
    setIsImportingVehicleHistory(true);
    try {
      let rows: string[][];
      if (/\.xlsx$/i.test(file.name)) {
        const ExcelJS = (await import('exceljs')).default;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        const worksheet = workbook.getWorksheet('سجل السيارة') || workbook.worksheets.find((sheet) => sheet.name !== 'بيانات السيارة');
        if (!worksheet) throw new Error('لم أجد ورقة سجل السيارة داخل ملف Excel.');
        rows = worksheetToRows(worksheet);
      } else {
        rows = parseCsvRows(await file.text());
      }
      if (rows.length < 2) throw new Error('القالب لا يحتوي على حركات.');
      const headerRowIndex = rows.findIndex((row) => row.some((value) => value.trim() === 'معرف الحركة'));
      if (headerRowIndex < 0) throw new Error('لم أجد صف عناوين قالب الحركات.');
      const headers = rows[headerRowIndex].map((value) => value.trim());
      const required = [
        'معرف الحركة', 'رقم اللوحة', 'التاريخ', 'اسم السائق', 'خط السير',
        'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)', 'التفصيل',
        ...expenseFields.map((field) => `${field.label} (جنيه)`),
      ];
      const missing = required.filter((header) => !headers.includes(header));
      if (missing.length) throw new Error(`القالب غير مطابق. الأعمدة الناقصة: ${missing.join('، ')}`);
      const column = (header: string) => headers.indexOf(header);
      const oldEntries = data.entries.filter((entry) => entry.vehicleId === vehicle.id);
      const oldIds = new Set(oldEntries.map((entry) => entry.id));
      const seenIds = new Set<string>();
      const imported: FuelEntry[] = [];
      const errors: string[] = [];

      rows.slice(headerRowIndex + 1).forEach((row, index) => {
        const value = (header: string) => (row[column(header)] || '').trim();
        const hasMovement = ['التاريخ', 'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)'].some((header) => value(header) !== '');
        if (!hasMovement) return;
        const rowPlate = value('رقم اللوحة').toLocaleLowerCase('ar-EG');
        const date = value('التاريخ');
        const start = parseTemplateNumber(value('القراءة السابقة'));
        const current = parseTemplateNumber(value('القراءة الحالية'));
        const actualFuelValue = parseTemplateNumber(value('مصروف الوقود الفعلي (جنيه)'));
        if (rowPlate !== vehicle.plate.trim().toLocaleLowerCase('ar-EG') || !date || !Number.isFinite(start) || !Number.isFinite(current) || current < start || !Number.isFinite(actualFuelValue) || actualFuelValue < 0) {
          errors.push(String(index + headerRowIndex + 2));
          return;
        }
        const suppliedId = value('معرف الحركة');
        if (suppliedId && seenIds.has(suppliedId)) {
          errors.push(String(index + headerRowIndex + 2));
          return;
        }
        if (suppliedId) seenIds.add(suppliedId);
        const id = suppliedId && oldIds.has(suppliedId) ? suppliedId : `e${Date.now()}-${index}`;
        const liters = calculateFuelLiters(vehicle.id, start, current);
        const expenses = Object.fromEntries(expenseFields.map((field) => [field.key, Math.max(0, Math.round(parseTemplateNumber(value(`${field.label} (جنيه)`)) || 0))])) as ExpenseBreakdown;
        imported.push({
          id, vehicleId: vehicle.id, date, startOdometer: start, currentOdometer: current,
          liters, pricePerLiter: liters > 0 ? actualFuelValue / liters : 0, fuelActualValue: actualFuelValue,
          driver: value('اسم السائق'), route: value('خط السير'), allowanceValue: calculateAllowanceValue(vehicle.id, liters),
          expenses, note: value('التفصيل'),
        });
      });

      if (errors.length) throw new Error(`راجع الصفوف التالية: ${errors.join('، ')}. تأكد من رقم اللوحة والعدادات والمصروف.`);
      if (!imported.length) throw new Error('لم يتم العثور على حركات صالحة في القالب.');
      const updatedCount = imported.filter((entry) => oldIds.has(entry.id)).length;
      const addedCount = imported.length - updatedCount;
      setData((current) => {
        const importedById = new Map(imported.map((entry) => [entry.id, entry]));
        const existingIds = new Set(current.entries.map((entry) => entry.id));
        const entries = current.entries.map((entry) => importedById.get(entry.id) || entry);
        imported.forEach((entry) => { if (!existingIds.has(entry.id)) entries.push(entry); });
        return { ...current, entries };
      });
      setNotice(`تم استيراد سجل ${vehicle.plate}: تحديث ${updatedCount} وإضافة ${addedCount}`);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'تعذر استيراد سجل السيارة.');
    } finally {
      setIsImportingVehicleHistory(false);
      input.value = '';
    }
  };

  const saveEntry = (event: React.FormEvent) => {
    event.preventDefault();
    if (fuelAmountDraft.trim() || (Object.values(expenseAmountDrafts) as string[]).some((amount) => amount.trim())) {
      setNotice('اضغط «إضافة» بجوار كل مبلغ لإدخاله في الإجمالي قبل الحفظ');
      return;
    }
    if (!entryForm.fuelActualValue.trim()) {
      setNotice('أضف مبلغ مصروف الوقود قبل الحفظ');
      return;
    }
    const start = Number(entryForm.startOdometer);
    const current = Number(entryForm.currentOdometer);
    if (current < start) {
      setNotice('العداد الحالي يجب أن يكون أكبر من أو يساوي بداية العداد');
      return;
    }
    const fuelActualValue = parseAmountExpression(entryForm.fuelActualValue);
    const parsedExpenses = Object.fromEntries(expenseFields.map(({ key }) => [key, parseAmountExpression(entryForm.expenses[key])])) as Record<keyof ExpenseBreakdown, number | null>;
    if (fuelActualValue === null || Object.values(parsedExpenses).some((amount) => amount === null)) {
      setNotice('راجع المبالغ المضافة وتأكد من صحتها');
      return;
    }
    const calculatedLiters = calculateFuelLiters(entryForm.vehicleId, start, current);
    const entry: FuelEntry = {
      id: editingEntryId || `e${Date.now()}`,
      vehicleId: entryForm.vehicleId,
      date: entryForm.date,
      startOdometer: start,
      currentOdometer: current,
      liters: calculatedLiters,
      pricePerLiter: calculatedLiters > 0 ? fuelActualValue / calculatedLiters : 0,
      fuelActualValue,
      fuelActualExpression: entryForm.fuelActualValue.trim(),
      driver: entryForm.driver.trim(),
      allowanceValue: calculateAllowanceValue(entryForm.vehicleId, calculatedLiters),
      route: entryForm.route.trim(),
      expenses: Object.fromEntries(expenseFields.map(({ key }) => [key, parsedExpenses[key] ?? 0])) as ExpenseBreakdown,
      expenseExpressions: Object.fromEntries(expenseFields.map(({ key }) => [key, entryForm.expenses[key].trim()])),
      note: entryForm.note.trim(),
    };
    setData((previous) => ({
      ...previous,
      entries: editingEntryId
        ? previous.entries.map((item) => item.id === editingEntryId ? entry : item)
        : [entry, ...previous.entries],
    }));
    setEntryDialog(false);
    setEditingEntryId(null);
    setNotice(editingEntryId ? 'تم تحديث الحركة' : 'تم حفظ سجل التعبئة وتحديث الفجوة التموينية');
  };

  const confirmDeleteEntry = () => {
    if (!pendingDeleteEntry || deleteConfirmationText.trim() !== 'موافق') return;
    setData((current) => ({ ...current, entries: current.entries.filter((entry) => entry.id !== pendingDeleteEntry.id) }));
    setPendingDeleteEntry(null);
    setDeleteConfirmationText('');
    setNotice('تم حذف سجل التعبئة');
  };

  const openDeleteConfirmation = (entry: FuelEntry) => {
    setPendingDeleteEntry(entry);
    setDeleteConfirmationText('');
  };

  const cancelDeleteConfirmation = () => {
    setPendingDeleteEntry(null);
    setDeleteConfirmationText('');
  };

  const exportExcel = async () => {
    const rows = [
      ['التاريخ', 'اسم السائق', 'خط السير', 'القراءة السابقة', 'القراءة الحالية', 'الفجوة كم', 'عدد اللترات', 'قيمة التموين وفقًا للتر (جنيه)', 'قيمة التموين الفعلية (جنيه)', 'الانحراف (جنيه)', 'التراكم (جنيه)', ...expenseFields.map((field) => `${field.label} (جنيه)`), 'إجمالي مصاريف السيارة (جنيه)', 'رقم اللوحة', 'ملاحظات'],
      ...filteredEntries.map((entry) => {
        const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
        const liters = litersForEntry(entry, vehicle);
        const allowanceValue = allowanceCost(liters, fuelAllowanceRate(data));
        const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters);
        const accumulated = data.entries
          .filter((item) => item.vehicleId === entry.vehicleId && item.date <= entry.date)
          .reduce((sum, item) => sum + entryDeviation(item, fuelAllowanceRate(data)), 0);
        return [entry.date, entry.driver || vehicle?.driver || '', entry.route || '', entry.startOdometer, entry.currentOdometer, Math.max(0, entry.currentOdometer - entry.startOdometer), liters, allowanceValue, entryFuelCost(entry), deviation, accumulated, ...expenseFields.map((field) => Math.round(entry.expenses?.[field.key] || 0)), entryTotalCost(entry), vehicle?.plate || '', entry.note];
      }),
    ];
    const financialSummary = filteredEntries.reduce((summary, entry) => {
      const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
      const liters = litersForEntry(entry, vehicle);
      return {
        fuel: summary.fuel + entryFuelCost(entry),
        other: summary.other + sumExpenses(entry.expenses),
        total: summary.total + entryTotalCost(entry),
        deviation: summary.deviation + entryDeviation(entry, fuelAllowanceRate(data), liters),
      };
    }, { fuel: 0, other: 0, total: 0, deviation: 0 });
    const ExcelJS = (await import('exceljs')).default;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('سجل التعبئة', { views: [{ rightToLeft: true }] });
    const lastColumn = String.fromCharCode(64 + rows[0].length);
    worksheet.mergeCells(`A1:${lastColumn}1`);
    worksheet.getCell('A1').value = 'ملخص سجل مصروفات السيارات';
    worksheet.getCell('A1').font = { name: 'Arial', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7E2033' } };
    worksheet.getRow(1).height = 32;

    worksheet.mergeCells(`A2:${lastColumn}2`);
    worksheet.getCell('A2').value = `تاريخ التصدير: ${new Date().toLocaleDateString('ar-EG')} | عدد الحركات: ${filteredEntries.length}`;
    worksheet.getCell('A2').font = { name: 'Arial', size: 11, italic: true, color: { argb: 'FF5D4A4D' } };
    worksheet.getCell('A2').alignment = { horizontal: 'center' };

    const summaryLabels = worksheet.addRow(['إجمالي التموين', 'إجمالي باقي المصاريف', 'إجمالي المصاريف', 'إجمالي الانحراف']);
    const summaryValues = worksheet.addRow([financialSummary.fuel, financialSummary.other, financialSummary.total, financialSummary.deviation]);
    [summaryLabels, summaryValues].forEach((row) => {
      row.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF263744' } };
      row.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    summaryLabels.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9D7A5' } };
    summaryValues.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF8E5' } };
    summaryValues.eachCell((cell) => { cell.numFmt = '#,##0.00 "جنيه"'; });
    summaryLabels.height = 23;
    summaryValues.height = 26;
    worksheet.addRow([]);
    const headerRow = worksheet.addRow(rows[0]);
    headerRow.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF17466E' } };
    headerRow.height = 30;
    rows.slice(1).forEach((row) => worksheet.addRow(row));
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber >= 7 && rowNumber % 2 === 1) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F9FB' } };
      if (rowNumber >= 6) row.alignment = { vertical: 'middle', horizontal: 'right' };
      if (rowNumber >= 6) row.font = { name: 'Arial', size: 10, color: { argb: 'FF263744' } };
    });
    headerRow.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.views = [{ rightToLeft: true, state: 'frozen', ySplit: 6 }];
    worksheet.autoFilter = { from: `A6`, to: `${lastColumn}${rows.length + 5}` };
    worksheet.columns.forEach((column, index) => { column.width = index === 0 ? 14 : index === 1 ? 22 : 18; });
    worksheet.getRow(5).height = 8;
    const buffer = await workbook.xlsx.writeBuffer();
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    link.download = 'سجل-مصروفات-السيارات.xlsx';
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  };

  const installApp = async () => {
    if (window.matchMedia('(display-mode: standalone)').matches) {
      setNotice('التطبيق مثبت بالفعل على هذا الجهاز');
      return;
    }
    const availablePrompt = installPrompt || window.deferredInstallPrompt;
    if (!availablePrompt) {
      setNotice('افتح قائمة Chrome ثم اختر «تثبيت التطبيق» أو «إضافة إلى الشاشة الرئيسية»');
      return;
    }
    await availablePrompt.prompt();
    const choice = await availablePrompt.userChoice;
    setNotice(choice.outcome === 'accepted' ? 'تم تثبيت التطبيق على جهازك' : 'تم إلغاء التثبيت');
    delete window.deferredInstallPrompt;
    setInstallPrompt(null);
  };

  const updateSystem = async () => {
    try {
      setNotice('جاري تطبيق التحديث...');
      await applyUpdate();
    } catch {
      setNotice('تعذر تطبيق التحديث. أعد المحاولة.');
    }
  };

  const handleUpdateNow = async () => {
    try {
      setNotice('جاري تطبيق التحديث...');
      await applyUpdate();
    } catch {
      setNotice('تعذر تطبيق التحديث. أعد المحاولة.');
    }
  };

  const handleUpdateLater = () => {
    setUpdateAvailable(false);
    setNotice('يمكنك التحديث لاحقاً من زر "تحديث النظام"');
  };

  const saveAllowanceRate = () => {
    const rate = Number(allowanceRateDraft);
    if (!Number.isFinite(rate) || rate <= 0) {
      setNotice('أدخل سعرًا صحيحًا أكبر من صفر');
      return;
    }
    setData((current) => ({ ...current, allowancePricePerLiter: rate }));
    setAllowanceRateDraft(rate.toFixed(2));
    setIsRateEditorOpen(false);
    setNotice('تم تحديث سعر التموين لجميع السيارات');
  };

  const entryVehicle = data.vehicles.find((vehicle) => vehicle.id === entryForm.vehicleId);
  const entryDistance = Math.max(0, Number(entryForm.currentOdometer) - Number(entryForm.startOdometer));
  const fuelActualValuePreview = parseAmountExpression(entryForm.fuelActualValue);
  const expenseAmountPreviews = Object.fromEntries(expenseFields.map(({ key }) => [key, parseAmountExpression(entryForm.expenses[key])])) as Record<keyof ExpenseBreakdown, number | null>;
  const hasInvalidAmount = fuelActualValuePreview === null || Object.values(expenseAmountPreviews).some((amount) => amount === null);
  const hasUnaddedAmounts = Boolean(fuelAmountDraft.trim() || (Object.values(expenseAmountDrafts) as string[]).some((amount) => amount.trim()));
  const entryCostPreview = Math.round((fuelActualValuePreview ?? 0) + Object.values(expenseAmountPreviews).reduce<number>((sum, amount) => sum + (amount ?? 0), 0));
  const navigationTiles = [
    ...fleetSections,
    ...(canManageUsers ? [fleetHomeActions[0]] : []),
    ...(tab === 'home' ? [fleetHomeActions[1], fleetHomeActions[2]] : []),
  ];
  const openSection = fleetSections.find((section) => section.key === tab)
    || (tab === 'users' ? fleetHomeActions[0] : undefined);

  return (
    <main className={`fleet-app fleet-font-${fontSizeKey}${themeKey === 'dark' ? ' fleet-theme-dark' : ''}${selectedSummary ? ' fleet-app-detail' : ''}${tab === 'home' ? ' fleet-app-home' : ''}`} dir="rtl" style={{ '--teal': theme.accent, '--theme-heading': theme.heading, '--ink': theme.ink, '--line': theme.line, '--surface': theme.surface, '--surface-soft': theme.soft, '--app-background': theme.background } as React.CSSProperties}>
      {updateAvailable && <div className="fleet-update-banner" role="alert">
        <span className="fleet-update-icon" aria-hidden="true">🔄</span>
        <div className="fleet-update-text">
          <strong>يوجد تحديث جديد</strong>
          <small>نسخة أحدث من التطبيق جاهزة للتثبيت</small>
        </div>
        <div className="fleet-update-actions">
          <button type="button" className="fleet-update-btn fleet-update-now" onClick={() => void handleUpdateNow()}>تحديث الآن</button>
          <button type="button" className="fleet-update-btn fleet-update-later" onClick={handleUpdateLater}>لاحقاً</button>
        </div>
      </div>}
      {!selectedSummary && <>
      <header className="fleet-header">
        <button className="fleet-brand fleet-brand-button" type="button" onClick={() => { setTab('home'); setSelectedVehicleId(null); }} aria-label="العودة للرئيسية">
          <span className="fleet-brand-mark"><img src="/car-mark.svg" alt="" /></span>
          <span className="fleet-brand-copy"><strong>مراقبة <em>سيارات أبو غالي</em></strong></span>
        </button>
        <div className="fleet-header-actions">
          <span className="fleet-live"><i /> متابعة مباشرة</span>
          <button className="fleet-theme-button" type="button" onClick={() => setIsThemePickerOpen((open) => !open)}>ألوان التطبيق</button>
          <button className="fleet-install" type="button" onClick={installApp}>تثبيت التطبيق</button>
          {onLogout && <button className="fleet-logout" type="button" onClick={onLogout}>تسجيل الخروج</button>}
        </div>
      </header>

      {isThemePickerOpen && <div className="fleet-theme-picker" aria-label="قوالب ألوان التطبيق">
        <strong>قالب الألوان</strong>
        <div>{(Object.keys(themePresets) as ThemeKey[]).map((key) => <button key={key} type="button" className={themeKey === key ? 'is-active' : ''} onClick={() => chooseTheme(key)}><i style={{ background: themePresets[key].accent }} />{themePresets[key].label}</button>)}</div>
        <strong>حجم الخط</strong>
        <div>{(Object.keys(fontSizePresets) as FontSizeKey[]).map((key) => <button key={key} type="button" className={fontSizeKey === key ? 'is-active' : ''} onClick={() => chooseFontSize(key)}>{fontSizePresets[key]}</button>)}</div>
      </div>}

      {tab === 'home' ? <nav className="fleet-tabs fleet-tabs-home" aria-label="أقسام متابعة السيارات">
        {navigationTiles.map(({ key, label, description, Icon }) => (
          <button key={key} type="button" onClick={() => {
            if (key === 'install') { void installApp(); return; }
            if (key === 'update') { void updateSystem(); return; }
            if (key === 'users') { setTab('users'); return; }
            setTab(key);
            if (key !== 'vehicles') setSelectedVehicleId(null);
          }}>
            <span className="fleet-tab-icon"><Icon aria-hidden="true" strokeWidth={1.8} /></span>
            <span className="fleet-tab-copy"><strong>{label}</strong></span>
          </button>
        ))}
      </nav> : openSection ? <nav className="fleet-section-bar" aria-label="القسم المفتوح">
        <div className="fleet-section-bar-center" dir="rtl">
          <button type="button" className="fleet-section-title-button" aria-label={`العودة للرئيسية من ${openSection.label}`} onClick={() => { setTab('home'); setSelectedVehicleId(null); }}>{openSection.label}</button>
          <button type="button" className="fleet-section-back-button" aria-label="رجوع إلى الصفحة الرئيسية" onClick={() => { setTab('home'); setSelectedVehicleId(null); }}><ArrowRight aria-hidden="true" /></button>
        </div>
      </nav> : null}
      </>}

      {notice && <button className="fleet-notice" type="button" onClick={() => setNotice('')}>{notice}<span>×</span></button>}

      {tab === 'users' && canManageUsers && <section className="fleet-section fleet-user-management">
        <div className="fleet-section-heading">
          <div><span className="fleet-overline">الحسابات والصلاحيات</span><h2>إدارة المستخدمين والصلاحيات</h2></div>
          {!editingUserId && <button type="button" className="fleet-button fleet-button-primary" onClick={startNewUser}>إضافة مستخدم</button>}
          <button type="button" className="fleet-link-button" onClick={() => setTab('home')}>العودة للرئيسية</button>
        </div>
        {editingUserId ? <form className="fleet-user-edit-page" onSubmit={addUser}>
          <div className="fleet-user-edit-header"><div><span className="fleet-overline">بيانات المستخدم</span><h3>{editingUserId === 'new' ? 'إضافة مستخدم جديد' : 'تعديل بيانات المستخدم'}</h3></div><button type="button" className="fleet-link-button" onClick={closeUserEditor}>العودة للقائمة</button></div>
          <div className="fleet-user-edit-grid">
            <label>اسم المستخدم<input required value={newUser.username} onChange={(event) => setNewUser((current) => ({ ...current, username: event.target.value }))} autoComplete="off" /></label>
            <label>اسم الموظف<input value={newUser.employeeName} onChange={(event) => setNewUser((current) => ({ ...current, employeeName: event.target.value }))} autoComplete="off" /></label>
            <label>الدور والصلاحية<select value={newUser.role} onChange={(event) => setNewUser((current) => ({ ...current, role: event.target.value as FleetRole }))}><option value="admin">مدير النظام</option><option value="manager">مدير فرع</option><option value="viewer">مشاهد</option><option value="accountant">محاسب</option><option value="warehouse">مسؤول مخزن</option></select></label>
            <label>كلمة مرور محلية جديدة<input type="password" value={newUser.password} onChange={(event) => setNewUser((current) => ({ ...current, password: event.target.value }))} placeholder="اختياري" autoComplete="new-password" /></label>
          </div>
          <p className="fleet-user-storage-note">تغيير كلمة مرور Firebase يتم من زر إرسال رابط التغيير بعد العودة إلى القائمة.</p>
          <div className="fleet-user-edit-actions"><button className="fleet-button fleet-button-primary" type="submit">{editingUserId === 'new' ? 'إضافة المستخدم' : 'حفظ بيانات المستخدم'}</button><button className="fleet-button fleet-button-light" type="button" onClick={closeUserEditor}>إلغاء</button></div>
        </form> : <>
        <div className="fleet-user-table-wrap"><table className="fleet-table fleet-user-table">
          <thead><tr><th>اسم المستخدم</th><th>اسم الموظف</th><th>الحالة</th><th>كلمة المرور</th><th>إجراءات</th></tr></thead>
          <tbody>{users.map((user) => <tr key={user.id}>
            <td><strong>{user.username}</strong>{user.id === currentUser.id && <small>الحساب الحالي</small>}</td>
            <td>{user.employeeName || '—'}</td>
            <td><button type="button" className={`fleet-user-status${user.active ? ' is-active' : ''}`} disabled={user.id === currentUser.id} onClick={() => updateUser(user.id, { active: !user.active })}>{user.active ? 'نشط' : 'معطل'}</button></td>
            <td><button type="button" className="fleet-link-button" onClick={() => void sendUserPasswordReset(user)}>إرسال رابط التغيير</button></td>
            <td className="fleet-user-actions"><button type="button" className="fleet-row-action fleet-row-action-permissions" onClick={() => setUserNotice(`صلاحيات ${user.username}: ${fleetRoleLabels[user.role]}`)}>الصلاحيات</button><button type="button" className="fleet-row-action fleet-row-action-lock" disabled={user.id === currentUser.id} onClick={() => updateUser(user.id, { active: !user.active })}>{user.active ? 'قفل' : 'فتح'}</button><button type="button" className="fleet-row-action fleet-row-action-edit" onClick={() => editUser(user)}>تعديل</button><button type="button" className="fleet-row-action fleet-row-action-delete" disabled={user.id === currentUser.id} onClick={() => deleteUser(user.id)}>حذف</button></td>
          </tr>)}</tbody>
        </table></div>
        <p className="fleet-user-storage-note">تغيير كلمة المرور أعلاه يتم عبر Firebase Authentication للحساب المسجل حاليا.</p>
        </>}
      </section>}

      {tab === 'overview' && <>
        <section className="ledger-filters" aria-label="فلترة كشف الوقود">
          <div className="ledger-filters-heading"><h2>فلترة سجل الوقود</h2></div>
          <div className="ledger-period-controls" role="group" aria-label="الفترة الزمنية لسجل الوقود">
            <button type="button" data-period="all" className={ledgerPeriod === 'all' ? 'is-active' : ''} aria-pressed={ledgerPeriod === 'all'} onClick={() => selectLedgerPeriod('all')}>كل السجلات</button>
            <div className="ledger-month-picker">
              <button
                type="button"
                className={`ledger-month-select${ledgerPeriod === 'month' ? ' is-active' : ''}`}
                aria-label="اختيار الشهر والسنة"
                aria-haspopup="dialog"
                aria-expanded={isLedgerMonthPickerOpen}
                onClick={() => {
                  setLedgerPickerYear(ledgerPeriod === 'month' ? Number(ledgerMonthYear.slice(0, 4)) : new Date().getFullYear());
                  setIsLedgerMonthPickerOpen((open) => !open);
                }}
              >
                {ledgerPeriod === 'month' ? new Intl.DateTimeFormat('ar-EG', { month: 'long', year: 'numeric' }).format(new Date(`${ledgerMonthYear}-01T00:00:00`)) : 'اختر الشهر والسنة'}
              </button>
              {isLedgerMonthPickerOpen && <div className="ledger-month-picker-panel" role="dialog" aria-label="اختيار الشهر والسنة">
                <div className="ledger-month-picker-header">
                  <button type="button" aria-label="السنة السابقة" onClick={() => setLedgerPickerYear((year) => year - 1)}>‹</button>
                  <strong>{ledgerPickerYear}</strong>
                  <button type="button" aria-label="السنة التالية" onClick={() => setLedgerPickerYear((year) => year + 1)}>›</button>
                </div>
                <div className="ledger-month-picker-months">
                  {monthNames.map((month, index) => (
                    <button
                      key={month}
                      type="button"
                      className={ledgerPeriod === 'month' && ledgerMonthYear === `${ledgerPickerYear}-${String(index + 1).padStart(2, '0')}` ? 'is-selected' : ''}
                      onClick={() => selectLedgerMonthYear(`${ledgerPickerYear}-${String(index + 1).padStart(2, '0')}`)}
                    >
                      {month}
                    </button>
                  ))}
                </div>
              </div>}
            </div>
            <button type="button" data-period="custom" className={ledgerPeriod === 'custom' ? 'is-active' : ''} aria-pressed={ledgerPeriod === 'custom'} onClick={() => selectLedgerPeriod('custom')}>تاريخ معين</button>
          </div>
          <div className={`ledger-filter-grid${ledgerPeriod === 'custom' ? ' ledger-filter-grid-custom' : ''}`}>
            <label className="ledger-filter-vehicle"><span>السيارة</span><select value={ledgerFilters.vehicleId} onChange={(event) => setLedgerFilters((current) => ({ ...current, vehicleId: event.target.value }))}><option value="all">كل السيارات</option>{data.vehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</select></label>
            {ledgerPeriod === 'custom' && <>
              <div className="ledger-filter-date-range" role="group" aria-label="نطاق التاريخ">
                <label><span>بداية الفترة</span><input type="date" lang="ar" dir="rtl" value={ledgerFilters.from} onChange={(event) => { setLedgerFilters((current) => ({ ...current, from: event.target.value })); }} /></label>
                <label><span>نهاية الفترة</span><input type="date" lang="ar" dir="rtl" value={ledgerFilters.to} onChange={(event) => { setLedgerFilters((current) => ({ ...current, to: event.target.value })); }} /></label>
              </div>
            </>}
            <label className="ledger-filter-search"><span>بحث في السجل</span><input type="search" value={ledgerFilters.search} onChange={(event) => setLedgerFilters((current) => ({ ...current, search: event.target.value }))} placeholder="السيارة أو اللوحة أو السائق" /></label>
            <button className="ledger-filter-reset" type="button" onClick={resetLedgerFilters}><span aria-hidden="true">↶</span> إعادة تعيين</button>
          </div>
        </section>
        <section className="ledger-summary-block" aria-label="ملخص أداء الأسطول">
          <div className="ledger-summary-heading"><div><span className="fleet-overline">ملخص الأداء</span><h2>صورة سريعة عن الأسطول</h2></div><span>{ledgerEntries.length} حركة مسجلة</span></div>
          <div className="ledger-summary">
            <div><span>السيارات في الكشف</span><strong>{numberFormat.format(new Set(ledgerEntries.map((entry) => entry.vehicleId)).size)}</strong><small>مركبة ضمن الفلتر</small></div>
            <div><span>المسافة الكلية</span><strong>{numberFormat.format(ledgerTotals.distance)} <small>كم</small></strong><small>ضمن الفترة المحددة</small></div>
            <div><span>الوقود المصروف</span><strong>{numberFormat.format(ledgerTotals.liters)} <small>لتر</small></strong><small>إجمالي التعبئة</small></div>
            <div><span>إجمالي المصروف</span><strong>{currencyFormat.format(ledgerTotals.cost)}</strong><small>وقود ومصاريف إضافية</small></div>
            <div><span>انحراف التموين</span><strong className={Math.abs(ledgerTotals.deviation) > 200 ? 'ledger-warning' : 'ledger-ok'}>{currencyFormat.format(ledgerTotals.deviation)}</strong><small>مقارنة بالسعر المعتمد</small></div>
          </div>
        </section>

        <section className="ledger-sheet">
          <div className="ledger-titlebar">
            <div><span>سجل التشغيل والوقود</span><strong>كشف متابعة مصروفات سيارات الشركة</strong></div>
            <span className="ledger-period">{ledgerEntries.length} حركة مسجلة</span>
          </div>
          <div className="fleet-table-wrap">
            <table className="fleet-table ledger-table">
              <thead><tr><th>م</th><th>رقم السيارة</th><th>اسم المركبة</th><th>السائق</th><th>تاريخ التعبئة</th><th>بداية العداد</th><th>قراءة العداد</th><th>الفجوة (كم)</th><th>الوقود المتوقع</th><th>الوقود المصروف</th><th>انحراف التموين (جنيه)</th><th>إجمالي المصروف</th></tr></thead>
              <tbody>{ledgerEntries.map((entry, index) => {
                const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
                const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
                const expected = vehicle ? litersForDistance(vehicle, distance) : 0;
                const liters = litersForEntry(entry, vehicle);
                const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters);
                return <tr key={entry.id} className="fleet-vehicle-tone" style={vehicleToneStyle(entry.vehicleId, data.vehicles)}>
                  <td className="ledger-index">{numberFormat.format(index + 1)}</td>
                  <td className="ledger-plate">{vehicle ? <button type="button" className="ledger-vehicle-link" onClick={() => { setSelectedVehicleId(vehicle.id); setTab('vehicles'); }}>{vehicle.plate}</button> : '—'}</td>
                  <td>{vehicle?.name || 'سيارة محذوفة'}</td>
                  <td>{entry.driver || vehicle?.driver || '—'}</td>
                  <td>{formatDateArabic(entry.date)}</td>
                  <td>{numberFormat.format(entry.startOdometer)}</td>
                  <td>{numberFormat.format(entry.currentOdometer)}</td>
                  <td>{numberFormat.format(distance)} كم</td>
                  <td>{numberFormat.format(expected)} لتر</td>
                  <td>{numberFormat.format(liters)} لتر</td>
                  <td className={Math.abs(deviation) > 200 ? 'ledger-warning' : 'ledger-ok'}>{currencyFormat.format(deviation)}</td>
                  <td>{currencyFormat.format(entryTotalCost(entry))}</td>
                </tr>;
              })}</tbody>
              <tfoot><tr><th colSpan={7}>إجمالي الكشف</th><th>{numberFormat.format(ledgerTotals.distance)} كم</th><th>{numberFormat.format(ledgerTotals.expected)} لتر</th><th>{numberFormat.format(ledgerTotals.liters)} لتر</th><th className={Math.abs(ledgerTotals.deviation) > 200 ? 'ledger-warning' : 'ledger-ok'}>{currencyFormat.format(ledgerTotals.deviation)}</th><th>{currencyFormat.format(ledgerTotals.cost)}</th></tr></tfoot>
            </table>
          </div>
          <div className="ledger-note">الفجوة = القراءة الحالية − القراءة السابقة. انحراف التموين = مصروف الوقود الفعلي − قيمة التموين المعتمدة.</div>
        </section>
      </>}

      {tab === 'vehicles' && <section className={`fleet-section${selectedSummary ? ' fleet-detail-page' : ''}`}>
        <section className="fleet-global-rate fleet-rate-compact" aria-label="سعر التموين الموحد">
          <div><strong>سعر اللتر المعتمد</strong><span>يُطبق تلقائيًا على حسابات جميع السيارات</span></div>
          {!isRateEditorOpen ? <button type="button" className="fleet-rate-value" onClick={() => setIsRateEditorOpen(true)}>{rateFormat.format(fuelAllowanceRate(data))} <small>جنيه / لتر</small><b>تغيير</b></button> : <div className="fleet-rate-editor"><label><input aria-label="سعر التموين المعتمد بالجنيه لكل لتر" type="number" min="0.01" step="0.01" value={allowanceRateDraft} onChange={(event) => setAllowanceRateDraft(event.target.value)} /><span>جنيه / لتر</span></label><button type="button" onClick={saveAllowanceRate}>حفظ</button><button type="button" className="fleet-rate-cancel" onClick={() => setIsRateEditorOpen(false)}>إلغاء</button></div>}
        </section>
        {selectedSummary ? <>
          <div className="fleet-section-heading fleet-detail-heading">
            <div><button type="button" className="fleet-link-button" onClick={() => setSelectedVehicleId(null)}>← العودة إلى السيارات</button><h2>{selectedSummary.vehicle.name}</h2><span className="fleet-overline">{selectedSummary.vehicle.plate} · {selectedSummary.vehicle.driver || 'السائق غير محدد'}</span></div>
            <div className="fleet-detail-actions">
              <button className="fleet-button fleet-button-light" type="button" onClick={() => openVehicleEditor(selectedSummary.vehicle)}>تعديل السيارة</button>
              <button className="fleet-button fleet-button-primary" type="button" onClick={() => startEntry(selectedSummary.vehicle.id)}>+ تسجيل تعبئة</button>
            </div>
          </div>
          <div className="fleet-detail-stats">
            <div><span>بداية العداد</span><strong>{numberFormat.format(selectedSummary.vehicle.openingOdometer)} كم</strong></div>
            <div><span>القراءة الحالية</span><strong>{numberFormat.format(selectedSummary.lastOdometer)} كم</strong></div>
            <div><span>المسافة المقطوعة</span><strong>{numberFormat.format(selectedSummary.distance)} كم</strong></div>
            <div><span>الوقود المتوقع</span><strong>{numberFormat.format(selectedSummary.expected)} لتر</strong></div>
            <div><span>عدد اللترات المصروفة</span><strong>{numberFormat.format(selectedSummary.liters)} لتر</strong></div>
            <div><span>معدل اللترات</span><strong>{numberFormat.format(vehicleFuelRate(selectedSummary.vehicle))} لتر / {numberFormat.format(vehicleDistanceBase(selectedSummary.vehicle))} كم</strong></div>
            <div><span>إجمالي مصاريف السيارة</span><strong>{currencyFormat.format(selectedSummary.cost)}</strong></div>
            <div><span>سعر التموين الموحد</span><strong>{rateFormat.format(fuelAllowanceRate(data))} جنيه / لتر</strong></div>
          </div>
          <div className="fleet-table-wrap fleet-detail-table-wrap">
            <table className="fleet-table ledger-table fleet-detail-ledger">
              <thead><tr><th>التاريخ</th><th>اسم السائق</th><th>خط السير</th><th>القراءة السابقة</th><th>القراءة الحالية</th><th>الفجوة (كم)</th><th>عدد اللترات</th><th>قيمة التموين المقررة</th><th>قيمة التموين الفعلية</th><th>الانحراف</th><th>التراكم</th>{expenseFields.map((field) => <th key={field.key}>{field.label}</th>)}<th>إجمالي مصاريف السيارة</th><th>إجراء</th></tr></thead>
              <tbody>{selectedEntryDetails.map(({ entry, distance, actualFuelValue, deviation, accumulated, totalCost }) => {
                const liters = litersForEntry(entry, selectedSummary.vehicle);
                return <tr key={entry.id} className="fleet-vehicle-tone" style={vehicleToneStyle(entry.vehicleId, data.vehicles)}><td>{formatDateArabic(entry.date)}</td><td>{entry.driver || selectedSummary.vehicle.driver || '—'}</td><td>{entry.route || '—'}</td><td>{numberFormat.format(entry.startOdometer)}</td><td>{numberFormat.format(entry.currentOdometer)}</td><td>{numberFormat.format(distance)} كم</td><td>{numberFormat.format(liters)} لتر</td><td>{currencyFormat.format(allowanceCost(liters, fuelAllowanceRate(data)))}</td><td>{currencyFormat.format(actualFuelValue)}</td><td className={Math.abs(deviation) > 200 ? 'ledger-warning' : 'ledger-ok'}>{currencyFormat.format(deviation)}</td><td className={accumulated > 200 ? 'ledger-warning' : 'ledger-ok'}>{currencyFormat.format(accumulated)}</td>{expenseFields.map((field) => <td key={field.key}>{currencyFormat.format(entry.expenses?.[field.key] || 0)}</td>)}<td>{currencyFormat.format(totalCost)}{entry.note ? <small className="ledger-row-note">{entry.note}</small> : null}</td><td className="ledger-actions"><button type="button" className="entry-edit-button" onClick={() => editEntry(entry)}>تعديل</button><button type="button" className="entry-delete-button" onClick={() => openDeleteConfirmation(entry)}>حذف</button></td></tr>;
              })}</tbody>
              {selectedEntries.length > 0 && <tfoot><tr><th colSpan={5}>إجمالي السيارة</th><th>{numberFormat.format(selectedSummary.distance)} كم</th><th>{numberFormat.format(selectedSummary.liters)} لتر</th><th colSpan={3}>{currencyFormat.format(selectedEntryDetails.reduce((sum, row) => sum + row.deviation, 0))}</th><th>—</th>{expenseFields.map((field) => <th key={field.key}>{currencyFormat.format(selectedEntries.reduce((sum, entry) => sum + (entry.expenses?.[field.key] || 0), 0))}</th>)}<th>{currencyFormat.format(selectedSummary.cost)}</th><th /></tr></tfoot>}
            </table>
            {selectedEntries.length === 0 && <p className="fleet-empty-state">لا توجد مصروفات مسجلة لهذه السيارة بعد.</p>}
          </div>
        </> : <>
          <div className="fleet-section-heading fleet-vehicles-heading">
            <div><span className="fleet-overline">بيانات الأسطول</span><h2>ملخص السيارات</h2></div>
            <div className="fleet-vehicle-actions">
              <button className="fleet-button fleet-button-primary" type="button" onClick={() => setVehicleDialog(true)}>+ إضافة سيارة</button>
            </div>
          </div>
          <div className="fleet-vehicle-grid">{summaries.map((summary) => {
            const { vehicle, lastOdometer, entries, distance, liters, expected, deviation, cost } = summary;
            return <article className="fleet-vehicle fleet-vehicle-card fleet-vehicle-tone" style={vehicleToneStyle(vehicle.id, data.vehicles)} key={vehicle.id} onClick={() => setSelectedVehicleId(vehicle.id)}>
              <div className="fleet-vehicle-top"><span className="fleet-car-icon"><img src="/car-mark.svg" alt="" /></span>{Math.abs(deviation) >= 0.5 && <span className={`fleet-gap ${Math.abs(deviation) > 200 ? 'gap-high' : 'gap-low'}`} aria-label={`انحراف ${currencyFormat.format(deviation)}`}>انحراف: {currencyFormat.format(deviation)}</span>}</div>
              <h3>{vehicle.name}</h3><p className="fleet-plate">{vehicle.plate}</p>
              <div className="fleet-vehicle-data"><span>السائق<strong>{vehicle.driver || 'غير محدد'}</strong></span><span>معدل اللترات<strong>{numberFormat.format(vehicleFuelRate(vehicle))} لتر / {numberFormat.format(vehicleDistanceBase(vehicle))} كم</strong></span><span>العداد الحالي<strong>{numberFormat.format(lastOdometer)} كم</strong></span><span className="fleet-fixed-reading" onClick={(event) => event.stopPropagation()}><label>القراءة السابقة<input type="number" min="0" value={openingOdometerDrafts[vehicle.id] ?? String(vehicle.openingOdometer)} aria-label={`القراءة السابقة لـ ${vehicle.name}`} onChange={(event) => setOpeningOdometerDrafts((current) => ({ ...current, [vehicle.id]: event.target.value }))} /><button type="button" onClick={() => saveOpeningOdometer(vehicle, entries.length)}>حفظ القراءة</button></label></span><span>المسافة المسجلة<strong>{numberFormat.format(distance)} كم</strong></span><span>الوقود المتوقع / المصروف<strong>{numberFormat.format(expected)} / {numberFormat.format(liters)} لتر</strong></span><span>عدد التعبئات<strong>{entries.length}</strong></span><span>إجمالي التكلفة<strong>{currencyFormat.format(cost)}</strong></span></div>
              <div className="fleet-card-actions"><button type="button" className="fleet-open-details" onClick={() => setSelectedVehicleId(vehicle.id)}>عرض تفاصيل السيارة <span aria-hidden="true">←</span></button><button type="button" className="fleet-vehicle-edit" onClick={() => openVehicleEditor(vehicle)}>تعديل السيارة</button></div>
            </article>;
          })}</div>
        </>}
      </section>}

      {tab === 'entries' && <section className="fleet-section fleet-entry-screen">
        <div className="fleet-section-heading entries-page-heading"><div><span className="fleet-overline">حركات الوقود والمصروفات</span><h2>سجل التعبئة</h2></div><span className="entries-count">{filteredEntries.length} حركة ظاهرة</span></div>
        <div className="entries-toolbar">
          <label className="entries-search"><span>بحث في السجل</span><input type="search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="رقم اللوحة أو السيارة أو السائق" /></label>
          <label><span>السيارة</span><select value={ledgerFilters.vehicleId} onChange={(event) => setLedgerFilters((current) => ({ ...current, vehicleId: event.target.value }))}><option value="all">كل السيارات</option>{data.vehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</select></label>
          <label><span>من تاريخ</span><input type="date" value={ledgerFilters.from} onChange={(event) => setLedgerFilters((current) => ({ ...current, from: event.target.value }))} /></label>
          <label><span>إلى تاريخ</span><input type="date" value={ledgerFilters.to} onChange={(event) => setLedgerFilters((current) => ({ ...current, to: event.target.value }))} /></label>
        </div>
        <div className="entries-actions">
          <button className="fleet-button fleet-button-primary" type="button" onClick={() => startEntry()}>+ تسجيل تعبئة</button>
          <button className="fleet-button fleet-button-light" type="button" onClick={() => void exportExcel()}>تصدير Excel</button>
          <button className="fleet-button fleet-button-light" type="button" onClick={() => { setSearch(''); setLedgerFilters({ vehicleId: 'all', from: '', to: '', search: '' }); }}>إعادة ضبط</button>
        </div>
        <div className="fleet-table-wrap"><table className="fleet-table fleet-entry-table"><thead><tr><th>التاريخ</th><th>السيارة</th><th>بداية العداد</th><th>العداد الحالي</th><th>الفجوة (كم)</th><th>الوقود المتوقع</th><th>المصروف</th><th>انحراف التموين (جنيه)</th><th>إجمالي المصروف</th><th>إجراء</th></tr></thead><tbody>
          {filteredEntries.map((entry) => { const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId); const distance = Math.max(0, entry.currentOdometer - entry.startOdometer); const expected = vehicle ? litersForDistance(vehicle, distance) : 0; const liters = litersForEntry(entry, vehicle); const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters); return <tr key={entry.id} className="fleet-vehicle-tone" style={vehicleToneStyle(entry.vehicleId, data.vehicles)}><td>{formatDateArabic(entry.date)}</td><td><strong>{vehicle?.name || 'سيارة محذوفة'}</strong><small>{vehicle?.plate}</small></td><td>{numberFormat.format(entry.startOdometer)}</td><td>{numberFormat.format(entry.currentOdometer)}</td><td>{numberFormat.format(distance)} كم</td><td>{numberFormat.format(expected)} لتر</td><td>{numberFormat.format(liters)} لتر</td><td className={Math.abs(deviation) > 200 ? 'ledger-warning' : 'ledger-ok'}>{currencyFormat.format(deviation)}</td><td>{currencyFormat.format(entryTotalCost(entry))}</td><td className="ledger-actions"><button type="button" className="entry-edit-button" onClick={() => editEntry(entry)}>تعديل</button><button type="button" className="entry-delete-button" onClick={() => openDeleteConfirmation(entry)}>حذف</button></td></tr>; })}
        </tbody></table></div>
      </section>}

      <footer className="fleet-footer"><span>مراقبة سيارات شركة أبو غالي</span></footer>

      {vehicleDialog && <div className="fleet-modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setVehicleDialog(false)}><form className="fleet-modal" onSubmit={saveVehicle}><div className="fleet-modal-title"><div><span className="fleet-overline">بيانات المركبة</span><h2>{editingVehicleId ? 'تعديل بيانات السيارة' : 'إضافة سيارة'}</h2></div><button type="button" aria-label="إغلاق" onClick={() => { setVehicleDialog(false); setEditingVehicleId(null); }}>×</button></div><label>اسم السيارة<input required value={vehicleForm.name} onChange={(event) => setVehicleForm({ ...vehicleForm, name: event.target.value })} placeholder="مثال: تويوتا هايلوكس" /></label><label>رقم اللوحة<input required value={vehicleForm.plate} onChange={(event) => setVehicleForm({ ...vehicleForm, plate: event.target.value })} /></label><label>اسم السائق<input value={vehicleForm.driver} onChange={(event) => setVehicleForm({ ...vehicleForm, driver: event.target.value })} /></label><div className="fleet-form-row"><label>بداية العداد<input required type="number" min="0" value={vehicleForm.openingOdometer} onChange={(event) => setVehicleForm({ ...vehicleForm, openingOdometer: event.target.value })} /></label><label>اللترات المحددة<input required type="number" min="0.1" step="0.1" value={vehicleForm.kmPerLiter} onChange={(event) => setVehicleForm({ ...vehicleForm, kmPerLiter: event.target.value })} /></label></div><div className="fleet-form-row"><label>مسافة الأساس بالكيلومتر<input required type="number" min="1" step="1" value={vehicleForm.distanceBaseKm} onChange={(event) => setVehicleForm({ ...vehicleForm, distanceBaseKm: event.target.value })} /></label><span className="fleet-rate-help">المعادلة: الفجوة ÷ مسافة الأساس × اللترات</span></div><button className="fleet-button fleet-button-primary fleet-submit" type="submit">{editingVehicleId ? 'حفظ التعديل' : 'حفظ السيارة'}</button></form></div>}

      {entryDialog && <div className="fleet-modal-backdrop fleet-entry-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setEntryDialog(false)}>
        <form className="fleet-modal fleet-entry-modal" onSubmit={saveEntry}>
          <header className="entry-form-header">
            <div><span>إدارة الأسطول / سجل المصاريف</span><h2>{editingEntryId ? 'تعديل حركة مسجلة' : 'إدخال حركة جديدة'}</h2></div>
            <button type="button" aria-label="إغلاق" onClick={() => setEntryDialog(false)}>×</button>
          </header>
          <div className="entry-form-body">
            <section className="entry-form-section">
              <div className="entry-section-title"><span>01</span><div><h3>بيانات الرحلة</h3><p>السيارة والتاريخ والسائق والمسار</p></div></div>
              <div className="entry-form-grid entry-trip-grid">
                <label className="entry-field"><span>اختيار السيارة برقم اللوحة</span><input type="search" inputMode="numeric" value={entryPlateSearch} onChange={(event) => setEntryPlateSearch(event.target.value)} placeholder="اكتب رقم اللوحة هنا" /><select required value={entryForm.vehicleId} onChange={(event) => { const next = event.target.value; const vehicleSummary = summaries.find((item) => item.vehicle.id === next); const latest = vehicleSummary?.lastOdometer ?? 0; setEntryPlateSearch(vehicleSummary?.vehicle.plate || ''); setEntryForm({ ...entryForm, vehicleId: next, driver: vehicleSummary?.vehicle.driver || '', startOdometer: String(latest), currentOdometer: '', liters: '0.0' }); }}>{data.vehicles.filter((vehicle) => !entryPlateSearch.trim() || vehicle.plate.includes(entryPlateSearch.trim())).map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</select><small className="entry-field-help">اكتب رقم اللوحة للبحث ثم اختر النتيجة باسم السيارة الكامل</small></label>
                <label className="entry-field"><span>التاريخ</span><input required type="date" value={entryForm.date} onChange={(event) => setEntryForm({ ...entryForm, date: event.target.value })} /></label>
                <label className="entry-field"><span>اسم السائق</span><input value={entryForm.driver} onChange={(event) => setEntryForm({ ...entryForm, driver: event.target.value })} /></label>
                <label className="entry-field"><span>خط السير</span><input value={entryForm.route} onChange={(event) => setEntryForm({ ...entryForm, route: event.target.value })} placeholder="نقطة الانطلاق والوجهة" /></label>
              </div>
            </section>

            <section className="entry-form-section">
              <div className="entry-section-title"><span>02</span><div><h3>العداد والتموين</h3><p>يُحسب فرق العداد تلقائيًا</p></div></div>
              <div className="entry-form-grid entry-fuel-grid">
                <label className="entry-field"><span>القراءة السابقة · آخر قراءة</span><input required type="number" min="0" value={entryForm.startOdometer} readOnly /></label>
                <label className="entry-field"><span>القراءة الحالية</span><input required type="number" min="0" value={entryForm.currentOdometer} onChange={(event) => { const currentOdometer = event.target.value; setEntryForm({ ...entryForm, currentOdometer, liters: calculateFuelLiters(entryForm.vehicleId, Number(entryForm.startOdometer), Number(currentOdometer)).toFixed(1) }); }} /></label>
                <div className="entry-distance"><span>المسافة</span><strong>{numberFormat.format(entryDistance)} <small>كم</small></strong></div>
                <label className="entry-field"><span>التموين الفعلي (لتر)</span><input type="number" min="0" step="0.1" value={entryForm.liters} readOnly /><small className="entry-field-help">يُحسب من فرق العداد ومعدل استهلاك السيارة</small></label>
                <label className="entry-field"><span>قيمة التموين وفقًا للتر (جنيه)</span><input type="number" min="0" step="1" value={calculateAllowanceValue(entryForm.vehicleId, calculateFuelLiters(entryForm.vehicleId, Number(entryForm.startOdometer), Number(entryForm.currentOdometer)))} readOnly /><small className="entry-field-help">{rateFormat.format(fuelAllowanceRate(data))} جنيه لكل لتر لجميع السيارات</small></label>
                <div className="entry-field">
                  <label className="entry-field-label" htmlFor="fuel-amount-input">مصروف الوقود الفعلي (جنيه)</label>
                  <div className="entry-amount-total"><span>الإجمالي</span><output aria-live="polite">{fuelActualValuePreview === null ? '—' : currencyFormat.format(fuelActualValuePreview)}</output></div>
                  <div className="entry-amount-entry"><input id="fuel-amount-input" dir="ltr" type="text" inputMode="decimal" value={fuelAmountDraft} onChange={(event) => setFuelAmountDraft(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); addFuelAmount(); } }} /><button type="button" className="entry-amount-add" onClick={addFuelAmount}>إضافة</button></div>
                  <div className="entry-amount-list" aria-label="المبالغ المضافة لمصروف الوقود">{splitAmountExpression(entryForm.fuelActualValue).map((amount, index) => <span className="entry-amount-chip" key={`${amount}-${index}`}><span dir="ltr">{amount}</span><button type="button" aria-label={`حذف مبلغ ${amount} من الوقود`} onClick={() => removeFuelAmount(index)}>×</button></span>)}</div>
                  {fuelActualValuePreview === null && <small className="entry-field-help" role="alert">راجع المبالغ المضافة</small>}
                </div>
              </div>
            </section>

            <section className="entry-form-section">
              <div className="entry-section-title"><span>03</span><div><h3>المصاريف الإضافية</h3><p>أدخل كل مبلغ واضغط إضافة ليُحسب في الإجمالي</p></div></div>
              <div className="entry-expense-grid">{expenseFields.map((field) => {
                const amount = expenseAmountPreviews[field.key];
                return <div className="entry-field" key={field.key}>
                  <label className="entry-field-label" htmlFor={`expense-amount-${field.key}`}>{field.label}</label>
                  <div className="entry-amount-total"><span>الإجمالي</span><output aria-live="polite">{amount === null ? '—' : currencyFormat.format(amount)}</output></div>
                  <div className="entry-amount-entry"><input id={`expense-amount-${field.key}`} dir="ltr" type="text" inputMode="decimal" aria-label={`${field.label} بالجنيه`} value={expenseAmountDrafts[field.key]} onChange={(event) => setExpenseAmountDrafts((current) => ({ ...current, [field.key]: event.target.value }))} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); addExpenseAmount(field.key); } }} /><button type="button" className="entry-amount-add" onClick={() => addExpenseAmount(field.key)}>إضافة</button></div>
                  <div className="entry-amount-list" aria-label={`المبالغ المضافة: ${field.label}`}>{splitAmountExpression(entryForm.expenses[field.key]).map((part, index) => <span className="entry-amount-chip" key={`${part}-${index}`}><span dir="ltr">{part}</span><button type="button" aria-label={`حذف ${part} من ${field.label}`} onClick={() => removeExpenseAmount(field.key, index)}>×</button></span>)}</div>
                  {amount === null && <small className="entry-field-help" role="alert">راجع المبالغ المضافة</small>}
                </div>;
              })}</div>
            </section>
          </div>
          <footer className="entry-form-actions"><div><span>إجمالي الحركة</span><strong>{hasInvalidAmount ? 'راجع المبالغ' : hasUnaddedAmounts ? 'اضغط إضافة أولًا' : currencyFormat.format(entryCostPreview)}</strong></div><button type="button" className="entry-cancel-button" onClick={() => { setEntryDialog(false); setEditingEntryId(null); }}>إلغاء</button><button className="fleet-button fleet-button-primary entry-save-button" type="submit">{editingEntryId ? 'حفظ التعديل' : 'حفظ الحركة'}</button></footer>
        </form>
      </div>}
      {pendingDeleteEntry && <div className="fleet-confirm-backdrop" role="presentation"><div className="fleet-confirm-dialog" role="alertdialog" aria-modal="true" aria-labelledby="delete-entry-title"><span className="fleet-confirm-icon">!</span><h2 id="delete-entry-title">تأكيد حذف السجل</h2><p>لحذف حركة التعبئة بتاريخ <strong>{formatDateArabic(pendingDeleteEntry.date)}</strong> اكتب كلمة <strong>موافق</strong> ثم اضغط تأكيد.</p><input className="fleet-confirm-input" aria-label="اكتب موافق للتأكيد" value={deleteConfirmationText} onChange={(event) => setDeleteConfirmationText(event.target.value)} autoFocus placeholder="اكتب: موافق" /><div><button type="button" className="fleet-confirm-cancel" onClick={cancelDeleteConfirmation}>إلغاء</button><button type="button" className="fleet-confirm-ok" disabled={deleteConfirmationText.trim() !== 'موافق'} onClick={confirmDeleteEntry}>تأكيد الحذف</button></div></div></div>}
    </main>
  );
};

export default CarFuelPage;