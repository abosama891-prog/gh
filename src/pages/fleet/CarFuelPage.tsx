import React, { useEffect, useMemo, useRef, useState } from 'react';
import type ExcelJS from 'exceljs';
import { Bell, BusFront, Car, CarFront, CarTaxiFront, Check, ChevronDown, Download, Fuel, Home, LayoutDashboard, LocateFixed, LogOut, MapPinned, Pencil, Plus, RefreshCw, RotateCcw, Settings, Trash2, Truck, Upload, Wrench } from 'lucide-react';
import { addGuest, deleteGuest as deleteGuestFromCloud, setGuestActive, subscribeToFleetUsers, updateGuest } from '../../services/fleetUserStore';
import type { FleetGuest, FleetUser } from '../../services/fleetUserStore';
import { isFirebaseConfigured } from '../../services/firebase';
import { readCloudFleetData, readCloudRouteEntries, saveCloudFleetData, saveCloudFleetDataStrict, saveCloudRouteEntries } from '../../services/fleetCloudStore';
import { firebaseAuth } from '../../services/firebase';
import { publishDeviationNotification, subscribeToDeviationNotifications } from '../../services/deviationNotificationStore';
import type { DeviationNotification } from '../../services/deviationNotificationStore';
import { applyUpdate } from '../../services/serviceWorker';
import { clearDeferredPrompt, getDeferredPrompt } from '../../services/pwaInstall';
import RouteMapPicker from './RouteMapPicker';
import { reverseGeocode } from './reverseGeocode';
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

type RouteEntry = {
  id: string;
  date: string;
  vehicleId: string | null;
  vehicleName: string;
  plate: string;
  departure: string;
  route: string;
  representative: string;
  driver: string;
};

type RouteForm = Omit<RouteEntry, 'id' | 'vehicleId'> & { vehicleId: string };

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

type EntryDraft = {
  form: EntryForm;
  fuelAmountDraft: string;
  expenseAmountDrafts: ExpenseInputs;
  entryPlateSearch: string;
};

type FleetData = { vehicles: Vehicle[]; entries: FuelEntry[]; allowancePricePerLiter?: number };
const STORAGE_KEY = 'company-fleet-fuel-v1';
const ROUTE_STORAGE_KEY = 'company-fleet-routes-v1';
const DEFAULT_ALLOWANCE_PRICE = 15.75;
const themePresets = {
  'dark-premium': {
    label: 'داكن فخم',
    emoji: '🌙',
    description: 'خلفية داكنة + ذهبي',
    background: '#0a0f1c',
    surface: '#111827',
    soft: '#1f2937',
    ink: '#f9fafb',
    inkSoft: '#d1d5db',
    inkMuted: '#9ca3af',
    line: '#1f2937',
    accent: '#fbbf24',
    accentDark: '#f59e0b',
    accentSoft: 'rgba(251,191,36,0.12)',
    heading: '#fbbf24',
    fontHeading: "'Playfair Display', 'Cairo', serif",
  },
  'emerald-light': {
    label: 'فاخر دافئ',
    emoji: '✨',
    description: 'عاجي + شامبين + برونزي',
    background: '#f3eee5',
    surface: '#fffdf8',
    soft: '#efe4d2',
    ink: '#2d261f',
    inkSoft: '#594b3e',
    inkMuted: '#867562',
    line: '#deccb0',
    accent: '#a87335',
    accentDark: '#805323',
    accentSoft: 'rgba(168,115,53,0.12)',
    heading: '#5d4025',
    fontHeading: "'Tajawal', 'Cairo', sans-serif",
  },
  'warm-luxury': {
    label: 'برتقالي احترافي',
    emoji: '🚚',
    description: 'برتقالي واضح + فحمي',
    background: '#f5f1ed',
    surface: '#ffffff',
    soft: '#f4e4db',
    ink: '#24211f',
    inkSoft: '#45403d',
    inkMuted: '#77706b',
    line: '#e4c7b8',
    accent: '#f25d2d',
    accentDark: '#d8441a',
    accentSoft: 'rgba(242,93,45,0.12)',
    heading: '#2a1f1b',
    fontHeading: "'Tajawal', 'Cairo', sans-serif",
  },
} as const;
type ThemeKey = keyof typeof themePresets;
const resolveTheme = (storedTheme: string | null): ThemeKey => {
  if (storedTheme === 'dark-premium' || storedTheme === 'emerald-light' || storedTheme === 'warm-luxury') {
    return storedTheme;
  }
  return 'warm-luxury';
};
const themeOptions: ThemeKey[] = ['dark-premium', 'emerald-light', 'warm-luxury'];
const fontSizePresets = { small: 'صغير', medium: 'متوسط', large: 'كبير' } as const;
type FontSizeKey = keyof typeof fontSizePresets;
const today = () => new Date().toISOString().slice(0, 10);
const createEmptyRouteForm = (): RouteForm => ({ date: today(), vehicleId: '', vehicleName: '', plate: '', departure: '', route: '', representative: '', driver: '' });
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

const isRouteEntry = (value: unknown): value is RouteEntry => {
  if (!value || typeof value !== 'object') return false;
  const entry = value as Partial<RouteEntry>;
  return typeof entry.id === 'string'
    && typeof entry.date === 'string'
    && (typeof entry.vehicleId === 'string' || entry.vehicleId === null)
    && typeof entry.vehicleName === 'string'
    && typeof entry.plate === 'string'
    && typeof entry.departure === 'string'
    && typeof entry.route === 'string'
    && typeof entry.representative === 'string'
    && typeof entry.driver === 'string';
};

const readRouteEntries = (): RouteEntry[] => {
  try {
    const stored = localStorage.getItem(ROUTE_STORAGE_KEY);
    if (!stored) return [];
    const parsed: unknown = JSON.parse(stored);
    return Array.isArray(parsed) ? parsed.filter(isRouteEntry) : [];
  } catch (error) {
    console.error('Failed to read daily route records', error);
    return [];
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
const createEmptyEntryForm = (): EntryForm => ({ vehicleId: '', driver: '', startOdometer: '', currentOdometer: '', liters: '', pricePerLiter: '2.18', fuelActualValue: '', allowanceValue: '', date: today(), route: '', note: '', expenses: emptyExpenseInputs() });
const readEntryDraft = (storageKey: string): EntryDraft | null => {
  try {
    const stored = localStorage.getItem(storageKey);
    if (!stored) return null;
    const parsed = JSON.parse(stored) as Partial<EntryDraft>;
    if (!parsed.form || typeof parsed.form !== 'object' || !parsed.form.expenses || typeof parsed.form.expenses !== 'object') return null;
    const defaults = createEmptyEntryForm();
    const form: EntryForm = {
      ...defaults,
      ...parsed.form,
      expenses: { ...defaults.expenses, ...parsed.form.expenses },
    };
    return {
      form,
      fuelAmountDraft: typeof parsed.fuelAmountDraft === 'string' ? parsed.fuelAmountDraft : '',
      expenseAmountDrafts: { ...emptyExpenseInputs(), ...parsed.expenseAmountDrafts },
      entryPlateSearch: typeof parsed.entryPlateSearch === 'string' ? parsed.entryPlateSearch : '',
    };
  } catch (error) {
    console.error('Failed to read the unsaved entry draft', error);
    return null;
  }
};
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

const normalizeTemplateDate = (value: string): string | null => {
  const normalized = value
    .trim()
    .replace(/[٠-٩]/g, (digit) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit)));
  const isoDate = normalized.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s].*)?$/);
  let year: number;
  let month: number;
  let day: number;
  if (isoDate) {
    year = Number(isoDate[1]);
    month = Number(isoDate[2]);
    day = Number(isoDate[3]);
  } else {
    const dayFirstDate = normalized.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if (!dayFirstDate) return null;
    const first = Number(dayFirstDate[1]);
    const second = Number(dayFirstDate[2]);
    year = Number(dayFirstDate[3]);
    if (year < 100) year += year < 50 ? 2000 : 1900;
    if (first > 12) {
      day = first;
      month = second;
    } else if (second > 12) {
      month = first;
      day = second;
    } else {
      day = first;
      month = second;
    }
  }
  if (year < 1 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
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
  if (value instanceof Date) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
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
  { key: 'routes', label: 'GPS', description: 'تسجيل ومتابعة تحركات السيارات', Icon: MapPinned },
] as const;
const fleetHomeActions = [
  { key: 'users', label: 'إدارة الزوار', description: 'إضافة الزوار وتحديث صلاحيات مشاهدتهم', Icon: Settings },
] as const;
type CarFuelPageProps = { onLogout?: () => void; currentUser: FleetUser };
type LedgerPeriod = 'all' | 'month' | 'custom';
type FleetRoute = { tab: 'home' | 'overview' | 'vehicles' | 'entries' | 'routes' | 'users'; selectedVehicleId: string | null; entryPage?: boolean };
type FleetHistoryState = { fleetRoute?: FleetRoute };

const dateInputValue = (date: Date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
const deviationClass = (value: number) => value < 0 ? 'ledger-ok' : value > 250 ? 'ledger-warning' : 'ledger-neutral';
const monthNames = Array.from({ length: 12 }, (_, index) => new Intl.DateTimeFormat('ar-EG', { month: 'long' }).format(new Date(2020, index, 1)));
const CarFuelPage: React.FC<CarFuelPageProps> = ({ onLogout, currentUser }) => {
  const entryDraftStorageKey = `company-fleet-entry-draft-v1-${encodeURIComponent(currentUser.email || currentUser.username)}`;
  const [initialEntryDraft] = useState(() => readEntryDraft(entryDraftStorageKey));
  const entryDraftSuppressedRef = useRef(false);
  const [data, setData] = useState<FleetData>(readData);
  const [routeEntries, setRouteEntries] = useState<RouteEntry[]>(readRouteEntries);
  const [cloudLoaded, setCloudLoaded] = useState(!isFirebaseConfigured);
  const [routesCloudLoaded, setRoutesCloudLoaded] = useState(!isFirebaseConfigured);
  const [routesCloudAvailable, setRoutesCloudAvailable] = useState(!isFirebaseConfigured);
  const [tab, setTab] = useState<'home' | 'overview' | 'vehicles' | 'entries' | 'routes' | 'users'>('home');
  const routeHistoryRef = useRef<FleetRoute | null>(null);
  const skipRoutePushRef = useRef(false);
  const exitConfirmedRef = useRef(false);
  const currentTabRef = useRef(tab);
  currentTabRef.current = tab;
  const [guests, setGuests] = useState<FleetGuest[]>([]);
  const [guestDraft, setGuestDraft] = useState({ password: '', employeeName: '' });
  const [editingGuest, setEditingGuest] = useState<FleetGuest | null>(null);
  const [showGuestDialog, setShowGuestDialog] = useState(false);
  const [isSavingGuest, setIsSavingGuest] = useState(false);
  const themeStorageKey = `cars-site-theme-${currentUser.email || currentUser.username}`;
  const dismissedDeviationNotificationsKey = `cars-site-dismissed-deviation-notifications-${encodeURIComponent(currentUser.email || currentUser.username)}`;
  const [themeKey, setThemeKey] = useState<ThemeKey>(() => resolveTheme(localStorage.getItem(themeStorageKey)));
  const [fontSizeKey, setFontSizeKey] = useState<FontSizeKey>(() => (localStorage.getItem(`${themeStorageKey}-font`) as FontSizeKey) || 'medium');
  const [openingOdometerDrafts, setOpeningOdometerDrafts] = useState<Record<string, string>>({});
  const [isThemePickerOpen, setIsThemePickerOpen] = useState(false);
  const theme = themePresets[themeKey];
  const [selectedVehicleId, setSelectedVehicleId] = useState<string | null>(null);
  const [isDetailSummaryExpanded, setIsDetailSummaryExpanded] = useState(false);
  const [isQuickNavOpen, setIsQuickNavOpen] = useState(false);
  const quickNavRef = useRef<HTMLDivElement>(null);
  const [expandedVehicleCardId, setExpandedVehicleCardId] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [ledgerFilters, setLedgerFilters] = useState({ vehicleId: 'all', from: '', to: '', search: '' });
  const [ledgerPeriod, setLedgerPeriod] = useState<LedgerPeriod>('all');
  const [ledgerMonthYear, setLedgerMonthYear] = useState('');
  const [isLedgerMonthPickerOpen, setIsLedgerMonthPickerOpen] = useState(false);
  const [ledgerPickerYear, setLedgerPickerYear] = useState(new Date().getFullYear());
  const [monthlySummaryMonth, setMonthlySummaryMonth] = useState(() => dateInputValue(new Date()).slice(0, 7));
  const [isMonthlySummaryCollapsed, setIsMonthlySummaryCollapsed] = useState(true);
  const [isLedgerSheetCollapsed, setIsLedgerSheetCollapsed] = useState(true);
  const [isDeviationSummaryCollapsed, setIsDeviationSummaryCollapsed] = useState(true);
  const [isFollowupSummaryCollapsed, setIsFollowupSummaryCollapsed] = useState(true);
  const [deviationSummaryVehicleId, setDeviationSummaryVehicleId] = useState('all');
  const [deviationSummaryMonth, setDeviationSummaryMonth] = useState(() => dateInputValue(new Date()).slice(0, 7));
  const [isDeviationSummaryPickerOpen, setIsDeviationSummaryPickerOpen] = useState(false);
  const [deviationSummaryPickerYear, setDeviationSummaryPickerYear] = useState(new Date().getFullYear());
  const [isMonthlySummaryPickerOpen, setIsMonthlySummaryPickerOpen] = useState(false);
  const [monthlySummaryPickerYear, setMonthlySummaryPickerYear] = useState(new Date().getFullYear());
  const monthlySummaryMonthSelectedRef = useRef(false);
  const deviationSummaryMonthSelectedRef = useRef(false);
  const [vehicleDialog, setVehicleDialog] = useState(false);
  const [editingVehicleId, setEditingVehicleId] = useState<string | null>(null);
  const vehicleImportRef = useRef<HTMLInputElement>(null);
  const allVehiclesImportRef = useRef<HTMLInputElement>(null);
  const [isImportingVehicles, setIsImportingVehicles] = useState(false);
  const [isImportingAllVehicleData, setIsImportingAllVehicleData] = useState(false);
  const vehicleHistoryImportRef = useRef<HTMLInputElement>(null);
  const [isImportingVehicleHistory, setIsImportingVehicleHistory] = useState(false);
  const [entryDialog, setEntryDialog] = useState(false);
  const [maintenanceDialog, setMaintenanceDialog] = useState(false);
  const [maintenanceForm, setMaintenanceForm] = useState({ vehicleId: '', amount: '', amountDraft: '' });
  const currentEntryDialogRef = useRef(entryDialog);
  currentEntryDialogRef.current = entryDialog;
  const [isFuelSectionExpanded, setIsFuelSectionExpanded] = useState(false);
  const [editingEntryId, setEditingEntryId] = useState<string | null>(null);
  const [allowanceRateDraft, setAllowanceRateDraft] = useState(String(DEFAULT_ALLOWANCE_PRICE));
  const [isInstallAvailable, setIsInstallAvailable] = useState(() => Boolean(getDeferredPrompt()));
  const [isAppInstalled, setIsAppInstalled] = useState(() => window.matchMedia('(display-mode: standalone)').matches);
  const [notice, setNotice] = useState('');
  const [noticeIsError, setNoticeIsError] = useState(false);
  const [deviationNotifications, setDeviationNotifications] = useState<DeviationNotification[]>([]);
  const [isDeviationNotificationPanelOpen, setIsDeviationNotificationPanelOpen] = useState(false);
  const dismissedDeviationNotificationsRef = useRef(new Set<string>());
  const pendingDeviationNavigationRef = useRef<DeviationNotification | null>(null);
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [pendingDeleteEntry, setPendingDeleteEntry] = useState<FuelEntry | null>(null);
  const [deleteConfirmationText, setDeleteConfirmationText] = useState('');
  const [isRateEditorOpen, setIsRateEditorOpen] = useState(false);
  const [isRateCardExpanded, setIsRateCardExpanded] = useState(false);
  const [vehicleForm, setVehicleForm] = useState({ name: '', plate: '', driver: '', openingOdometer: '', kmPerLiter: '20', distanceBaseKm: '105' });
  const [entryForm, setEntryForm] = useState<EntryForm>(() => initialEntryDraft?.form ?? createEmptyEntryForm());
  const [routeForm, setRouteForm] = useState<RouteForm>(createEmptyRouteForm);
  const [editingRouteId, setEditingRouteId] = useState<string | null>(null);
  const [locatingRouteField, setLocatingRouteField] = useState<'departure' | 'route' | null>(null);
  const [routeMapField, setRouteMapField] = useState<'departure' | 'route' | null>(null);
  const [routeDateFilter, setRouteDateFilter] = useState(today);
  const [isExportingRoutes, setIsExportingRoutes] = useState(false);
  const [entryPlateSearch, setEntryPlateSearch] = useState(initialEntryDraft?.entryPlateSearch ?? '');
  const [fuelAmountDraft, setFuelAmountDraft] = useState(initialEntryDraft?.fuelAmountDraft ?? '');
  const [expenseAmountDrafts, setExpenseAmountDrafts] = useState<ExpenseInputs>(() => initialEntryDraft?.expenseAmountDrafts ?? emptyExpenseInputs());
  const [expandedExpenseKey, setExpandedExpenseKey] = useState<keyof ExpenseBreakdown | null>(null);
  const canManageUsers = currentUser.role === 'admin' && Boolean(firebaseAuth?.currentUser);
  const canEditFleet = currentUser.role === 'admin' && Boolean(firebaseAuth?.currentUser);

  useEffect(() => {
    const handleUpdateAvailable = () => setUpdateAvailable(true);
    window.addEventListener('sw-update-available', handleUpdateAvailable);
    return () => window.removeEventListener('sw-update-available', handleUpdateAvailable);
  }, []);

  useEffect(() => {
    if (!isQuickNavOpen) return undefined;
    const closeOnOutsideClick = (event: PointerEvent) => {
      if (event.target instanceof Node && !quickNavRef.current?.contains(event.target)) {
        setIsQuickNavOpen(false);
      }
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setIsQuickNavOpen(false);
    };
    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [isQuickNavOpen]);

  useEffect(() => {
    if (!entryDialog || editingEntryId || entryDraftSuppressedRef.current) return;
    const draft: EntryDraft = { form: entryForm, fuelAmountDraft, expenseAmountDrafts, entryPlateSearch };
    try {
      localStorage.setItem(entryDraftStorageKey, JSON.stringify(draft));
    } catch (error) {
      console.error('Failed to save the unsaved entry draft', error);
      setNotice('تعذر حفظ مسودة الحركة على هذا الجهاز');
      setNoticeIsError(true);
    }
  }, [entryDialog, editingEntryId, entryDraftStorageKey, entryForm, entryPlateSearch, expenseAmountDrafts, fuelAmountDraft]);

  useEffect(() => {
    if (!canManageUsers) return undefined;
    return subscribeToFleetUsers(setGuests, (error) => {
      setNotice(error.message);
      setNoticeIsError(true);
    });
  }, [canManageUsers]);

  useEffect(() => {
    const currentRoute: FleetRoute = { tab, selectedVehicleId, entryPage: entryDialog };
    const historyState = window.history.state as FleetHistoryState | null;
    if (!routeHistoryRef.current) {
      window.history.replaceState({ ...historyState, fleetRoute: currentRoute }, '');
      routeHistoryRef.current = currentRoute;
      return;
    }

    const previousRoute = routeHistoryRef.current;
    routeHistoryRef.current = currentRoute;
    if (!previousRoute || (previousRoute.tab === tab && previousRoute.selectedVehicleId === selectedVehicleId && previousRoute.entryPage === entryDialog)) return;
    if (skipRoutePushRef.current) {
      skipRoutePushRef.current = false;
      return;
    }
    window.history.pushState({ fleetRoute: currentRoute }, '');
  }, [entryDialog, selectedVehicleId, tab]);

  useEffect(() => {
    const handlePopState = (event: PopStateEvent) => {
      if (exitConfirmedRef.current) {
        exitConfirmedRef.current = false;
        return;
      }

      const historyState = event.state as FleetHistoryState | null;
      const previousRoute = historyState?.fleetRoute;
      if (currentEntryDialogRef.current) {
        if (previousRoute) {
          skipRoutePushRef.current = true;
          setTab(previousRoute.tab);
          setSelectedVehicleId(previousRoute.selectedVehicleId);
          setEntryDialog(previousRoute.entryPage ?? false);
        } else {
          setEntryDialog(false);
        }
        return;
      }

      if (currentTabRef.current === 'home' && window.matchMedia('(max-width: 768px)').matches) {
        const shouldExit = window.confirm('هل تريد الخروج من التطبيق؟');
        if (!shouldExit) {
          const currentRoute = routeHistoryRef.current;
          if (currentRoute) window.history.pushState({ fleetRoute: currentRoute }, '');
          return;
        }
        exitConfirmedRef.current = true;
        window.history.back();
        return;
      }

      if (!previousRoute) return;

      skipRoutePushRef.current = true;
      setTab(previousRoute.tab);
      setSelectedVehicleId(previousRoute.selectedVehicleId);
      setEntryDialog(previousRoute.entryPage ?? false);
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []);

  useEffect(() => {
    const storedTheme = localStorage.getItem(themeStorageKey) as ThemeKey | null;
    setThemeKey(resolveTheme(storedTheme));
    const storedFontSize = localStorage.getItem(`${themeStorageKey}-font`) as FontSizeKey | null;
    setFontSizeKey(storedFontSize && storedFontSize in fontSizePresets ? storedFontSize : 'medium');
  }, [themeStorageKey]);

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', themeKey);
  }, [themeKey]);

  useEffect(() => {
    if (!cloudLoaded || monthlySummaryMonthSelectedRef.current || data.entries.length === 0) return;
    const latestEntryMonth = data.entries.reduce((latest, entry) => entry.date.slice(0, 7) > latest ? entry.date.slice(0, 7) : latest, '');
    if (!latestEntryMonth) return;
    setMonthlySummaryMonth(latestEntryMonth);
    setMonthlySummaryPickerYear(Number(latestEntryMonth.slice(0, 4)));
  }, [cloudLoaded, data.entries]);

  useEffect(() => {
    if (!cloudLoaded || deviationSummaryMonthSelectedRef.current || data.entries.length === 0) return;
    const latestEntryMonth = data.entries.reduce((latest, entry) => entry.date.slice(0, 7) > latest ? entry.date.slice(0, 7) : latest, '');
    if (!latestEntryMonth) return;
    setDeviationSummaryMonth(latestEntryMonth);
    setDeviationSummaryPickerYear(Number(latestEntryMonth.slice(0, 4)));
  }, [cloudLoaded, data.entries]);

  useEffect(() => {
    if (tab !== 'overview') return;
    setIsFollowupSummaryCollapsed(true);
    setIsLedgerSheetCollapsed(true);
    setIsMonthlySummaryCollapsed(true);
    setIsMonthlySummaryPickerOpen(false);
    setIsDeviationSummaryPickerOpen(false);
    const pendingNotification = pendingDeviationNavigationRef.current;
    if (pendingNotification) {
      setDeviationSummaryMonth(pendingNotification.month);
      setDeviationSummaryPickerYear(Number(pendingNotification.month.slice(0, 4)));
      setDeviationSummaryVehicleId(pendingNotification.vehicleId);
      setIsDeviationSummaryCollapsed(false);
      pendingDeviationNavigationRef.current = null;
    } else {
      setIsDeviationSummaryCollapsed(true);
    }
  }, [tab]);

  const chooseTheme = (nextTheme: ThemeKey) => {
    setThemeKey(nextTheme);
    localStorage.setItem(themeStorageKey, nextTheme);
  };

  const chooseFontSize = (nextFontSize: FontSizeKey) => {
    setFontSizeKey(nextFontSize);
    localStorage.setItem(`${themeStorageKey}-font`, nextFontSize);
  };

  const handleRouteVehicleChange = (vehicleId: string) => {
    if (vehicleId === 'other') {
      setRouteForm((current) => ({ ...current, vehicleId, vehicleName: '', plate: '', driver: '' }));
      return;
    }
    const vehicle = data.vehicles.find((item) => item.id === vehicleId);
    setRouteForm((current) => vehicle
      ? { ...current, vehicleId, vehicleName: vehicle.name, plate: vehicle.plate, driver: vehicle.driver }
      : { ...current, vehicleId, vehicleName: '', plate: '' });
  };

  const useCurrentLocation = (field: 'departure' | 'route') => {
    if (!navigator.geolocation) {
      setNoticeIsError(true);
      setNotice('هذا الجهاز أو المتصفح لا يدعم تحديد الموقع GPS');
      return;
    }
    setLocatingRouteField(field);
    navigator.geolocation.getCurrentPosition(
      ({ coords }) => {
        void reverseGeocode({ latitude: coords.latitude, longitude: coords.longitude })
          .then((address) => {
            setRouteForm((current) => ({ ...current, [field]: address }));
            setNoticeIsError(false);
            setNotice(field === 'departure' ? `تم تحديد نقطة الانطلاق: ${address}` : `تم تحديد الوجهة: ${address}`);
          })
          .catch((error: unknown) => {
            console.error('Could not resolve current GPS location to an address', error);
            setNoticeIsError(true);
            setNotice('تم تحديد إحداثيات GPS، لكن تعذر جلب اسم العنوان. استخدم اختيار الخريطة أو اكتب المكان يدويًا.');
          })
          .finally(() => setLocatingRouteField(null));
      },
      (error) => {
        setLocatingRouteField(null);
        setNoticeIsError(true);
        if (error.code === error.PERMISSION_DENIED) setNotice('لم يتم السماح بالوصول إلى GPS. فعّل إذن الموقع للمتصفح ثم أعد المحاولة.');
        else if (error.code === error.POSITION_UNAVAILABLE) setNotice('تعذر تحديد موقع GPS الحالي. تحقق من تشغيل خدمات الموقع.');
        else if (error.code === error.TIMEOUT) setNotice('استغرق تحديد موقع GPS وقتًا طويلًا. أعد المحاولة.');
        else setNotice('حدث خطأ أثناء تحديد موقع GPS. أعد المحاولة.');
      },
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 60000 },
    );
  };

  const routeMapInitialCoordinate = (field: 'departure' | 'route') => {
    const match = routeForm[field].match(/(?:GPS:\s*)?(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/i);
    if (!match) return undefined;
    const latitude = Number(match[1]);
    const longitude = Number(match[2]);
    if (latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) return undefined;
    return { latitude, longitude };
  };

  const setRouteMapAddress = (field: 'departure' | 'route', address: string) => {
    setRouteForm((current) => ({ ...current, [field]: address }));
    setRouteMapField(null);
    setNoticeIsError(false);
    setNotice(field === 'departure' ? `تم تحديد نقطة الانطلاق: ${address}` : `تم تحديد الوجهة: ${address}`);
  };

  const saveRouteEntry = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!canEditFleet) return;
    const selectedVehicle = data.vehicles.find((vehicle) => vehicle.id === routeForm.vehicleId);
    const isOtherVehicle = routeForm.vehicleId === 'other';
    if (!routeForm.date || !routeForm.vehicleId || (!selectedVehicle && !isOtherVehicle)) {
      setNoticeIsError(true);
      setNotice('اختر سيارة من الأسطول أو حدد سيارة إيجار / أخرى');
      return;
    }
    if (isOtherVehicle && (!routeForm.vehicleName.trim() || !routeForm.plate.trim())) {
      setNoticeIsError(true);
      setNotice('أدخل وصف السيارة ورقمها أو اكتب «إيجار» في خانة الرقم');
      return;
    }
    if (!routeForm.departure.trim() || !routeForm.route.trim() || !routeForm.driver.trim()) {
      setNoticeIsError(true);
      setNotice('أدخل نقطة الانطلاق وخط السير واسم السائق');
      return;
    }
    const savedRoute: RouteEntry = {
      id: editingRouteId || `r${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      date: routeForm.date,
      vehicleId: selectedVehicle?.id ?? null,
      vehicleName: selectedVehicle?.name ?? routeForm.vehicleName.trim(),
      plate: selectedVehicle?.plate ?? routeForm.plate.trim(),
      departure: routeForm.departure.trim(),
      route: routeForm.route.trim(),
      representative: routeForm.representative.trim(),
      driver: routeForm.driver.trim(),
    };
    setRouteEntries((current) => editingRouteId
      ? current.map((entry) => entry.id === editingRouteId ? savedRoute : entry)
      : [savedRoute, ...current]);
    setRouteDateFilter(savedRoute.date);
    setRouteForm(createEmptyRouteForm());
    setEditingRouteId(null);
    setNoticeIsError(false);
    setNotice(editingRouteId ? 'تم تحديث خط السير' : 'تم تسجيل خط السير');
  };

  const editRouteEntry = (entry: RouteEntry) => {
    if (!canEditFleet) return;
    const registeredVehicleId = entry.vehicleId && data.vehicles.some((vehicle) => vehicle.id === entry.vehicleId)
      ? entry.vehicleId
      : null;
    setRouteForm({
      ...entry,
      vehicleId: registeredVehicleId ?? 'other',
    });
    setEditingRouteId(entry.id);
    setRouteDateFilter(entry.date);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const deleteRouteEntry = (entry: RouteEntry) => {
    if (!canEditFleet || !window.confirm(`هل تريد حذف خط سير السيارة ${entry.plate} بتاريخ ${formatDateArabic(entry.date)}؟`)) return;
    setRouteEntries((current) => current.filter((item) => item.id !== entry.id));
    if (editingRouteId === entry.id) {
      setEditingRouteId(null);
      setRouteForm(createEmptyRouteForm());
    }
    setNoticeIsError(false);
    setNotice('تم حذف سجل خط السير');
  };

  const exportRouteEntries = async () => {
    if (!routeEntriesForDate.length) {
      setNoticeIsError(true);
      setNotice('لا توجد سجلات لتصديرها في هذا التاريخ');
      return;
    }
    setIsExportingRoutes(true);
    try {
      const ExcelJS = (await import('exceljs')).default;
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'مراقبة سيارات شركة أبو غالي';
      workbook.subject = 'خط سير السيارات اليومي';
      const worksheet = workbook.addWorksheet(routeDateFilter, { views: [{ rightToLeft: true, showGridLines: false, state: 'frozen', ySplit: 2 }] });
      worksheet.columns = [
        { width: 8 }, { width: 16 }, { width: 19 }, { width: 19 },
        { width: 34 }, { width: 36 }, { width: 24 }, { width: 24 },
      ];
      worksheet.mergeCells('A1:H1');
      worksheet.getCell('A1').value = `خط سير السيارات اليومي - ${formatDateArabic(routeDateFilter)}`;
      worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      worksheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9A682C' } };
      worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(1).height = 32;
      const header = ['م', 'التاريخ', 'الوصف', 'رقم السيارة', 'نقطة الإنطلاق', 'خط السير', 'المندوب', 'السائق'];
      const headerRow = worksheet.addRow(header);
      headerRow.height = 26;
      headerRow.eachCell((cell) => {
        cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5D4025' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      routeEntriesForDate.forEach((entry, index) => {
        const row = worksheet.addRow([index + 1, formatDateArabic(entry.date), entry.vehicleName, entry.plate, entry.departure, entry.route, entry.representative || '—', entry.driver]);
        row.eachCell((cell) => {
          cell.font = { name: 'Arial', size: 11 };
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          if (index % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAF4E9' } };
        });
      });
      const buffer = await workbook.xlsx.writeBuffer();
      const url = URL.createObjectURL(new Blob([buffer as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const link = document.createElement('a');
      link.href = url;
      link.download = `خط-سير-السيارات-${routeDateFilter}.xlsx`;
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setNoticeIsError(false);
      setNotice('تم تصدير كشف خط السير إلى Excel');
    } catch (error) {
      console.error('Failed to export daily route sheet', error);
      setNoticeIsError(true);
      setNotice('تعذر تصدير ملف خط السير. أعد المحاولة.');
    } finally {
      setIsExportingRoutes(false);
    }
  };

  const routeEntriesForDate = routeEntries
    .filter((entry) => entry.date === routeDateFilter)
    .sort((first, second) => first.departure.localeCompare(second.departure, 'ar'));

  const saveOpeningOdometer = (vehicle: Vehicle, entriesCount: number) => {
    if (!canEditFleet) return;
    const value = Number(openingOdometerDrafts[vehicle.id] ?? vehicle.openingOdometer);
    if (!Number.isFinite(value) || value < 0) {
      setNotice('أدخل قراءة سابقة صحيحة');
      return;
    }
    setData((current) => ({ ...current, vehicles: current.vehicles.map((item) => item.id === vehicle.id ? { ...item, openingOdometer: value } : item) }));
    setOpeningOdometerDrafts((current) => ({ ...current, [vehicle.id]: String(value) }));
    setNotice(entriesCount ? `تم حفظ القراءة السابقة لسيارة ${vehicle.name} وتطبيقها كقراءة بداية لاحقة` : `تم حفظ القراءة السابقة وتطبيقها على سيارة ${vehicle.name}`);
  };

  const saveGuest = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!canManageUsers) return;
    setIsSavingGuest(true);
    try {
      if (editingGuest) {
        await updateGuest(editingGuest.id, guestDraft);
        setNotice('تم تحديث بيانات الزائر');
      } else {
        await addGuest(guestDraft);
        setNotice('تمت إضافة الزائر');
      }
      setGuestDraft({ password: '', employeeName: '' });
      setEditingGuest(null);
      setShowGuestDialog(false);
      setNoticeIsError(false);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'تعذر إضافة الزائر');
      setNoticeIsError(true);
    }
    finally {
      setIsSavingGuest(false);
    }
  };

  const toggleGuestActive = async (guest: FleetGuest) => {
    if (!canManageUsers) return;
    try {
      await setGuestActive(guest.id, !guest.active);
      setNotice(guest.active ? 'تم إيقاف الزائر' : 'تم تفعيل الزائر');
      setNoticeIsError(false);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'تعذر تحديث حالة الزائر');
      setNoticeIsError(true);
    }
  };

  const deleteGuest = async (guest: FleetGuest) => {
    if (!canManageUsers) return;
    if (!window.confirm(`هل تريد حذف الزائر ${guest.employeeName}؟`)) return;
    try {
      await deleteGuestFromCloud(guest.id);
      setNotice('تم حذف الزائر');
      setNoticeIsError(false);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'تعذر حذف الزائر');
      setNoticeIsError(true);
    }
  };

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    if (cloudLoaded && canEditFleet) void saveCloudFleetData(data);
  }, [canEditFleet, cloudLoaded, data]);

  useEffect(() => {
    try {
      localStorage.setItem(ROUTE_STORAGE_KEY, JSON.stringify(routeEntries));
    } catch (error) {
      console.error('Failed to save daily route records locally', error);
      setNoticeIsError(true);
      setNotice('تعذر حفظ سجلات خط السير على هذا الجهاز');
    }
    if (isFirebaseConfigured && routesCloudLoaded && routesCloudAvailable && canEditFleet) {
      void saveCloudRouteEntries(routeEntries).catch((error) => {
        console.error('Could not save daily route records to Firestore', error);
        setRoutesCloudAvailable(false);
        setNoticeIsError(true);
        setNotice('تعذر حفظ خطوط السير في السحابة. ستظل محفوظة على هذا الجهاز.');
      });
    }
  }, [canEditFleet, routeEntries, routesCloudAvailable, routesCloudLoaded]);

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
    try {
      const storedIds = JSON.parse(localStorage.getItem(dismissedDeviationNotificationsKey) || '[]') as unknown;
      if (Array.isArray(storedIds)) {
        dismissedDeviationNotificationsRef.current = new Set(storedIds.filter((id): id is string => typeof id === 'string'));
      }
    } catch (error) {
      console.error('Failed to read dismissed deviation notifications', error);
      setNoticeIsError(true);
      setNotice('تعذر تحميل حالة إشعارات الانحرافات المحفوظة على هذا الجهاز.');
    }
    if (!isFirebaseConfigured) return undefined;
    return subscribeToDeviationNotifications((notifications) => {
      const unreadNotifications = notifications.filter((notification) => (
        notification
        && typeof notification.id === 'string'
        && typeof notification.vehicleId === 'string'
        && typeof notification.month === 'string'
        && typeof notification.deviation === 'number'
        && !dismissedDeviationNotificationsRef.current.has(notification.id)
      ));
      if (unreadNotifications.length) {
        setDeviationNotifications((current) => {
          const existingIds = new Set(current.map(({ id }) => id));
          return [...current, ...unreadNotifications.filter(({ id }) => !existingIds.has(id))];
        });
      }
    }, (error) => {
      setNoticeIsError(true);
      setNotice(`تعذر الاتصال بإشعارات الانحرافات للمستخدمين: ${error.message}`);
    });
  }, [dismissedDeviationNotificationsKey]);

  useEffect(() => {
    if (!isFirebaseConfigured) return undefined;
    let cancelled = false;
    readCloudRouteEntries<RouteEntry>().then((cloudRoutes) => {
      if (cancelled) return;
      if (cloudRoutes) setRouteEntries(cloudRoutes.filter(isRouteEntry));
      setRoutesCloudAvailable(true);
      setRoutesCloudLoaded(true);
    }).catch((error) => {
      if (cancelled) return;
      console.error('Could not read daily route records from Firestore', error);
      setRoutesCloudAvailable(false);
      setRoutesCloudLoaded(true);
      setNoticeIsError(true);
      setNotice('تعذر الاتصال بسجل خط السير السحابي. ستظهر السجلات المحفوظة على هذا الجهاز فقط.');
    });
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    setAllowanceRateDraft(String(fuelAllowanceRate(data)));
  }, [data.allowancePricePerLiter]);

  useEffect(() => {
    setIsDetailSummaryExpanded(false);
  }, [selectedVehicleId]);

  useEffect(() => {
    const updateInstallAvailability = () => setIsInstallAvailable(Boolean(getDeferredPrompt()));
    const markInstalled = () => {
      setIsAppInstalled(true);
      setIsInstallAvailable(false);
    };
    const displayMode = window.matchMedia('(display-mode: standalone)');
    const syncDisplayMode = () => {
      if (displayMode.matches) markInstalled();
    };
    window.addEventListener('pwa-install-available', updateInstallAvailability);
    window.addEventListener('pwa-installed', markInstalled);
    displayMode.addEventListener('change', syncDisplayMode);
    return () => {
      window.removeEventListener('pwa-install-available', updateInstallAvailability);
      window.removeEventListener('pwa-installed', markInstalled);
      displayMode.removeEventListener('change', syncDisplayMode);
    };
  }, []);

  const summaries = useMemo(() => data.vehicles.map((vehicle) => {
    const entries = data.entries.filter((entry) => entry.vehicleId === vehicle.id);
    const distance = entries.reduce((total, entry) => total + Math.max(0, entry.currentOdometer - entry.startOdometer), 0);
    const liters = entries.reduce((total, entry) => total + litersForEntry(entry, vehicle), 0);
    const expected = litersForDistance(vehicle, distance);
    const cost = entries.reduce((total, entry) => total + entryTotalCost(entry), 0);
    const deviation = entries.reduce((total, entry) => total + entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)), 0);
    const actualLiters = fuelAllowanceRate(data) > 0
      ? entries.reduce((total, entry) => total + entryFuelCost(entry) / fuelAllowanceRate(data), 0)
      : 0;
    const kmPerLiter = actualLiters > 0 ? distance / actualLiters : 0;
    const lastOdometer = entries.reduce((latest, entry) => Math.max(latest, entry.currentOdometer), vehicle.openingOdometer);
    return { vehicle, entries, distance, liters, expected, deviation, cost, kmPerLiter, lastOdometer };
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
    .sort((a, b) => a.date.localeCompare(b.date));
  const selectedAdditionalExpenses = selectedEntries.reduce((total, entry) => total + sumExpenses(entry.expenses), 0);
  const selectedEntryDetails = selectedEntries.reduce<Array<{
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
    const pricePerLiter = fuelAllowanceRate(data);
    const actualLiters = pricePerLiter > 0
      ? entryFuelCost(entry) / pricePerLiter
      : 0;
    const accumulated = actualLiters > 0
      ? distance / actualLiters
      : 0;
    rows.push({ entry, distance, expectedLiters, actualFuelValue, deviation, accumulated, totalCost: entryTotalCost(entry) });
    return rows;
  }, []);
  const selectedFuelTotals = selectedEntryDetails.reduce((totals, row) => {
    const liters = litersForEntry(row.entry, selectedSummary?.vehicle);
    return {
      allowance: totals.allowance + allowanceCost(liters, fuelAllowanceRate(data)),
      actual: totals.actual + row.actualFuelValue,
      deviation: totals.deviation + row.deviation,
      actualLiters: totals.actualLiters + (fuelAllowanceRate(data) > 0 ? row.actualFuelValue / fuelAllowanceRate(data) : 0),
    };
  }, { allowance: 0, actual: 0, deviation: 0, actualLiters: 0 });
  const totalKmPerLiter = selectedFuelTotals.actualLiters > 0
    ? selectedSummary.distance / selectedFuelTotals.actualLiters
    : 0;

  const filteredEntries = data.entries
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date))
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
    .sort((a, b) => a.date.localeCompare(b.date))
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
  const monthlyVehicleSummaries = useMemo(() => data.vehicles.map((vehicle) => {
    const entries = data.entries.filter((entry) => entry.vehicleId === vehicle.id && entry.date.startsWith(monthlySummaryMonth));
    const expenses = Object.fromEntries(expenseFields.map(({ key }) => [
      key,
      entries.reduce((total, entry) => total + (Number(entry.expenses?.[key]) || 0), 0),
    ])) as ExpenseBreakdown;
    return {
      vehicle,
      entries,
      fuel: entries.reduce((total, entry) => total + entryFuelCost(entry), 0),
      expenses,
      additionalExpenses: sumExpenses(expenses),
      allowance: entries.reduce((total, entry) => total + allowanceCost(litersForEntry(entry, vehicle), fuelAllowanceRate(data)), 0),
      deviation: entries.reduce((total, entry) => total + entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)), 0),
      distance: entries.reduce((total, entry) => total + Math.max(0, entry.currentOdometer - entry.startOdometer), 0),
      liters: entries.reduce((total, entry) => total + litersForEntry(entry, vehicle), 0),
      actualLiters: fuelAllowanceRate(data) > 0
        ? entries.reduce((total, entry) => total + entryFuelCost(entry) / fuelAllowanceRate(data), 0)
        : 0,
    };
  }), [data, monthlySummaryMonth]);
  const monthlySummaryTotals = monthlyVehicleSummaries.reduce((total, summary) => ({
    fuel: total.fuel + summary.fuel,
    expenses: total.expenses + summary.additionalExpenses,
    allowance: total.allowance + summary.allowance,
    deviation: total.deviation + summary.deviation,
    distance: total.distance + summary.distance,
    liters: total.liters + summary.liters,
    actualLiters: total.actualLiters + summary.actualLiters,
  }), { fuel: 0, expenses: 0, allowance: 0, deviation: 0, distance: 0, liters: 0, actualLiters: 0 });
  const deviationSummaries = useMemo(() => data.vehicles.map((vehicle) => {
    const entries = data.entries.filter((entry) => entry.vehicleId === vehicle.id && entry.date.startsWith(deviationSummaryMonth));
    const deviation = entries.reduce((total, entry) => total + entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)), 0);
    return { vehicle, entries, deviation };
  }).filter((summary) => summary.deviation > 250), [data, deviationSummaryMonth]);
  const visibleDeviationSummaries = deviationSummaries.filter((summary) => deviationSummaryVehicleId === 'all' || summary.vehicle.id === deviationSummaryVehicleId);
  const visibleDeviationTotal = visibleDeviationSummaries.reduce((total, summary) => total + summary.deviation, 0);

  const openDeviationNotification = (notification: DeviationNotification) => {
    const dismissedIds = new Set(dismissedDeviationNotificationsRef.current);
    dismissedIds.add(notification.id);
    dismissedDeviationNotificationsRef.current = dismissedIds;
    try {
      localStorage.setItem(dismissedDeviationNotificationsKey, JSON.stringify([...dismissedIds].slice(-200)));
    } catch (error) {
      console.error('Failed to save dismissed deviation notification', error);
      setNoticeIsError(true);
      setNotice('تعذر حفظ حالة الإشعار؛ قد يظهر مجددًا عند إعادة فتح التطبيق.');
    }
    setDeviationNotifications((current) => current.filter(({ id }) => id !== notification.id));
    setDeviationSummaryMonth(notification.month);
    setDeviationSummaryPickerYear(Number(notification.month.slice(0, 4)));
    setDeviationSummaryVehicleId(notification.vehicleId);
    setIsDeviationSummaryPickerOpen(false);
    if (tab === 'overview') {
      setIsFollowupSummaryCollapsed(true);
      setIsLedgerSheetCollapsed(true);
      setIsMonthlySummaryCollapsed(true);
      setIsDeviationSummaryCollapsed(false);
    } else {
      pendingDeviationNavigationRef.current = notification;
      setTab('overview');
    }
  };

  const startEntry = (vehicleId = data.vehicles[0]?.id) => {
    if (!canEditFleet) return;
    const savedDraft = readEntryDraft(entryDraftStorageKey);
    if (savedDraft) {
      entryDraftSuppressedRef.current = false;
      setEditingEntryId(null);
      setEntryForm(savedDraft.form);
      setEntryPlateSearch(savedDraft.entryPlateSearch);
      setFuelAmountDraft(savedDraft.fuelAmountDraft);
      setExpenseAmountDrafts(savedDraft.expenseAmountDrafts);
      setEntryDialog(true);
      setNotice('تم استعادة مسودة الحركة السابقة');
      setNoticeIsError(false);
      return;
    }
    const summary = summaries.find((item) => item.vehicle.id === vehicleId);
    const start = summary?.lastOdometer ?? 0;
    const entryDate = today();
    entryDraftSuppressedRef.current = false;
    setEditingEntryId(null);
    setFuelAmountDraft('');
    setEntryPlateSearch('');
    setExpenseAmountDrafts(emptyExpenseInputs());
    setExpandedExpenseKey(null);
    setEntryForm({ ...createEmptyEntryForm(), vehicleId: vehicleId ?? '', driver: summaries.find((item) => item.vehicle.id === vehicleId)?.vehicle.driver || '', startOdometer: String(start), liters: '0.0', pricePerLiter: '0', date: entryDate });
    setEntryDialog(true);
  };

  const openMaintenanceDialog = () => {
    if (!canEditFleet) return;
    setMaintenanceForm({
      vehicleId: ledgerFilters.vehicleId !== 'all' ? ledgerFilters.vehicleId : data.vehicles[0]?.id ?? '',
      amount: '',
      amountDraft: '',
    });
    setMaintenanceDialog(true);
  };

  const addMaintenanceAmount = () => {
    const amount = parseSingleAmount(maintenanceForm.amountDraft.trim());
    if (amount === null || amount <= 0) {
      setNotice('أدخل مبلغ صيانة صحيحًا أكبر من صفر');
      setNoticeIsError(true);
      return;
    }
    setMaintenanceForm((current) => ({
      ...current,
      amount: [...splitAmountExpression(current.amount), String(amount)].join('+'),
      amountDraft: '',
    }));
    setNotice('');
    setNoticeIsError(false);
  };

  const removeMaintenanceAmount = (indexToRemove: number) => {
    setMaintenanceForm((current) => ({
      ...current,
      amount: splitAmountExpression(current.amount).filter((_, index) => index !== indexToRemove).join('+'),
    }));
  };

  const saveMaintenance = (event: React.FormEvent) => {
    event.preventDefault();
    if (!canEditFleet) return;
    const vehicle = data.vehicles.find((item) => item.id === maintenanceForm.vehicleId);
    const amount = parseAmountExpression(maintenanceForm.amount);
    if (!vehicle) {
      setNotice('اختر سيارة موجودة في قائمة السيارات');
      setNoticeIsError(true);
      return;
    }
    if (maintenanceForm.amountDraft.trim()) {
      setNotice('اضغط «إضافة» لإدخال مبلغ الصيانة قبل الحفظ');
      setNoticeIsError(true);
      return;
    }
    if (amount === null || amount <= 0) {
      setNotice('أدخل مبلغ صيانة صحيحًا أكبر من صفر');
      setNoticeIsError(true);
      return;
    }
    const latestEntry = data.entries
      .filter((entry) => entry.vehicleId === vehicle.id)
      .reduce<FuelEntry | null>((latest, entry) => !latest || entry.date > latest.date ? entry : latest, null);
    if (!latestEntry) {
      setNotice(`لا يوجد سجل سابق لسيارة ${vehicle.name}؛ لم يتم حفظ الصيانة`);
      setNoticeIsError(true);
      setMaintenanceDialog(false);
      return;
    }
    const currentMaintenance = latestEntry.expenses?.maintenance ?? 0;
    const currentExpression = latestEntry.expenseExpressions?.maintenance
      ?? (currentMaintenance ? String(currentMaintenance) : '');
    setData((current) => ({
      ...current,
      entries: current.entries.map((entry) => entry.id === latestEntry.id
        ? {
          ...entry,
          expenses: { ...entry.expenses, maintenance: currentMaintenance + amount },
          expenseExpressions: {
            ...entry.expenseExpressions,
            maintenance: [currentExpression, maintenanceForm.amount].filter(Boolean).join('+'),
          },
        }
        : entry),
    }));
    setMaintenanceDialog(false);
    setNotice(`تمت إضافة الصيانة إلى آخر سجل لسيارة ${vehicle.name}`);
    setNoticeIsError(false);
  };

  const closeEntryPage = () => {
    setEditingEntryId(null);
    const currentRoute = (window.history.state as FleetHistoryState | null)?.fleetRoute;
    if (currentRoute?.entryPage) {
      window.history.back();
      return;
    }
    setEntryDialog(false);
  };

  const editEntry = (entry: FuelEntry) => {
    if (!canEditFleet) return;
    entryDraftSuppressedRef.current = true;
    const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
    setEditingEntryId(entry.id);
    setFuelAmountDraft('');
    setExpenseAmountDrafts(emptyExpenseInputs());
    setExpandedExpenseKey(null);
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
    if (!canEditFleet) return;
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
    if (!canEditFleet) return;
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
    if (!canEditFleet) return;
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
        const openingOdometerText = (row[column('بداية العداد')] || '').trim();
        const kmPerLiterText = (row[column('معدل الاستهلاك (كم/لتر)')] || '').trim();
        const openingOdometer = parseTemplateNumber(openingOdometerText);
        const kmPerLiter = parseTemplateNumber(kmPerLiterText);
        if (!plate && !name && !row.some((value) => value.trim())) return;

        const errors: string[] = [];
        if (!plate) errors.push('رقم اللوحة مطلوب');
        if (!name) errors.push('اسم السيارة مطلوب');
        if (!openingOdometerText || !Number.isFinite(openingOdometer) || openingOdometer < 0) errors.push('بداية العداد يجب أن تكون رقمًا غير سالب');
        if (!kmPerLiterText || !Number.isFinite(kmPerLiter) || kmPerLiter <= 0) errors.push('معدل الاستهلاك يجب أن يكون أكبر من صفر');
        if (errors.length) {
          rowErrors.push(`صف ${index + 2}: ${errors.join('، ')}`);
          return;
        }
        const normalizedPlate = plate.toLocaleLowerCase('ar-EG');
        if (seenPlates.has(normalizedPlate)) {
          rowErrors.push(`صف ${index + 2}: رقم اللوحة مكرر داخل الملف`);
          return;
        }
        seenPlates.add(normalizedPlate);
        imported.push({
          plate,
          name,
          driver,
          openingOdometer,
          kmPerLiter,
          fuelLitersPer105Km: 105 / kmPerLiter,
          fuelDistanceBaseKm: 105,
        });
      });

      if (rowErrors.length) throw new Error(`تعذر استيراد السيارات بسبب الأخطاء التالية:\n${rowErrors.join('\n')}`);
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

  const downloadAllVehicleTemplate = async () => {
    if (data.vehicles.length === 0) {
      setNotice('أضف سيارة واحدة على الأقل قبل تنزيل القالب.');
      setNoticeIsError(true);
      return;
    }
    const ExcelJS = (await import('exceljs')).default;
    const normalizedPlates = data.vehicles.map((vehicle) => vehicle.plate.trim().toLocaleLowerCase('ar-EG'));
    if (new Set(normalizedPlates).size !== normalizedPlates.length) {
      throw new Error('لا يمكن إنشاء القائمة المنسدلة لوجود أرقام لوحات مكررة. صحح أرقام اللوحات أولًا.');
    }
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'مراقبة سيارات شركة أبو غالي';
    workbook.subject = 'قالب بيانات ومصروفات جميع السيارات';
    workbook.calcProperties.fullCalcOnLoad = true;

    const vehicleSheet = workbook.addWorksheet('بيانات السيارات', { views: [{ rightToLeft: true, showGridLines: false, state: 'frozen', ySplit: 1 }] });
    const vehicleHeaders = ['معرف السيارة', 'رقم اللوحة', 'اسم السيارة', 'اسم السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)', 'القراءة السابقة الافتراضية'];
    vehicleSheet.columns = [24, 20, 26, 24, 18, 24, 26].map((width) => ({ width }));
    vehicleSheet.addRow(vehicleHeaders);
    data.vehicles.forEach((vehicle) => {
      const latestReading = summaries.find((summary) => summary.vehicle.id === vehicle.id)?.lastOdometer ?? vehicle.openingOdometer;
      vehicleSheet.addRow([
        vehicle.id, vehicle.plate, vehicle.name, vehicle.driver, vehicle.openingOdometer, vehicle.kmPerLiter, latestReading,
      ]);
    });
    const vehicleHeader = vehicleSheet.getRow(1);
    vehicleHeader.height = 34;
    vehicleHeader.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF176B64' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    vehicleSheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: data.vehicles.length + 1, column: vehicleHeaders.length } };
    workbook.definedNames.add(`'بيانات السيارات'!$B$2:$B$${data.vehicles.length + 1}`, 'VehiclePlates');

    const entryHeaders = [
      'معرف السيارة', 'رقم اللوحة', 'معرف الحركة', 'التاريخ', 'اسم السائق', 'خط السير',
      'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)',
      ...expenseFields.map((field) => `${field.label} (جنيه)`), 'التفصيل',
    ];
    const entryColumnWidths = [24, 20, 25, 16, 22, 24, 18, 18, 24, ...expenseFields.map(() => 18), 28];
    data.vehicles.forEach((vehicle, vehicleIndex) => {
      const entrySheet = workbook.addWorksheet(`سيارة ${vehicleIndex + 1}`, {
        views: [{ rightToLeft: true, showGridLines: false, state: 'frozen', ySplit: 4, xSplit: 3 }],
      });
      entrySheet.columns = entryColumnWidths.map((width) => ({ width }));
      entrySheet.mergeCells(1, 1, 1, entryHeaders.length);
      entrySheet.getCell(1, 1).value = `سجل السيارة - ${vehicle.plate} - ${vehicle.name}`;
      entrySheet.getCell(1, 1).font = { name: 'Arial', size: 15, bold: true, color: { argb: 'FFFFFFFF' } };
      entrySheet.getCell(1, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF176B64' } };
      entrySheet.getCell(1, 1).alignment = { horizontal: 'center', vertical: 'middle' };
      entrySheet.getRow(1).height = 34;
      entrySheet.addRow(['معرف السيارة', vehicle.id, 'رقم اللوحة', vehicle.plate, 'اسم السيارة', vehicle.name]);
      entrySheet.addRow(['معدل الاستهلاك (كم/لتر)', vehicle.kmPerLiter, 'بداية العداد', vehicle.openingOdometer, 'القراءة السابقة الافتراضية', summaries.find((summary) => summary.vehicle.id === vehicle.id)?.lastOdometer ?? vehicle.openingOdometer]);
      entrySheet.mergeCells(3, 7, 3, entryHeaders.length);
      entrySheet.getCell(3, 7).value = 'الورقة خاصة بهذه السيارة. اختر رقم اللوحة في صف جديد، وسيظهر معرّف السيارة والعداد السابق تلقائيًا.';
      entrySheet.getCell(3, 7).font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF52675E' } };
      entrySheet.getCell(3, 7).alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
      entrySheet.getRow(3).height = 28;
      entrySheet.addRow(entryHeaders);
      const entryHeader = entrySheet.getRow(4);
      entryHeader.height = 42;
      entryHeader.eachCell((cell) => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF247C73' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      const vehicleEntries = data.entries
        .filter((entry) => entry.vehicleId === vehicle.id)
        .slice()
        .sort((a, b) => a.date.localeCompare(b.date));
      vehicleEntries.forEach((entry) => {
        entrySheet.addRow([
          vehicle.id, vehicle.plate, entry.id, new Date(`${entry.date}T00:00:00`),
          entry.driver || vehicle.driver, entry.route || '', entry.startOdometer, entry.currentOdometer,
          entryFuelCost(entry), ...expenseFields.map((field) => entry.expenses?.[field.key] || 0), entry.note || '',
        ]);
      });
      for (let index = 0; index < 10; index += 1) entrySheet.addRow([null, vehicle.plate]);
      for (let rowNumber = 5; rowNumber <= entrySheet.rowCount; rowNumber += 1) {
        const row = entrySheet.getRow(rowNumber);
        row.getCell(2).dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['=VehiclePlates'],
          showErrorMessage: true,
          errorTitle: 'رقم لوحة غير موجود',
          error: 'اختر رقم لوحة من السيارات المسجلة في ورقة بيانات السيارات.',
          showInputMessage: true,
          promptTitle: 'اختيار السيارة',
          prompt: 'سيظهر معرّف السيارة والعداد السابق تلقائيًا.',
        };
        if (row.getCell(3).value === null || row.getCell(3).value === undefined) {
          row.getCell(1).value = {
            formula: `IF(B${rowNumber}="","",IFERROR(INDEX('بيانات السيارات'!$A$2:$A$${data.vehicles.length + 1},MATCH(B${rowNumber},'بيانات السيارات'!$B$2:$B$${data.vehicles.length + 1},0)),""))`,
          };
          const previousRowsEnd = rowNumber - 1;
          const previousReadingFormula = previousRowsEnd >= 5
            ? `IFERROR(LOOKUP(2,1/(($B$5:B${previousRowsEnd}=B${rowNumber})*($H$5:H${previousRowsEnd}<>"")),$H$5:H${previousRowsEnd}),VLOOKUP(B${rowNumber},'بيانات السيارات'!$B$2:$G$${data.vehicles.length + 1},6,FALSE))`
            : `VLOOKUP(B${rowNumber},'بيانات السيارات'!$B$2:$G$${data.vehicles.length + 1},6,FALSE)`;
          row.getCell(7).value = { formula: `IF(B${rowNumber}="","",IFERROR(${previousReadingFormula},""))` };
        }
      }
      entrySheet.getColumn(4).numFmt = 'yyyy-mm-dd';
      entrySheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: entrySheet.rowCount, column: entryHeaders.length } };
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buffer as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = 'قالب-جميع-السيارات.xlsx';
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const importAllVehicleTemplate = async (event: React.ChangeEvent<HTMLInputElement>) => {
    if (!canEditFleet) return;
    const input = event.currentTarget;
    const file = input.files?.[0];
    if (!file) return;
    setIsImportingAllVehicleData(true);
    setNoticeIsError(false);
    try {
      if (!/\.xlsx$/i.test(file.name)) throw new Error('قالب جميع السيارات يجب أن يكون بصيغة Excel (.xlsx).');
      const ExcelJS = (await import('exceljs')).default;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
      const vehicleSheet = workbook.getWorksheet('بيانات السيارات');
      if (!vehicleSheet) throw new Error('لم أجد ورقة «بيانات السيارات» في الملف.');
      const vehicleRows = worksheetToRows(vehicleSheet);
      const vehicleHeaders = vehicleRows[0]?.map((value) => value.trim()) || [];
      const requiredVehicleHeaders = ['معرف السيارة', 'رقم اللوحة', 'اسم السيارة', 'اسم السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)'];
      const missingVehicleHeaders = requiredVehicleHeaders.filter((header) => !vehicleHeaders.includes(header));
      if (missingVehicleHeaders.length) throw new Error(`أعمدة بيانات السيارات الناقصة: ${missingVehicleHeaders.join('، ')}`);
      const requiredEntryHeaders = [
        'معرف السيارة', 'رقم اللوحة', 'معرف الحركة', 'التاريخ', 'اسم السائق', 'خط السير',
        'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)', 'التفصيل',
        ...expenseFields.map((field) => `${field.label} (جنيه)`),
      ];

      const vehiclesById = new Map(data.vehicles.map((vehicle) => [vehicle.id, vehicle]));
      const importedVehicles = new Map<string, Vehicle>();
      const vehicleErrors: string[] = [];
      const seenVehicleIds = new Set<string>();
      vehicleRows.slice(1).forEach((row, index) => {
        const value = (header: string) => (row[vehicleHeaders.indexOf(header)] || '').trim();
        const vehicleId = value('معرف السيارة');
        if (!vehicleId && !row.some((cell) => cell.trim())) return;
        const lineNumber = index + 2;
        const existing = vehiclesById.get(vehicleId);
        if (!existing) {
          vehicleErrors.push(`ورقة السيارات، صف ${lineNumber}: معرف السيارة غير موجود`);
          return;
        }
        if (seenVehicleIds.has(vehicleId)) {
          vehicleErrors.push(`ورقة السيارات، صف ${lineNumber}: معرف السيارة مكرر`);
          return;
        }
        seenVehicleIds.add(vehicleId);
        const plate = value('رقم اللوحة');
        const name = value('اسم السيارة');
        const driver = value('اسم السائق');
        const openingText = value('بداية العداد');
        const rateText = value('معدل الاستهلاك (كم/لتر)');
        const openingOdometer = parseTemplateNumber(openingText);
        const kmPerLiter = parseTemplateNumber(rateText);
        const errors: string[] = [];
        if (plate.toLocaleLowerCase('ar-EG') !== existing.plate.trim().toLocaleLowerCase('ar-EG')) errors.push('رقم اللوحة لا يطابق السيارة المرتبطة بالمعرف');
        if (!name) errors.push('اسم السيارة مطلوب');
        if (!openingText || !Number.isFinite(openingOdometer) || openingOdometer < 0) errors.push('بداية العداد غير صحيحة');
        if (!rateText || !Number.isFinite(kmPerLiter) || kmPerLiter <= 0) errors.push('معدل الاستهلاك يجب أن يكون أكبر من صفر');
        if (errors.length) {
          vehicleErrors.push(`ورقة السيارات، صف ${lineNumber}: ${errors.join('، ')}`);
          return;
        }
        importedVehicles.set(vehicleId, {
          ...existing, plate, name, driver, openingOdometer, kmPerLiter,
          fuelLitersPer105Km: 105 / kmPerLiter, fuelDistanceBaseKm: 105,
        });
      });
      if (vehicleErrors.length) throw new Error(`تعذر الاستيراد بسبب أخطاء في بيانات السيارات:\n${vehicleErrors.join('\n')}`);
      const missingVehicleIds = data.vehicles.filter((vehicle) => !importedVehicles.has(vehicle.id)).map((vehicle) => vehicle.plate);
      if (missingVehicleIds.length) throw new Error(`القالب لا يحتوي على بيانات هذه السيارات: ${missingVehicleIds.join('، ')}`);

      const existingEntriesById = new Map(data.entries.map((entry) => [entry.id, entry]));
      const seenEntryIds = new Set<string>();
      const importedEntries: FuelEntry[] = [];
      const entryErrors: string[] = [];
      let skippedRows = 0;
      const vehicleSheets = workbook.worksheets.filter((sheet) => sheet !== vehicleSheet);
      const seenSheetVehicleIds = new Set<string>();
      vehicleSheets.forEach((sheet) => {
        const rows = worksheetToRows(sheet);
        const metadata = rows[1] || [];
        const vehicleId = metadata[1]?.trim() || '';
        const plate = metadata[3]?.trim() || '';
        const vehicle = importedVehicles.get(vehicleId);
        const existingVehicle = vehiclesById.get(vehicleId);
        if (metadata[0]?.trim() !== 'معرف السيارة' || metadata[2]?.trim() !== 'رقم اللوحة'
          || !vehicle || !existingVehicle
          || plate.toLocaleLowerCase('ar-EG') !== existingVehicle.plate.trim().toLocaleLowerCase('ar-EG')) {
          entryErrors.push(`ورقة «${sheet.name}»: بيانات تعريف السيارة غير صحيحة أو لا تطابق ورقة بيانات السيارات`);
          return;
        }
        if (seenSheetVehicleIds.has(vehicleId)) {
          entryErrors.push(`ورقة «${sheet.name}»: يوجد أكثر من ورقة للسيارة نفسها`);
          return;
        }
        seenSheetVehicleIds.add(vehicleId);

        const entryHeaderIndex = rows.findIndex((row) => row.includes('معرف السيارة') && row.includes('معرف الحركة'));
        if (entryHeaderIndex < 0) {
          entryErrors.push(`ورقة «${sheet.name}»: لم أجد عناوين سجل الحركات`);
          return;
        }
        const entryHeaders = rows[entryHeaderIndex].map((value) => value.trim());
        const missingEntryHeaders = requiredEntryHeaders.filter((header) => !entryHeaders.includes(header));
        if (missingEntryHeaders.length) {
          entryErrors.push(`ورقة «${sheet.name}»: الأعمدة الناقصة ${missingEntryHeaders.join('، ')}`);
          return;
        }
        const column = (header: string) => entryHeaders.indexOf(header);
        rows.slice(entryHeaderIndex + 1).forEach((row, index) => {
          const value = (header: string) => (row[column(header)] || '').trim();
          const dateText = value('التاريخ');
          const startText = value('القراءة السابقة');
          const currentText = value('القراءة الحالية');
          const hasMovementData = [
            'معرف الحركة', 'التاريخ', 'اسم السائق', 'خط السير', 'القراءة السابقة',
            'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)', 'التفصيل',
            ...expenseFields.map((field) => `${field.label} (جنيه)`),
          ].some((header) => value(header) !== '');
          if (!hasMovementData) return;
          if (!dateText || !startText || !currentText) {
            skippedRows += 1;
            return;
          }
          const lineNumber = index + entryHeaderIndex + 2;
          const rowVehicleId = value('معرف السيارة');
          const rowVehicle = importedVehicles.get(rowVehicleId);
          const rowPlate = value('رقم اللوحة');
          const errors: string[] = [];
          if (rowVehicleId !== vehicleId || !rowVehicle) errors.push('معرف السيارة لا يطابق هذه الورقة');
          else if (rowPlate.trim().toLocaleLowerCase('ar-EG') !== rowVehicle.plate.trim().toLocaleLowerCase('ar-EG')) errors.push('رقم اللوحة لا يطابق معرف السيارة');
          const date = normalizeTemplateDate(dateText);
          if (!date) errors.push('التاريخ غير صحيح');
          const start = parseTemplateNumber(startText);
          const current = parseTemplateNumber(currentText);
          if (!Number.isFinite(start) || start < 0) errors.push('القراءة السابقة غير صحيحة');
          if (!Number.isFinite(current) || current < 0 || current < start) errors.push('القراءة الحالية يجب أن تكون أكبر من أو تساوي القراءة السابقة');
          const fuelText = value('مصروف الوقود الفعلي (جنيه)');
          const actualFuelValue = fuelText ? parseTemplateNumber(fuelText) : 0;
          if (!Number.isFinite(actualFuelValue) || actualFuelValue < 0) errors.push('مصروف الوقود غير صحيح');
          const expenses = {} as ExpenseBreakdown;
          expenseFields.forEach((field) => {
            const rawAmount = value(`${field.label} (جنيه)`);
            const amount = rawAmount ? parseTemplateNumber(rawAmount) : 0;
            if (!Number.isFinite(amount) || amount < 0) errors.push(`${field.label} غير صحيح`);
            else expenses[field.key] = Math.round(amount);
          });
          const suppliedId = value('معرف الحركة');
          const oldEntry = suppliedId ? existingEntriesById.get(suppliedId) : undefined;
          if (oldEntry && oldEntry.vehicleId !== rowVehicleId) errors.push('معرف الحركة تابع لسيارة أخرى');
          if (suppliedId && seenEntryIds.has(suppliedId)) errors.push('معرف الحركة مكرر داخل الملف');
          if (errors.length) {
            entryErrors.push(`ورقة «${sheet.name}»، صف ${lineNumber}: ${errors.join('، ')}`);
            return;
          }
          if (suppliedId) seenEntryIds.add(suppliedId);
          const id = oldEntry ? suppliedId : `e${Date.now()}-${lineNumber}-${rowVehicleId}`;
          const liters = litersForDistance(rowVehicle, current - start);
          importedEntries.push({
            id, vehicleId: rowVehicleId, date: date || '', startOdometer: start, currentOdometer: current, liters,
            pricePerLiter: liters > 0 ? actualFuelValue / liters : 0, fuelActualValue: actualFuelValue,
            driver: value('اسم السائق'), route: value('خط السير'),
            allowanceValue: allowanceCost(liters, fuelAllowanceRate(data)), expenses, note: value('التفصيل'),
          });
        });
      });
      const missingVehicleSheets = data.vehicles.filter((vehicle) => !seenSheetVehicleIds.has(vehicle.id)).map((vehicle) => vehicle.plate);
      if (missingVehicleSheets.length) entryErrors.push(`القالب لا يحتوي على ورقة مستقلة لهذه السيارات: ${missingVehicleSheets.join('، ')}`);
      if (entryErrors.length) throw new Error(`تعذر الاستيراد بسبب أخطاء في الحركات:\n${entryErrors.join('\n')}`);

      const importedById = new Map(importedEntries.map((entry) => [entry.id, entry]));
      const currentEntryIds = new Set(data.entries.map((entry) => entry.id));
      const entries = data.entries.map((entry) => importedById.get(entry.id) || entry);
      importedEntries.forEach((entry) => { if (!currentEntryIds.has(entry.id)) entries.push(entry); });
      const nextData: FleetData = {
        ...data,
        vehicles: data.vehicles.map((vehicle) => importedVehicles.get(vehicle.id) || vehicle),
        entries,
      };
      const previousLocalData = localStorage.getItem(STORAGE_KEY);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(nextData));
      try {
        await saveCloudFleetDataStrict(nextData);
      } catch (error) {
        if (previousLocalData === null) localStorage.removeItem(STORAGE_KEY);
        else localStorage.setItem(STORAGE_KEY, previousLocalData);
        throw new Error(`فشل حفظ الاستيراد في السحابة ولم تُعتمد التغييرات: ${error instanceof Error ? error.message : 'خطأ غير معروف'}`);
      }
      setData(nextData);
      setNoticeIsError(false);
      setNotice(`تم استيراد ${importedVehicles.size} سيارة و${importedEntries.length} حركة. تم تخطي ${skippedRows} صفوف غير مكتملة. تم ربط كل حركة بمعرف السيارة ورقم لوحتها.`);
    } catch (error) {
      setNoticeIsError(true);
      setNotice(error instanceof Error ? error.message : 'تعذر استيراد ملف جميع السيارات.');
    } finally {
      setIsImportingAllVehicleData(false);
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
      'الانحراف (جنيه - محسوب)', 'كم/لتر',
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

    const movementRows = selectedEntries;
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
    if (!canEditFleet) return;
    const input = event.currentTarget;
    const file = input.files?.[0];
    if (!file || !selectedSummary) return;
    const vehicle = selectedSummary.vehicle;
    setIsImportingVehicleHistory(true);
    try {
      let rows: string[][];
      let vehicleInfoRows: string[][] | undefined;
      if (/\.xlsx$/i.test(file.name)) {
        const ExcelJS = (await import('exceljs')).default;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        const infoSheet = workbook.getWorksheet('بيانات السيارة');
        const worksheet = workbook.getWorksheet('سجل السيارة');
        if (!worksheet) throw new Error('لم أجد ورقة سجل السيارة داخل ملف Excel.');
        rows = worksheetToRows(worksheet);
        if (infoSheet) vehicleInfoRows = worksheetToRows(infoSheet);
      } else {
        rows = parseCsvRows(await file.text());
      }
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
      const usedIds = new Set(data.entries.map((entry) => entry.id));
      const imported: FuelEntry[] = [];
      const errors: string[] = [];
      let skippedRows = 0;
      let importedVehicle = vehicle;

      if (vehicleInfoRows) {
        const info = vehicleInfoRows[2] || [];
        const settings = vehicleInfoRows[3] || [];
        const rowErrors: string[] = [];
        if (info[0]?.trim() !== 'رقم اللوحة' || info[2]?.trim() !== 'اسم السيارة' || info[4]?.trim() !== 'اسم السائق'
          || settings[0]?.trim() !== 'بداية العداد' || settings[2]?.trim() !== 'معدل الاستهلاك (كم/لتر)') {
          rowErrors.push('ورقة بيانات السيارة لا تطابق القالب المعتمد');
        }
        const importedPlate = (info[1] || '').trim();
        const importedName = (info[3] || '').trim();
        const openingOdometerText = (settings[1] || '').trim();
        const kmPerLiterText = (settings[3] || '').trim();
        const openingOdometer = parseTemplateNumber(openingOdometerText);
        const kmPerLiter = parseTemplateNumber(kmPerLiterText);
        if (importedPlate.toLocaleLowerCase('ar-EG') !== vehicle.plate.trim().toLocaleLowerCase('ar-EG')) {
          rowErrors.push(`رقم اللوحة يجب أن يطابق السيارة المحددة (${vehicle.plate})`);
        }
        if (!importedName) rowErrors.push('اسم السيارة مطلوب في ورقة بيانات السيارة');
        if (!openingOdometerText || !Number.isFinite(openingOdometer) || openingOdometer < 0) rowErrors.push('بداية العداد يجب أن تكون رقمًا صحيحًا غير سالب');
        if (!kmPerLiterText || !Number.isFinite(kmPerLiter) || kmPerLiter <= 0) rowErrors.push('معدل الاستهلاك يجب أن يكون أكبر من صفر');
        if (rowErrors.length) throw new Error(`خطأ في بيانات السيارة: ${rowErrors.join('، ')}`);
        importedVehicle = {
          ...vehicle,
          plate: importedPlate,
          name: importedName,
          driver: (info[5] || '').trim(),
          openingOdometer,
          kmPerLiter,
          fuelLitersPer105Km: 105 / kmPerLiter,
          fuelDistanceBaseKm: 105,
        };
      }

      rows.slice(headerRowIndex + 1).forEach((row, index) => {
        const value = (header: string) => (row[column(header)] || '').trim();
        const editableHeaders = [
          'معرف الحركة', 'التاريخ', 'اسم السائق', 'خط السير',
          'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)', 'التفصيل',
          ...expenseFields.map((field) => `${field.label} (جنيه)`),
        ];
        const hasMovement = editableHeaders.some((header) => value(header) !== '');
        if (!hasMovement) return;
        const lineNumber = index + headerRowIndex + 2;
        const rawDate = value('التاريخ');
        const startText = value('القراءة السابقة');
        const currentText = value('القراءة الحالية');
        if (!rawDate || !startText || !currentText) {
          skippedRows += 1;
          return;
        }
        const rowPlate = value('رقم اللوحة').toLocaleLowerCase('ar-EG');
        const start = parseTemplateNumber(startText);
        const current = parseTemplateNumber(currentText);
        const actualFuelText = value('مصروف الوقود الفعلي (جنيه)');
        const actualFuelValue = actualFuelText ? parseTemplateNumber(actualFuelText) : 0;
        const rowErrors: string[] = [];
        if (rowPlate !== vehicle.plate.trim().toLocaleLowerCase('ar-EG')) rowErrors.push('رقم اللوحة لا يطابق السيارة المحددة');
        const date = normalizeTemplateDate(rawDate);
        if (!date) rowErrors.push('التاريخ غير صحيح (استخدم YYYY-MM-DD أو DD/MM/YYYY)');
        if (!Number.isFinite(start) || start < 0) rowErrors.push('القراءة السابقة غير صحيحة');
        if (!Number.isFinite(current) || current < 0 || current < start) rowErrors.push('القراءة الحالية يجب أن تكون أكبر من أو تساوي القراءة السابقة');
        if (!Number.isFinite(actualFuelValue) || actualFuelValue < 0) rowErrors.push('مصروف الوقود الفعلي غير صحيح');
        const expenses = {} as ExpenseBreakdown;
        expenseFields.forEach((field) => {
          const rawAmount = value(`${field.label} (جنيه)`);
          const amount = rawAmount ? parseTemplateNumber(rawAmount) : 0;
          if (!Number.isFinite(amount) || amount < 0) rowErrors.push(`${field.label} غير صحيح`);
          else expenses[field.key] = Math.round(amount);
        });
        if (rowErrors.length) {
          errors.push(`صف ${lineNumber}: ${rowErrors.join('، ')}`);
          return;
        }
        const suppliedId = value('معرف الحركة');
        if (suppliedId && seenIds.has(suppliedId)) {
          errors.push(`صف ${lineNumber}: معرف الحركة مكرر`);
          return;
        }
        if (suppliedId) seenIds.add(suppliedId);
        let id = suppliedId && oldIds.has(suppliedId) ? suppliedId : `e${Date.now()}-${lineNumber}`;
        let uniqueSuffix = 1;
        while (!oldIds.has(id) && usedIds.has(id)) {
          id = `e${Date.now()}-${lineNumber}-${uniqueSuffix}`;
          uniqueSuffix += 1;
        }
        usedIds.add(id);
        const liters = litersForDistance(importedVehicle, current - start);
        imported.push({
          id, vehicleId: vehicle.id, date, startOdometer: start, currentOdometer: current,
          liters, pricePerLiter: liters > 0 ? actualFuelValue / liters : 0, fuelActualValue: actualFuelValue,
          driver: value('اسم السائق'), route: value('خط السير'), allowanceValue: allowanceCost(liters, fuelAllowanceRate(data)),
          expenses, note: value('التفصيل'),
        });
      });

      if (errors.length) throw new Error(`تعذر الاستيراد بسبب أخطاء في البيانات:\n${errors.join('\n')}`);
      if (!imported.length && !vehicleInfoRows) throw new Error(`لم يتم العثور على حركات مكتملة للاستيراد.${skippedRows ? ` تم تخطي ${skippedRows} صفوف غير مكتملة.` : ''}`);
      const updatedCount = imported.filter((entry) => oldIds.has(entry.id)).length;
      const addedCount = imported.length - updatedCount;
      const importedById = new Map(imported.map((entry) => [entry.id, entry]));
      const existingIds = new Set(data.entries.map((entry) => entry.id));
      const entries = data.entries.map((entry) => importedById.get(entry.id) || entry);
      imported.forEach((entry) => { if (!existingIds.has(entry.id)) entries.push(entry); });
      const nextData: FleetData = {
        ...data,
        vehicles: data.vehicles.map((item) => item.id === vehicle.id ? importedVehicle : item),
        entries,
      };
      const previousLocalData = localStorage.getItem(STORAGE_KEY);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(nextData));
      try {
        await saveCloudFleetDataStrict(nextData);
      } catch (error) {
        if (previousLocalData === null) localStorage.removeItem(STORAGE_KEY);
        else localStorage.setItem(STORAGE_KEY, previousLocalData);
        throw new Error(`فشل حفظ الاستيراد في السحابة ولم تُعتمد التغييرات: ${error instanceof Error ? error.message : 'خطأ غير معروف'}`);
      }
      setData(nextData);
      const movementNotice = imported.length
        ? `تحديث ${updatedCount} حركة وإضافة ${addedCount}`
        : 'لا توجد حركات جديدة للاستيراد';
      const skippedNotice = skippedRows ? ` تم تخطي ${skippedRows} صفوف غير مكتملة.` : '';
      setNotice(`تم استيراد بيانات ${vehicle.plate}: ${movementNotice}.${skippedNotice} ظهرت البيانات المكتملة في الملخصات والتقارير وسجل السيارة.`);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'تعذر استيراد سجل السيارة.');
    } finally {
      setIsImportingVehicleHistory(false);
      input.value = '';
    }
  };

  const saveEntry = (event: React.FormEvent) => {
    event.preventDefault();
    if (!canEditFleet) return;
    if (!entryForm.date) {
      setNotice('اختر تاريخ الحركة من التقويم أو أدخله في حقل التاريخ');
      return;
    }
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
    const entryMonth = entry.date.slice(0, 7);
    const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
    const previousMonthlyDeviation = !editingEntryId
      ? data.entries
        .filter((item) => item.vehicleId === entry.vehicleId && item.date.startsWith(entryMonth))
        .reduce((total, item) => total + entryDeviation(item, fuelAllowanceRate(data), litersForEntry(item, vehicle)), 0)
      : 0;
    const nextMonthlyDeviation = previousMonthlyDeviation + entryDeviation(entry, fuelAllowanceRate(data), calculatedLiters);
    const newDeviationNotification = !editingEntryId && previousMonthlyDeviation <= 250 && nextMonthlyDeviation > 250 && vehicle
      ? {
        id: encodeURIComponent(`${entryMonth}_${entry.vehicleId}_${entry.id}`),
        vehicleId: vehicle.id,
        plate: vehicle.plate,
        month: entryMonth,
        deviation: nextMonthlyDeviation,
        createdAt: new Date().toISOString(),
      } satisfies DeviationNotification
      : null;
    if (!editingEntryId) {
      try {
        localStorage.removeItem(entryDraftStorageKey);
      } catch (error) {
        console.error('Failed to clear the saved entry draft', error);
        setNotice('تعذر مسح المسودة المحفوظة. لم يتم حفظ الحركة');
        setNoticeIsError(true);
        return;
      }
    }
    const nextData: FleetData = {
      ...data,
      entries: editingEntryId
        ? data.entries.map((item) => item.id === editingEntryId ? entry : item)
        : [entry, ...data.entries],
    };
    setData(nextData);
    closeEntryPage();
    setNotice(editingEntryId ? 'تم تحديث الحركة' : 'تم حفظ سجل التعبئة وتحديث الفجوة التموينية');
    if (newDeviationNotification) {
      if (!isFirebaseConfigured) {
        setNoticeIsError(true);
        setNotice('تم حفظ الحركة محليًا، لكن لا يمكن إرسال إشعار الانحراف لبقية المستخدمين دون اتصال سحابي.');
      } else {
        void Promise.resolve(saveCloudFleetDataStrict(nextData))
          .then(() => publishDeviationNotification(newDeviationNotification))
          .catch((error: unknown) => {
            console.error('Failed to publish deviation notification', error);
            setNoticeIsError(true);
            setNotice(`تم حفظ الحركة، لكن تعذر إرسال إشعار الانحراف للمستخدمين: ${error instanceof Error ? error.message : 'خطأ غير معروف'}`);
          });
      }
    }
  };

  const confirmDeleteEntry = () => {
    if (!canEditFleet) return;
    if (!pendingDeleteEntry || deleteConfirmationText.trim() !== 'موافق') return;
    setData((current) => ({ ...current, entries: current.entries.filter((entry) => entry.id !== pendingDeleteEntry.id) }));
    setPendingDeleteEntry(null);
    setDeleteConfirmationText('');
    setNotice('تم حذف سجل التعبئة');
  };

  const openDeleteConfirmation = (entry: FuelEntry) => {
    if (!canEditFleet) return;
    setPendingDeleteEntry(entry);
    setDeleteConfirmationText('');
  };

  const cancelDeleteConfirmation = () => {
    setPendingDeleteEntry(null);
    setDeleteConfirmationText('');
  };

  const exportExcel = async () => {
    const rows = [
      ['التاريخ', 'اسم السائق', 'خط السير', 'القراءة السابقة', 'القراءة الحالية', 'الفجوة كم', 'عدد اللترات', 'قيمة التموين وفقًا للتر (جنيه)', 'قيمة التموين الفعلية (جنيه)', 'الانحراف (جنيه)', 'كم/لتر', ...expenseFields.map((field) => `${field.label} (جنيه)`), 'إجمالي مصاريف السيارة (جنيه)', 'رقم اللوحة', 'ملاحظات'],
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
      setIsAppInstalled(true);
      setNotice('التطبيق مثبت بالفعل على هذا الجهاز');
      return;
    }
    const availablePrompt = getDeferredPrompt();
    if (!availablePrompt) {
      setNotice('افتح قائمة Chrome ثم اختر «تثبيت التطبيق» أو «إضافة إلى الشاشة الرئيسية»');
      return;
    }
    try {
      await availablePrompt.prompt();
      const choice = await availablePrompt.userChoice;
      setNotice(choice.outcome === 'accepted' ? 'جاري تثبيت التطبيق...' : 'تم إلغاء التثبيت');
    } catch (error) {
      console.error('Install prompt error', error);
      setNotice('تعذر فتح نافذة التثبيت. أعد المحاولة.');
    } finally {
      clearDeferredPrompt();
      setIsInstallAvailable(false);
    }
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
    if (!canEditFleet) return;
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
  const navigationTiles = fleetSections;
  const openSection = fleetSections.find((section) => section.key === tab)
    || (tab === 'users' ? fleetHomeActions[0] : undefined);

  return (
    <main className={`fleet-app fleet-font-${fontSizeKey}${themeKey === 'dark-premium' ? ' fleet-theme-dark' : ''}${selectedSummary ? ' fleet-app-detail' : ''}${tab === 'home' ? ' fleet-app-home' : ''}`} dir="rtl" style={{ '--teal': theme.accent, '--theme-heading': theme.heading, '--ink': theme.ink, '--muted': theme.inkMuted, '--line': theme.line, '--surface': theme.surface, '--surface-soft': theme.soft, '--app-background': theme.background, '--accent-dark': theme.accentDark, '--accent-soft': theme.accentSoft, '--font-heading': theme.fontHeading } as React.CSSProperties}>
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
      <header className="fleet-header">
        <div className="fleet-quick-nav-anchor" ref={quickNavRef}>
          <button className="fleet-brand fleet-brand-button" type="button" aria-expanded={isQuickNavOpen} aria-haspopup="menu" onClick={() => setIsQuickNavOpen((open) => !open)} aria-label="فتح قائمة الأقسام">
            <span className="fleet-brand-mark"><img src="/icon-speedometer.svg" alt="" /></span>
            <span className="fleet-brand-copy"><strong>Abu Ghali</strong><small>Muraqabat</small></span>
          </button>
          {isQuickNavOpen && <nav className="fleet-quick-nav" aria-label="اختيار القسم" role="menu">
            <button type="button" className={tab === 'home' ? 'is-active' : ''} role="menuitem" onClick={() => { setTab('home'); setSelectedVehicleId(null); setIsQuickNavOpen(false); }}><Home aria-hidden="true" /><span>الرئيسية</span></button>
            {navigationTiles.map(({ key, label, Icon }) => <button key={key} type="button" className={tab === key ? 'is-active' : ''} role="menuitem" onClick={() => { setTab(key); setSelectedVehicleId(null); setIsQuickNavOpen(false); }}><Icon aria-hidden="true" /><span>{label}</span></button>)}
            {canManageUsers && <button type="button" className={tab === 'users' ? 'is-active' : ''} role="menuitem" onClick={() => { setTab('users'); setSelectedVehicleId(null); setIsQuickNavOpen(false); }}><Settings aria-hidden="true" /><span>إدارة الزوار</span></button>}
          </nav>}
        </div>
        <div className="fleet-header-actions">
          <div className="fleet-notification-anchor">
            <button
              className="fleet-header-icon-button fleet-notification-button"
              type="button"
              title="الإشعارات"
              aria-label={deviationNotifications.length ? `الإشعارات، ${deviationNotifications.length} جديد` : 'الإشعارات'}
              aria-expanded={isDeviationNotificationPanelOpen}
              aria-haspopup="dialog"
              onClick={() => setIsDeviationNotificationPanelOpen((open) => !open)}
            >
              <Bell aria-hidden="true" />
              {deviationNotifications.length > 0 && <span className="fleet-notification-count">{deviationNotifications.length > 99 ? '99+' : deviationNotifications.length}</span>}
            </button>
            {isDeviationNotificationPanelOpen && <section className="fleet-notification-panel" role="dialog" aria-label="الإشعارات">
              <div className="fleet-notification-panel-heading">
                <strong>الإشعارات</strong>
                <span>{deviationNotifications.length ? `${deviationNotifications.length} جديد` : 'لا توجد إشعارات جديدة'}</span>
              </div>
              {deviationNotifications.length ? <div className="fleet-notification-list">
                {deviationNotifications.map((notification) => <button
                  key={notification.id}
                  className="fleet-notification-item"
                  type="button"
                  onClick={() => {
                    setIsDeviationNotificationPanelOpen(false);
                    openDeviationNotification(notification);
                  }}
                >
                  <span className="fleet-notification-item-icon"><Bell size={17} aria-hidden="true" /></span>
                  <span className="fleet-notification-item-copy">
                    <strong>انحراف جديد — السيارة {notification.plate}</strong>
                    <span>{notification.month} · {notification.deviation.toLocaleString('ar-EG')} جنيه</span>
                  </span>
                  <span className="fleet-notification-item-action">عرض</span>
                </button>)}
              </div> : <p className="fleet-notification-empty">ستظهر هنا إشعارات الانحرافات الجديدة.</p>}
            </section>}
          </div>
          {currentUser.isGuest && <span className="fleet-user-greeting">مرحباً، {currentUser.guestName}</span>}
          <span className="fleet-live"><i /> متابعة مباشرة</span>
          <button className="fleet-theme-button" type="button" onClick={() => setIsThemePickerOpen((open) => !open)}>ألوان التطبيق</button>
          {canManageUsers && <button className="fleet-header-icon-button fleet-header-settings" type="button" title="إدارة الزوار" aria-label="إدارة الزوار" onClick={() => setTab('users')}><Settings aria-hidden="true" /><span>إدارة الزوار</span></button>}
          <button className="fleet-header-icon-button fleet-header-update" type="button" title="تحديث النظام" aria-label="تحديث النظام" onClick={() => void updateSystem()}><RefreshCw aria-hidden="true" /><span>تحديث النظام</span></button>
          {(isInstallAvailable || isAppInstalled) && <button className="fleet-header-icon-button fleet-header-install" type="button" title={isAppInstalled ? 'التطبيق مثبت' : 'تثبيت التطبيق'} aria-label={isAppInstalled ? 'التطبيق مثبت' : 'تثبيت التطبيق'} onClick={installApp} disabled={isAppInstalled}><Download aria-hidden="true" /><span>{isAppInstalled ? 'مثبت' : 'تثبيت التطبيق'}</span></button>}
          {onLogout && <button className="fleet-header-icon-button fleet-header-logout" type="button" title="تسجيل الخروج" aria-label="تسجيل الخروج" onClick={onLogout}><LogOut aria-hidden="true" /><span>تسجيل الخروج</span></button>}
        </div>
      </header>

      {isThemePickerOpen && <div className="fleet-theme-picker" aria-label="قوالب ألوان التطبيق">
        <div className="fleet-theme-picker-heading">
          <div><strong>اختر القالب</strong><small>خصص ألوان وخطوط لوحة السيارات</small></div>
          <span className="fleet-theme-current">{themePresets[themeKey].emoji} {themePresets[themeKey].label}</span>
        </div>
        <div className="fleet-theme-options">{themeOptions.map((key) => {
          const preset = themePresets[key];
          return <button key={key} type="button" className={`fleet-theme-option${themeKey === key ? ' is-active' : ''}`} aria-pressed={themeKey === key} onClick={() => chooseTheme(key)} style={{ '--preset-accent': preset.accent, '--preset-heading': preset.heading, '--preset-surface': preset.surface, '--preset-background': preset.background } as React.CSSProperties}>
            <span className="fleet-theme-option-heading">
              <span className="fleet-theme-option-name"><i aria-hidden="true">{preset.emoji}</i><strong>{preset.label}</strong></span>
              {themeKey === key && <span className="fleet-theme-active"><Check size={13} aria-hidden="true" /> نشط</span>}
            </span>
            <span className="fleet-theme-option-preview"><span className="fleet-theme-swatches"><i /><i /><i /></span><small>{preset.description}</small></span>
          </button>;
        })}</div>
        <div className="fleet-font-picker">
          <div><strong>حجم الخط</strong><small>اختر الحجم الأنسب للقراءة</small></div>
          <div className="fleet-font-options">{(Object.keys(fontSizePresets) as FontSizeKey[]).map((key) => <button key={key} type="button" className={fontSizeKey === key ? 'is-active' : ''} aria-pressed={fontSizeKey === key} onClick={() => chooseFontSize(key)}>{fontSizePresets[key]}</button>)}</div>
        </div>
      </div>}

      {tab === 'home' ? <nav className="fleet-tabs fleet-tabs-home" aria-label="أقسام متابعة السيارات">
        {navigationTiles.map(({ key, label, description, Icon }) => (
          <button key={key} type="button" className={`fleet-home-section fleet-home-section-${key}`} onClick={() => {
            setTab(key);
            if (key !== 'vehicles') setSelectedVehicleId(null);
          }}>
            <span className="fleet-tab-icon">{key === 'routes' ? <img className="fleet-gps-nav-icon" src="/gps-icon.png" alt="" /> : <Icon aria-hidden="true" strokeWidth={1.8} />}</span>
            <span className="fleet-tab-copy"><strong>{label}</strong></span>
          </button>
        ))}
      </nav> : openSection ? <nav className={`fleet-section-bar fleet-section-bar-${openSection.key}`} aria-label={`قسم ${openSection.label}`}>
        <div className="fleet-section-bar-center" dir="rtl">
          <span className="fleet-section-title-button">{openSection.label}</span>
          <button type="button" className="fleet-section-back-button" onClick={() => { setTab('home'); setSelectedVehicleId(null); }}>رجوع</button>
        </div>
      </nav> : null}
      {notice && <button className={`fleet-notice${noticeIsError ? ' is-error'  : ''}`} type="button" onClick={() => { setNotice(''); setNoticeIsError(false); }}>{notice}<span>×</span></button>}

      {tab === 'users' && canManageUsers && <section className="fleet-section fleet-user-management">
        <div className="fleet-section-heading">
          <div><span className="fleet-overline">إدارة الزوار</span><h2>قائمة الزوار</h2></div>
          <button type="button" className="fleet-button fleet-button-primary" onClick={() => { setGuestDraft({ password: '', employeeName: '' }); setEditingGuest(null); setShowGuestDialog(true); }}>إضافة زائر</button>
          <button type="button" className="fleet-link-button" onClick={() => setTab('home')}>العودة للرئيسية</button>
        </div>
        <div className="fleet-user-table-wrap"><table className="fleet-table fleet-user-table">
          <thead><tr><th>كلمة المرور</th><th>اسم الموظف</th><th>الحالة</th><th>إجراءات</th></tr></thead>
          <tbody>{guests.map((guest) => <tr key={guest.id}>
            <td><strong dir="ltr" aria-label="كلمة المرور مخفية">••••••</strong></td>
            <td><strong>{guest.employeeName}</strong></td>
            <td><button type="button" className={`fleet-user-status${guest.active ? ' is-active' : ''}`} onClick={() => void toggleGuestActive(guest)}>{guest.active ? 'نشط' : 'موقوف'}</button></td>
            <td className="fleet-user-actions">
              <button type="button" className="fleet-row-action fleet-row-action-edit fleet-icon-action" aria-label={`تعديل الزائر ${guest.employeeName}`} title="تعديل الزائر" onClick={() => { setEditingGuest(guest); setGuestDraft({ password: '', employeeName: guest.employeeName }); setShowGuestDialog(true); }}><Pencil size={15} aria-hidden="true" /></button>
              <button type="button" className="fleet-row-action fleet-row-action-delete fleet-icon-action" aria-label={`حذف الزائر ${guest.employeeName}`} title="حذف الزائر" onClick={() => void deleteGuest(guest)}><Trash2 size={15} aria-hidden="true" /></button>
            </td>
          </tr>)}</tbody>
        </table></div>
        {guests.length === 0 && <p className="fleet-empty-state">لا يوجد زوار مسجلون حتى الآن.</p>}
        <p className="fleet-user-storage-note">تُحفظ قائمة الزوار في Firestore وتُحدّث تلقائياً على الأجهزة المتصلة.</p>
        {showGuestDialog && <div className="fleet-modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && !isSavingGuest && setShowGuestDialog(false)}>
          <form className="fleet-modal" onSubmit={saveGuest} role="dialog" aria-modal="true" aria-labelledby="guest-editor-title">
            <div className="fleet-modal-title">
              <div><span className="fleet-overline">بيانات الزائر</span><h2 id="guest-editor-title">{editingGuest ? 'تعديل بيانات الزائر' : 'إضافة زائر'}</h2></div>
              <button type="button" aria-label="إغلاق" disabled={isSavingGuest} onClick={() => setShowGuestDialog(false)}>×</button>
            </div>
            <label>كلمة المرور<input required={!editingGuest || !editingGuest.passwordHash} type="password" autoComplete="new-password" maxLength={72} placeholder={editingGuest?.passwordHash ? 'اتركها فارغة للاحتفاظ بكلمة المرور الحالية' : 'أدخل كلمة المرور'} value={guestDraft.password} onChange={(event) => setGuestDraft({ ...guestDraft, password: event.target.value })} /></label>
            <label>اسم الموظف<input required type="text" autoComplete="off" maxLength={80} value={guestDraft.employeeName} onChange={(event) => setGuestDraft({ ...guestDraft, employeeName: event.target.value })} /></label>
            <div className="fleet-user-edit-actions">
              <button className="fleet-button fleet-button-primary" type="submit" disabled={isSavingGuest}>{isSavingGuest ? 'جارٍ الحفظ...' : editingGuest ? 'حفظ التعديل' : 'حفظ الزائر'}</button>
              <button className="fleet-button fleet-button-light" type="button" disabled={isSavingGuest} onClick={() => setShowGuestDialog(false)}>إلغاء</button>
            </div>
          </form>
        </div>}
      </section>}

      {tab === 'routes' && <section className="fleet-section route-tracking-section">
        <div className="route-section-heading">
          <div><span className="fleet-overline">متابعة التشغيل اليومي</span><h2>خط سير السيارات اليومي</h2><p>سجل مستقل لتحركات السيارات، على نفس ترتيب كشف Excel اليومي.</p></div>
          <div className="route-section-count"><strong>{numberFormat.format(routeEntriesForDate.length)}</strong><span>خط سير في هذا اليوم</span></div>
        </div>
        {!routesCloudLoaded
          ? <div className="route-storage-status" role="status">جاري تحميل سجلات خطوط السير...</div>
          : (!routesCloudAvailable || !isFirebaseConfigured) && <div className="route-storage-status" role="status">{routesCloudAvailable ? 'السجلات محفوظة على هذا الجهاز فقط؛ المزامنة السحابية غير مفعلة.' : 'تعذر الاتصال بسجل خط السير السحابي؛ السجلات الحالية محفوظة على هذا الجهاز فقط.'}</div>}
        {!canEditFleet && <div className="fleet-user-permission-note">يمكنك مشاهدة سجلات خط السير. التعديل والإضافة متاحان للمسؤول فقط.</div>}

        {canEditFleet && <form className="route-entry-form" onSubmit={saveRouteEntry}>
          <div className="route-form-heading"><span className="route-form-icon"><MapPinned aria-hidden="true" /></span><div><h3>{editingRouteId ? 'تعديل خط السير' : 'تسجيل خط سير جديد'}</h3><p>بيانات هذا القسم محفوظة بشكل مستقل عن سجل السيارات والوقود.</p></div></div>
          <div className="route-form-grid">
            <label><span>التاريخ <i>*</i></span><input type="date" lang="ar" dir="rtl" required value={routeForm.date} onChange={(event) => setRouteForm((current) => ({ ...current, date: event.target.value }))} /></label>
            <label><span>رقم السيارة <i>*</i></span><select required value={routeForm.vehicleId} onChange={(event) => handleRouteVehicleChange(event.target.value)}><option value="">اختر سيارة</option>{data.vehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}<option value="other">سيارة إيجار / أخرى</option></select></label>
            {routeForm.vehicleId === 'other' && <>
              <label><span>وصف السيارة <i>*</i></span><input required value={routeForm.vehicleName} onChange={(event) => setRouteForm((current) => ({ ...current, vehicleName: event.target.value }))} placeholder="مثال: جامبو أو ديماكس" /></label>
              <label><span>رقم السيارة <i>*</i></span><input required value={routeForm.plate} onChange={(event) => setRouteForm((current) => ({ ...current, plate: event.target.value }))} placeholder="رقم اللوحة أو إيجار" /></label>
            </>}
            <label><span>نقطة الانطلاق <i>*</i></span><input required value={routeForm.departure} onChange={(event) => setRouteForm((current) => ({ ...current, departure: event.target.value }))} placeholder="اكتب المكان أو وقت التحرك" /><div className="route-location-actions"><button className="route-gps-button" type="button" onClick={() => useCurrentLocation('departure')} disabled={locatingRouteField !== null}><LocateFixed aria-hidden="true" />{locatingRouteField === 'departure' ? 'جارٍ تحديد الموقع...' : 'موقعي الحالي GPS'}</button><button className="route-gps-button" type="button" onClick={() => setRouteMapField('departure')}><MapPinned aria-hidden="true" />اختيار من الخريطة</button></div></label>
            <label><span>خط السير / الوجهة <i>*</i></span><input required value={routeForm.route} onChange={(event) => setRouteForm((current) => ({ ...current, route: event.target.value }))} placeholder="اكتب الوجهة أو تفاصيل خط السير" /><div className="route-location-actions"><button className="route-gps-button" type="button" onClick={() => useCurrentLocation('route')} disabled={locatingRouteField !== null}><LocateFixed aria-hidden="true" />{locatingRouteField === 'route' ? 'جارٍ تحديد الموقع...' : 'موقعي الحالي GPS'}</button><button className="route-gps-button" type="button" onClick={() => setRouteMapField('route')}><MapPinned aria-hidden="true" />اختيار من الخريطة</button></div></label>
            <label><span>المندوب</span><input value={routeForm.representative} onChange={(event) => setRouteForm((current) => ({ ...current, representative: event.target.value }))} placeholder="اسم المندوب" /></label>
            <label><span>السائق <i>*</i></span><input required value={routeForm.driver} onChange={(event) => setRouteForm((current) => ({ ...current, driver: event.target.value }))} placeholder="اسم السائق" /></label>
          </div>
          <div className="route-form-actions"><button className="fleet-button fleet-button-primary" type="submit" disabled={!routesCloudLoaded}><Plus aria-hidden="true" />{editingRouteId ? 'حفظ التعديل' : 'إضافة خط السير'}</button>{editingRouteId && <button className="fleet-button fleet-button-light" type="button" onClick={() => { setEditingRouteId(null); setRouteForm(createEmptyRouteForm()); }}>إلغاء التعديل</button>}</div>
        </form>}

        <section className="route-daily-sheet" aria-label="كشف خطوط السير">
          <div className="route-daily-toolbar">
            <div><span className="fleet-overline">كشف اليوم</span><h3>سجل تحركات السيارات</h3></div>
            <label className="route-date-filter"><span>تاريخ الكشف</span><input type="date" lang="ar" dir="rtl" value={routeDateFilter} onChange={(event) => setRouteDateFilter(event.target.value)} /></label>
            <button className="fleet-button fleet-button-light route-export-button" type="button" onClick={() => void exportRouteEntries()} disabled={isExportingRoutes || routeEntriesForDate.length === 0}><Download aria-hidden="true" />{isExportingRoutes ? 'جاري تجهيز Excel...' : 'تصدير إلى Excel'}</button>
          </div>
          {routeEntriesForDate.length === 0 ? <div className="fleet-empty-state route-empty-state"><MapPinned aria-hidden="true" /><strong>لا توجد خطوط سير مسجلة لهذا اليوم</strong><span>اختر تاريخًا آخر أو أضف أول خط سير من النموذج أعلاه.</span></div> : <div className="fleet-table-wrap route-table-wrap">
            <table className="fleet-table route-table">
              <thead><tr><th>م</th><th>الوصف</th><th>رقم السيارة</th><th>نقطة الانطلاق</th><th>خط السير</th><th>المندوب</th><th>السائق</th>{canEditFleet && <th>إجراءات</th>}</tr></thead>
              <tbody>{routeEntriesForDate.map((entry, index) => <tr key={entry.id}>
                <td className="ledger-index">{numberFormat.format(index + 1)}</td>
                <td><strong>{entry.vehicleName}</strong></td>
                <td className="ledger-plate">{entry.plate}</td>
                <td>{entry.departure}</td>
                <td>{entry.route}</td>
                <td>{entry.representative || '—'}</td>
                <td>{entry.driver}</td>
                {canEditFleet && <td><div className="route-row-actions"><button type="button" aria-label={`تعديل خط سير ${entry.plate}`} title="تعديل" onClick={() => editRouteEntry(entry)}><Pencil aria-hidden="true" /></button><button type="button" aria-label={`حذف خط سير ${entry.plate}`} title="حذف" onClick={() => deleteRouteEntry(entry)}><Trash2 aria-hidden="true" /></button></div></td>}
              </tr>)}</tbody>
            </table>
          </div>}
        </section>
        {routeMapField && <RouteMapPicker initialCoordinate={routeMapInitialCoordinate(routeMapField)} onCancel={() => setRouteMapField(null)} onConfirm={(address) => setRouteMapAddress(routeMapField, address)} />}
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
        <section className={`ledger-summary-block${isFollowupSummaryCollapsed ? ' is-collapsed' : ''}`} aria-label="ملخص أداء كشف المتابعة" onDoubleClick={() => setIsFollowupSummaryCollapsed((collapsed) => !collapsed)}>
          <div className="ledger-summary-heading"><div><span className="fleet-overline">{ledgerMonthYear ? `ملخص الأداء · ${new Intl.DateTimeFormat('ar-EG', { month: 'long', year: 'numeric' }).format(new Date(`${ledgerMonthYear}-01T00:00:00`))}` : 'ملخص الأداء · حسب فلاتر الكشف'}</span><h2><button type="button" className="ledger-summary-toggle" aria-expanded={!isFollowupSummaryCollapsed} aria-controls="ledger-followup-summary-content" onClick={() => setIsFollowupSummaryCollapsed((collapsed) => !collapsed)}>صورة سريعة للمتابعة<ChevronDown aria-hidden="true" /></button></h2></div><span>{numberFormat.format(ledgerEntries.length)} حركة مسجلة</span></div>
          {!isFollowupSummaryCollapsed && <div id="ledger-followup-summary-content" className="ledger-summary">
            <div><span>عدد الحركات المسجلة</span><strong>{numberFormat.format(ledgerEntries.length)}</strong><small>ضمن الفترة والفلاتر المحددة</small></div>
            <div><span>المسافة الكلية</span><strong>{numberFormat.format(ledgerTotals.distance)} <small>كم</small></strong><small>ضمن الفترة المحددة</small></div>
            <div><span>الوقود المصروف</span><strong>{numberFormat.format(ledgerTotals.liters)} <small>لتر</small></strong><small>إجمالي التعبئة</small></div>
            <div><span>إجمالي المصروف</span><strong>{currencyFormat.format(ledgerTotals.cost)}</strong><small>وقود ومصاريف إضافية</small></div>
            <div><span>انحراف التموين</span><strong className={deviationClass(ledgerTotals.deviation)}>{currencyFormat.format(ledgerTotals.deviation)}</strong><small>مقارنة بالسعر المعتمد</small></div>
          </div>}
        </section>

        <section className="ledger-sheet ledger-report-sheet" onDoubleClick={() => setIsLedgerSheetCollapsed((collapsed) => !collapsed)}>
          <div className={`ledger-titlebar ledger-report-titlebar${isLedgerSheetCollapsed ? ' is-collapsed' : ''}`}>
            <div className="ledger-sheet-title">
              <span>سجل التشغيل والوقود</span>
              <button type="button" className="ledger-sheet-toggle" aria-expanded={!isLedgerSheetCollapsed} aria-controls="ledger-sheet-content" onClick={() => setIsLedgerSheetCollapsed((collapsed) => !collapsed)}>
                <strong>كشف متابعة مصروفات سيارات الشركة</strong><ChevronDown aria-hidden="true" />
              </button>
            </div>
            <span className="ledger-period">{ledgerEntries.length} حركة مسجلة</span>
          </div>
          {!isLedgerSheetCollapsed && <div id="ledger-sheet-content">
            <div className="fleet-table-wrap">
              <table className="fleet-table ledger-table">
                <thead><tr><th>م</th><th>رقم السيارة</th><th>السائق</th><th>تاريخ التعبئة</th><th>بداية العداد</th><th>قراءة العداد</th><th>الفجوة (كم)</th><th>الوقود المتوقع</th><th>الوقود المصروف</th><th>انحراف التموين (جنيه)</th><th>إجمالي المصروف</th></tr></thead>
                <tbody>{ledgerEntries.map((entry, index) => {
                  const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
                  const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
                  const expected = vehicle ? litersForDistance(vehicle, distance) : 0;
                  const liters = litersForEntry(entry, vehicle);
                  const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters);
                  return <tr key={entry.id} className="fleet-vehicle-tone" style={vehicleToneStyle(entry.vehicleId, data.vehicles)}>
                    <td className="ledger-index">{numberFormat.format(index + 1)}</td>
                    <td className="ledger-plate">{vehicle ? <button type="button" className="ledger-vehicle-link" onClick={() => { setSelectedVehicleId(vehicle.id); setTab('vehicles'); }}>{vehicle.plate}</button> : '—'}</td>
                    <td>{entry.driver || vehicle?.driver || '—'}</td>
                    <td>{formatDateArabic(entry.date)}</td>
                    <td>{numberFormat.format(entry.startOdometer)}</td>
                    <td>{numberFormat.format(entry.currentOdometer)}</td>
                    <td>{numberFormat.format(distance)} كم</td>
                    <td>{numberFormat.format(expected)} لتر</td>
                    <td>{numberFormat.format(liters)} لتر</td>
                    <td className={deviationClass(deviation)}>{currencyFormat.format(deviation)}</td>
                    <td>{currencyFormat.format(entryTotalCost(entry))}</td>
                  </tr>;
                })}</tbody>
                <tfoot><tr><th colSpan={6}>إجمالي الكشف</th><th>{numberFormat.format(ledgerTotals.distance)} كم</th><th>{numberFormat.format(ledgerTotals.expected)} لتر</th><th>{numberFormat.format(ledgerTotals.liters)} لتر</th><th className={deviationClass(ledgerTotals.deviation)}>{currencyFormat.format(ledgerTotals.deviation)}</th><th>{currencyFormat.format(ledgerTotals.cost)}</th></tr></tfoot>
              </table>
            </div>
            <div className="ledger-note">الفجوة = القراءة الحالية − القراءة السابقة. انحراف التموين = مصروف الوقود الفعلي − قيمة التموين المعتمدة.</div>
          </div>}
        </section>

        <section className="ledger-sheet monthly-summary-sheet" aria-label="الملخص الشهري لمصروفات السيارات" onDoubleClick={() => setIsMonthlySummaryCollapsed((collapsed) => !collapsed)}>
          <div className="ledger-titlebar monthly-summary-titlebar">
            <div className="monthly-summary-title">
              <button type="button" className="monthly-summary-toggle" aria-expanded={!isMonthlySummaryCollapsed} aria-controls="monthly-summary-content" onClick={() => setIsMonthlySummaryCollapsed((collapsed) => !collapsed)}>
                <span>ملخص شهري</span><ChevronDown aria-hidden="true" />
              </button>
              <span>تجميع مصروفات السيارات حسب الشهر</span>
            </div>
            <div className="monthly-summary-controls">
              <span className="ledger-period">{numberFormat.format(monthlyVehicleSummaries.reduce((count, summary) => count + summary.entries.length, 0))} حركة</span>
              {!isMonthlySummaryCollapsed && <div className="ledger-month-picker">
                <button
                  type="button"
                  className="ledger-month-select"
                  aria-label="اختيار شهر الملخص"
                  aria-haspopup="dialog"
                  aria-expanded={isMonthlySummaryPickerOpen}
                  onClick={() => {
                    setMonthlySummaryPickerYear(Number(monthlySummaryMonth.slice(0, 4)));
                    setIsMonthlySummaryPickerOpen((open) => !open);
                  }}
                >
                  {new Intl.DateTimeFormat('ar-EG', { month: 'long', year: 'numeric' }).format(new Date(`${monthlySummaryMonth}-01T00:00:00`))}
                </button>
                {isMonthlySummaryPickerOpen && <div className="ledger-month-picker-panel" role="dialog" aria-label="اختيار شهر الملخص">
                  <div className="ledger-month-picker-header">
                    <button type="button" aria-label="السنة السابقة" onClick={() => setMonthlySummaryPickerYear((year) => year - 1)}>‹</button>
                    <strong>{monthlySummaryPickerYear}</strong>
                    <button type="button" aria-label="السنة التالية" onClick={() => setMonthlySummaryPickerYear((year) => year + 1)}>›</button>
                  </div>
                  <div className="ledger-month-picker-months">
                    {monthNames.map((month, index) => (
                      <button
                        key={month}
                        type="button"
                        className={monthlySummaryMonth === `${monthlySummaryPickerYear}-${String(index + 1).padStart(2, '0')}` ? 'is-selected' : ''}
                        onClick={() => {
                          const selectedMonth = `${monthlySummaryPickerYear}-${String(index + 1).padStart(2, '0')}`;
                          monthlySummaryMonthSelectedRef.current = true;
                          setMonthlySummaryMonth(selectedMonth);
                          setDeviationSummaryVehicleId('all');
                          setIsMonthlySummaryPickerOpen(false);
                        }}
                      >
                        {month}
                      </button>
                    ))}
                  </div>
                </div>}
              </div>}
            </div>
          </div>
          {!isMonthlySummaryCollapsed && <div id="monthly-summary-content">
            <div className="fleet-table-wrap monthly-summary-table-wrap">
              <table className="fleet-table ledger-table monthly-summary-table">
                <thead><tr><th title="رقم السيارة">رقم السيارة</th><th title="إجمالي التموين الفعلي">التموين الفعلي</th><th title="إجمالي المصروفات الإضافية">المصروفات الإضافية</th><th title="إجمالي مصروفات السيارة">إجمالي المصروفات</th><th title="إجمالي قيمة التموين المقررة">التموين المقرر</th><th title="إجمالي الانحراف">الانحراف</th><th title="إجمالي كم/لتر">كم/لتر</th></tr></thead>
                <tbody>{monthlyVehicleSummaries.map((summary) => (
                  <tr key={summary.vehicle.id} className="fleet-vehicle-tone" style={vehicleToneStyle(summary.vehicle.id, data.vehicles)}>
                    <td className="ledger-plate"><button type="button" className="ledger-vehicle-link" onClick={() => { setSelectedVehicleId(summary.vehicle.id); setTab('vehicles'); }}>{summary.vehicle.plate}</button></td>
                    <td>{currencyFormat.format(summary.fuel)}</td>
                    <td>{currencyFormat.format(summary.additionalExpenses)}</td>
                    <td className="monthly-summary-total">{currencyFormat.format(summary.fuel + summary.additionalExpenses)}</td>
                    <td>{currencyFormat.format(Math.round(summary.allowance))}</td>
                    <td className={deviationClass(summary.deviation)}>{currencyFormat.format(Math.round(summary.deviation))}</td>
                    <td>{rateFormat.format(summary.actualLiters > 0 ? summary.distance / summary.actualLiters : 0)}</td>
                  </tr>
                ))}</tbody>
                <tfoot><tr><th>الإجمالي</th><th>{currencyFormat.format(monthlySummaryTotals.fuel)}</th><th>{currencyFormat.format(monthlySummaryTotals.expenses)}</th><th>{currencyFormat.format(monthlySummaryTotals.fuel + monthlySummaryTotals.expenses)}</th><th>{currencyFormat.format(Math.round(monthlySummaryTotals.allowance))}</th><th className={deviationClass(monthlySummaryTotals.deviation)}>{currencyFormat.format(Math.round(monthlySummaryTotals.deviation))}</th><th>{rateFormat.format(monthlySummaryTotals.actualLiters > 0 ? monthlySummaryTotals.distance / monthlySummaryTotals.actualLiters : 0)}</th></tr></tfoot>
              </table>
            </div>
            {monthlyVehicleSummaries.every((summary) => summary.entries.length === 0) && <p className="fleet-empty-state">لا توجد مصروفات مسجلة في هذا الشهر.</p>}
          </div>}
        </section>

        <section className="ledger-sheet deviation-summary-sheet" aria-label="ملخص الانحرافات الشهرية" onDoubleClick={() => setIsDeviationSummaryCollapsed((collapsed) => !collapsed)}>
          <div className={`ledger-titlebar monthly-summary-titlebar deviation-summary-titlebar${isDeviationSummaryCollapsed ? ' is-collapsed' : ''}`}>
            <div className="monthly-summary-title">
              <button type="button" className="monthly-summary-toggle" aria-expanded={!isDeviationSummaryCollapsed} aria-controls="deviation-summary-content" onClick={() => setIsDeviationSummaryCollapsed((collapsed) => !collapsed)}>
                <span>ملخص الانحرافات</span><ChevronDown aria-hidden="true" />
              </button>
              <span>السيارات التي تجاوز انحرافها الإجمالي 250 خلال الشهر المختار</span>
            </div>
            <div className="monthly-summary-controls deviation-summary-controls">
              <span className="ledger-period">{numberFormat.format(deviationSummaries.length)} سيارة</span>
              {!isDeviationSummaryCollapsed && <div className="ledger-month-picker deviation-month-picker">
                <button
                  type="button"
                  className="ledger-month-select"
                  aria-label="اختيار شهر ملخص الانحرافات"
                  aria-haspopup="dialog"
                  aria-expanded={isDeviationSummaryPickerOpen}
                  onClick={() => {
                    setDeviationSummaryPickerYear(Number(deviationSummaryMonth.slice(0, 4)));
                    setIsDeviationSummaryPickerOpen((open) => !open);
                  }}
                >
                  {new Intl.DateTimeFormat('ar-EG', { month: 'long', year: 'numeric' }).format(new Date(`${deviationSummaryMonth}-01T00:00:00`))}
                </button>
                {isDeviationSummaryPickerOpen && <div className="ledger-month-picker-panel" role="dialog" aria-label="اختيار شهر ملخص الانحرافات">
                  <div className="ledger-month-picker-header">
                    <button type="button" aria-label="السنة السابقة لملخص الانحرافات" onClick={() => setDeviationSummaryPickerYear((year) => year - 1)}>‹</button>
                    <strong>{deviationSummaryPickerYear}</strong>
                    <button type="button" aria-label="السنة التالية لملخص الانحرافات" onClick={() => setDeviationSummaryPickerYear((year) => year + 1)}>›</button>
                  </div>
                  <div className="ledger-month-picker-months">
                    {monthNames.map((month, index) => (
                      <button
                        key={month}
                        type="button"
                        className={deviationSummaryMonth === `${deviationSummaryPickerYear}-${String(index + 1).padStart(2, '0')}` ? 'is-selected' : ''}
                        onClick={() => {
                          const selectedMonth = `${deviationSummaryPickerYear}-${String(index + 1).padStart(2, '0')}`;
                          deviationSummaryMonthSelectedRef.current = true;
                          setDeviationSummaryMonth(selectedMonth);
                          setDeviationSummaryVehicleId('all');
                          setIsDeviationSummaryPickerOpen(false);
                        }}
                      >
                        {month}
                      </button>
                    ))}
                  </div>
                </div>}
              </div>}
              {!isDeviationSummaryCollapsed && <label className="deviation-summary-filter"><span>السيارة</span>
                <select value={deviationSummaryVehicleId} onChange={(event) => setDeviationSummaryVehicleId(event.target.value)} aria-label="تصفية ملخص الانحرافات حسب السيارة">
                  <option value="all">كل السيارات المتجاوزة</option>
                  {deviationSummaries.map(({ vehicle }) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate}</option>)}
                </select>
              </label>}
            </div>
          </div>
          {!isDeviationSummaryCollapsed && <div id="deviation-summary-content">
            {visibleDeviationSummaries.length > 0 ? <div className="fleet-table-wrap deviation-summary-table-wrap">
              <table className="fleet-table ledger-table deviation-summary-table">
                <thead><tr><th>رقم السيارة</th><th>عدد الحركات</th><th>إجمالي الانحراف</th></tr></thead>
                <tbody>{visibleDeviationSummaries.map((summary) => (
                  <tr key={summary.vehicle.id} className="fleet-vehicle-tone" style={vehicleToneStyle(summary.vehicle.id, data.vehicles)}>
                    <td className="ledger-plate"><button type="button" className="ledger-vehicle-link" onClick={() => { setSelectedVehicleId(summary.vehicle.id); setTab('vehicles'); }}>{summary.vehicle.plate}</button></td>
                    <td>{numberFormat.format(summary.entries.length)}</td>
                    <td className="ledger-warning deviation-summary-amount">{currencyFormat.format(Math.round(summary.deviation))}</td>
                  </tr>
                ))}</tbody>
                <tfoot><tr><th colSpan={2}>إجمالي الانحراف للسيارات المعروضة</th><th className="ledger-warning">{currencyFormat.format(Math.round(visibleDeviationTotal))}</th></tr></tfoot>
              </table>
            </div> : <p className="fleet-empty-state">{deviationSummaries.length === 0 ? 'لا توجد سيارات تجاوز انحرافها الإجمالي 250 في هذا الشهر.' : 'لا توجد نتائج للسيارة المحددة.'}</p>}
          </div>}
        </section>
      </>}

      {tab === 'vehicles' && <section className={`fleet-section${selectedSummary ? ' fleet-detail-page' : ''}`}>
        <section className={`fleet-global-rate fleet-rate-compact${selectedSummary ? ` fleet-rate-collapsible${isRateCardExpanded ? ' is-expanded' : ' is-collapsed'}` : ''}`} aria-label="سعر التموين الموحد">
          {selectedSummary && <button type="button" className="fleet-rate-collapse-toggle" aria-expanded={isRateCardExpanded} onClick={() => setIsRateCardExpanded((expanded) => !expanded)}>
            <span>سعر اللتر المعتمد</span><strong>{rateFormat.format(fuelAllowanceRate(data))} <small>جنيه / لتر</small></strong><ChevronDown aria-hidden="true" />
          </button>}
          {(!selectedSummary || isRateCardExpanded) && <>
            <div><strong>سعر اللتر المعتمد</strong><span>يُطبق تلقائيًا على حسابات جميع السيارات</span></div>
            {!isRateEditorOpen ? <button type="button" className="fleet-rate-value" onClick={() => setIsRateEditorOpen(true)}>{rateFormat.format(fuelAllowanceRate(data))} <small>جنيه / لتر</small><b>تغيير</b></button> : <div className="fleet-rate-editor"><label><input aria-label="سعر التموين المعتمد بالجنيه لكل لتر" type="number" min="0.01" step="0.01" value={allowanceRateDraft} onChange={(event) => setAllowanceRateDraft(event.target.value)} /><span>جنيه / لتر</span></label><button type="button" onClick={saveAllowanceRate}>حفظ</button><button type="button" className="fleet-rate-cancel" onClick={() => setIsRateEditorOpen(false)}>إلغاء</button></div>}
          </>}
        </section>
        {selectedSummary ? <>
          <div className="fleet-section-heading fleet-detail-heading">
            <div><button type="button" className="fleet-link-button" onClick={() => setSelectedVehicleId(null)}>← العودة إلى السيارات</button><h2>{selectedSummary.vehicle.name}</h2><span className="fleet-overline">{selectedSummary.vehicle.plate} · {selectedSummary.vehicle.driver || 'السائق غير محدد'}</span></div>
            <div className="fleet-detail-actions">
              <button className="fleet-button fleet-button-light" type="button" onClick={() => void downloadSelectedVehicleTemplate().catch((error: unknown) => setNotice(error instanceof Error ? `تعذر إنشاء قالب السيارة: ${error.message}` : 'تعذر إنشاء قالب السيارة.'))}><Download size={14} aria-hidden="true" /> تنزيل قالب السيارة</button>
              {canEditFleet && <>
                <button className="fleet-button fleet-button-light" type="button" disabled={isImportingVehicleHistory} onClick={() => vehicleHistoryImportRef.current?.click()}><Upload size={14} aria-hidden="true" /> {isImportingVehicleHistory ? 'جاري الاستيراد...' : 'استيراد بيانات السيارة'}</button>
                <input ref={vehicleHistoryImportRef} className="fleet-file-input" type="file" accept=".xlsx,.csv" aria-label="اختر قالب السيارة للاستيراد" onChange={importSelectedVehicleTemplate} />
              </>}
              <button className="fleet-button fleet-button-light fleet-icon-action" type="button" aria-label={`تعديل السيارة ${selectedSummary.vehicle.name}`} title="تعديل السيارة" onClick={() => openVehicleEditor(selectedSummary.vehicle)}><Pencil size={16} aria-hidden="true" /></button>
              <button className="fleet-button fleet-button-primary" type="button" onClick={() => startEntry(selectedSummary.vehicle.id)}>+ تسجيل تعبئة</button>
            </div>
          </div>
          <p className="fleet-template-note">القالب خاص بالسيارة المحددة: عدّل بياناتها وحركاتها ثم استورده لتحديث الملخصات والتقارير وسجل السيارة.</p>
          <section className={`fleet-detail-summary${isDetailSummaryExpanded ? ' is-expanded' : ' is-collapsed'}`} aria-label="ملخص السيارة">
            <button type="button" className="fleet-detail-summary-toggle" aria-expanded={isDetailSummaryExpanded} onClick={() => setIsDetailSummaryExpanded((expanded) => !expanded)}>
              <span><small>إجمالي التموين الفعلي</small><strong>{currencyFormat.format(Math.round(selectedFuelTotals.actual))}</strong></span>
              <span><small>إجمالي الانحراف</small><strong className={`fleet-gap${selectedFuelTotals.deviation > 250 ? ' gap-high' : selectedFuelTotals.deviation < 0 ? ' gap-low' : ' gap-normal'}`}>{currencyFormat.format(Math.round(selectedFuelTotals.deviation))}</strong></span>
              <ChevronDown aria-hidden="true" />
            </button>
            {isDetailSummaryExpanded && <div className="fleet-detail-stats">
            <div><span>بداية العداد</span><strong>{numberFormat.format(selectedSummary.vehicle.openingOdometer)} كم</strong></div>
            <div><span>القراءة الحالية</span><strong>{numberFormat.format(selectedSummary.lastOdometer)} كم</strong></div>
            <div><span>المسافة المقطوعة</span><strong>{numberFormat.format(selectedSummary.distance)} كم</strong></div>
            <div><span>الوقود المتوقع</span><strong>{numberFormat.format(selectedSummary.expected)} لتر</strong></div>
            <div><span>عدد اللترات المصروفة</span><strong>{numberFormat.format(selectedSummary.liters)} لتر</strong></div>
            <div><span>معدل اللترات</span><strong>{numberFormat.format(vehicleFuelRate(selectedSummary.vehicle))} لتر / {numberFormat.format(vehicleDistanceBase(selectedSummary.vehicle))} كم</strong></div>
            <div><span>إجمالي مصاريف السيارة</span><strong>{currencyFormat.format(selectedSummary.cost)}</strong></div>
            <div className="fleet-extra-expense-stat"><span>إجمالي المصاريف الإضافية</span><strong>{currencyFormat.format(selectedAdditionalExpenses)}</strong></div>
            <div><span>سعر التموين الموحد</span><strong>{rateFormat.format(fuelAllowanceRate(data))} جنيه / لتر</strong></div>
            </div>}
          </section>
          <div className="fleet-table-wrap fleet-detail-table-wrap">
            <table className="fleet-table ledger-table fleet-detail-ledger">
              <thead><tr><th>التاريخ</th><th>اسم السائق</th><th>خط السير</th><th>القراءة السابقة</th><th>القراءة الحالية</th><th>الفجوة (كم)</th><th>عدد اللترات</th><th>قيمة التموين المقررة</th><th>قيمة التموين الفعلية</th><th>الانحراف</th><th>كم/لتر</th>{expenseFields.map((field) => <th key={field.key}>{field.label}</th>)}<th>إجمالي مصاريف السيارة</th><th>إجراء</th></tr></thead>
              <tbody>{selectedEntryDetails.map(({ entry, distance, actualFuelValue, deviation, accumulated, totalCost }) => {
                const liters = litersForEntry(entry, selectedSummary.vehicle);
                return <tr key={entry.id} className="fleet-vehicle-tone" style={vehicleToneStyle(entry.vehicleId, data.vehicles)}><td>{formatDateArabic(entry.date)}</td><td>{entry.driver || selectedSummary.vehicle.driver || '—'}</td><td>{entry.route || '—'}</td><td>{numberFormat.format(entry.startOdometer)}</td><td>{numberFormat.format(entry.currentOdometer)}</td><td>{numberFormat.format(distance)} كم</td><td>{numberFormat.format(liters)} لتر</td><td>{numberFormat.format(Math.round(allowanceCost(liters, fuelAllowanceRate(data))))}</td><td>{numberFormat.format(Math.round(actualFuelValue))}</td>                <td className={deviationClass(deviation)}>{numberFormat.format(Math.round(deviation))}</td><td className={accumulated > 200 ? 'ledger-warning' : 'ledger-ok'}>{rateFormat.format(accumulated)}</td>{expenseFields.map((field) => <td key={field.key}>{numberFormat.format(entry.expenses?.[field.key] || 0)}</td>)}<td>{numberFormat.format(Math.round(totalCost))}{entry.note ? <small className="ledger-row-note">{entry.note}</small> : null}</td><td className="ledger-actions"><button type="button" className="entry-edit-button fleet-icon-action" aria-label="تعديل حركة الوقود" title="تعديل" onClick={() => editEntry(entry)}><Pencil size={15} aria-hidden="true" /></button><button type="button" className="entry-delete-button fleet-icon-action" aria-label="حذف حركة الوقود" title="حذف" onClick={() => openDeleteConfirmation(entry)}><Trash2 size={15} aria-hidden="true" /></button></td></tr>;
              })}</tbody>
              {selectedEntries.length > 0 && <tfoot><tr><th colSpan={5}>إجمالي السيارة</th><th>{numberFormat.format(selectedSummary.distance)} كم</th><th>{numberFormat.format(selectedSummary.liters)} لتر</th><th>{numberFormat.format(Math.round(selectedFuelTotals.allowance))}</th><th>{numberFormat.format(Math.round(selectedFuelTotals.actual))}</th><th>{numberFormat.format(Math.round(selectedFuelTotals.deviation))}</th><th>{rateFormat.format(totalKmPerLiter)}</th>{expenseFields.map((field) => <th key={field.key}>{numberFormat.format(selectedEntries.reduce((sum, entry) => sum + (entry.expenses?.[field.key] || 0), 0))}</th>)}<th>{numberFormat.format(Math.round(selectedSummary.cost))}</th><th /></tr></tfoot>}
            </table>
            {selectedEntries.length === 0 && <p className="fleet-empty-state">لا توجد مصروفات مسجلة لهذه السيارة بعد.</p>}
          </div>
        </> : <>
          <div className="fleet-section-heading fleet-vehicles-heading">
            <div><span className="fleet-overline">بيانات الأسطول</span><h2>ملخص السيارات</h2></div>
            <div className="fleet-vehicle-actions">
              <button className="fleet-button fleet-button-light" type="button" onClick={() => void downloadAllVehicleTemplate().catch((error: unknown) => { setNoticeIsError(true); setNotice(error instanceof Error ? `تعذر تنزيل قالب جميع السيارات: ${error.message}` : 'تعذر تنزيل قالب جميع السيارات.'); })}><Download size={14} aria-hidden="true" /> تنزيل قالب الجميع</button>
              {canEditFleet && <>
                <button className="fleet-button fleet-button-light" type="button" disabled={isImportingAllVehicleData} onClick={() => allVehiclesImportRef.current?.click()}><Upload size={14} aria-hidden="true" /> {isImportingAllVehicleData ? 'جاري الاستيراد...' : 'استيراد جميع السيارات'}</button>
                <input ref={allVehiclesImportRef} className="fleet-file-input" type="file" accept=".xlsx" aria-label="اختر قالب جميع السيارات للاستيراد" onChange={importAllVehicleTemplate} />
              </>}
              <button className="fleet-button fleet-button-primary" type="button" onClick={() => setVehicleDialog(true)}>+ إضافة سيارة</button>
            </div>
          </div>
          <div className="fleet-vehicle-grid">{summaries.map((summary) => {
            const { vehicle, lastOdometer, entries, distance, liters, expected, deviation, cost, kmPerLiter } = summary;
            const vehicleIndex = data.vehicles.findIndex((item) => item.id === vehicle.id);
            const VehicleIcon = [CarFront, Truck, BusFront, CarTaxiFront, Car][Math.max(0, vehicleIndex) % 5];
            const isExpanded = expandedVehicleCardId === vehicle.id;
            return <article className={`fleet-vehicle fleet-vehicle-card fleet-vehicle-tone${isExpanded ? ' is-expanded' : ''}`} style={vehicleToneStyle(vehicle.id, data.vehicles)} key={vehicle.id}>
              <div className="fleet-vehicle-summary-row">
                <button type="button" className="fleet-vehicle-summary-toggle" aria-expanded={isExpanded} aria-controls={`vehicle-card-details-${vehicle.id}`} onClick={() => setExpandedVehicleCardId((current) => current === vehicle.id ? null : vehicle.id)} onDoubleClick={() => setSelectedVehicleId(vehicle.id)}>
                  <span className="fleet-car-icon"><VehicleIcon size={25} strokeWidth={2.1} aria-hidden="true" /></span>
                  <span className="fleet-vehicle-summary-copy"><strong>{vehicle.plate}</strong><small>{vehicle.name}</small></span>
                </button>
                <button type="button" className="fleet-vehicle-edit-icon" aria-label={`تعديل السيارة ${vehicle.name}`} title="تعديل السيارة" onClick={() => openVehicleEditor(vehicle)}><Settings size={14} strokeWidth={2.2} aria-hidden="true" /></button>
                <span className="fleet-vehicle-summary-spacer" />
                {Math.abs(deviation) >= 0.5 && <span className={`fleet-gap ${deviation > 250 ? 'gap-high' : deviation < 0 ? 'gap-low' : 'gap-normal'}`}>الانحراف: {currencyFormat.format(deviation)}</span>}
                <span className="fleet-vehicle-expand-icon" aria-hidden="true">{isExpanded ? '−' : '+'}</span>
              </div>
              {isExpanded && <div className="fleet-vehicle-expanded" id={`vehicle-card-details-${vehicle.id}`}>
                <div className="fleet-vehicle-data"><span>السائق<strong>{vehicle.driver || 'غير محدد'}</strong></span><span>معدل اللترات<strong>{numberFormat.format(vehicleFuelRate(vehicle))} لتر / {numberFormat.format(vehicleDistanceBase(vehicle))} كم</strong></span><span>العداد الحالي<strong>{numberFormat.format(lastOdometer)} كم</strong></span><span className="fleet-fixed-reading"><label>القراءة السابقة<input type="number" min="0" value={openingOdometerDrafts[vehicle.id] ?? String(vehicle.openingOdometer)} aria-label={`القراءة السابقة لـ ${vehicle.name}`} onChange={(event) => setOpeningOdometerDrafts((current) => ({ ...current, [vehicle.id]: event.target.value }))} /><button type="button" onClick={() => saveOpeningOdometer(vehicle, entries.length)}>حفظ القراءة</button></label></span><span>المسافة المسجلة<strong>{numberFormat.format(distance)} كم</strong></span><span>كم/لتر الإجمالي<strong>{rateFormat.format(kmPerLiter)}</strong></span><span>إجمالي الانحراف<strong className={`fleet-gap${deviation > 250 ? ' gap-high' : deviation < 0 ? ' gap-low' : ''}`}>{currencyFormat.format(deviation)}</strong></span><span>إجمالي التكلفة<strong>{currencyFormat.format(cost)}</strong></span></div>
                <div className="fleet-card-actions"><button type="button" className="fleet-open-details" onClick={() => setSelectedVehicleId(vehicle.id)}>عرض تفاصيل السيارة <span aria-hidden="true">←</span></button></div>
              </div>}
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
          <button className="fleet-button fleet-button-primary fleet-icon-button" type="button" aria-label="تسجيل تعبئة" title="تسجيل تعبئة" onClick={() => startEntry()}><Fuel size={19} aria-hidden="true" /></button>
          <button className="fleet-button fleet-button-light fleet-icon-button" type="button" aria-label="إضافة صيانة" title="إضافة صيانة" onClick={openMaintenanceDialog}><Wrench size={18} aria-hidden="true" /></button>
          <button className="fleet-button fleet-button-light fleet-icon-button" type="button" aria-label="تصدير Excel" title="تصدير Excel" onClick={() => void exportExcel()}><svg className="excel-export-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 2.75h9l5 5v13.5H5z" fill="#fff" stroke="#217346" strokeWidth="1.5" strokeLinejoin="round" /><path d="M14 2.75v5h5" fill="#d9f0e2" stroke="#217346" strokeWidth="1.5" strokeLinejoin="round" /><path d="M2 8h11v12H2z" rx="1.5" fill="#217346" /><path d="m5 11 5 6m0-6-5 6" fill="none" stroke="#fff" strokeWidth="1.7" strokeLinecap="round" /></svg></button>
          <button className="fleet-button fleet-button-light fleet-icon-button" type="button" aria-label="إعادة ضبط" title="إعادة ضبط" onClick={() => { setSearch(''); setLedgerFilters({ vehicleId: 'all', from: '', to: '', search: '' }); }}><RotateCcw size={18} aria-hidden="true" /></button>
        </div>
        <div className="fleet-table-wrap"><table className="fleet-table fleet-entry-table"><thead><tr><th>التاريخ</th><th>السيارة</th><th>بداية العداد</th><th>العداد الحالي</th><th>الفجوة (كم)</th><th>الوقود المتوقع</th><th>المصروف</th><th>انحراف التموين (جنيه)</th><th>إجمالي المصروف</th><th>إجراء</th></tr></thead><tbody>
          {filteredEntries.map((entry) => { const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId); const distance = Math.max(0, entry.currentOdometer - entry.startOdometer); const expected = vehicle ? litersForDistance(vehicle, distance) : 0; const liters = litersForEntry(entry, vehicle); const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters); return <tr key={entry.id} className="fleet-vehicle-tone" style={vehicleToneStyle(entry.vehicleId, data.vehicles)}><td>{formatDateArabic(entry.date)}</td><td><strong>{vehicle?.name || 'سيارة محذوفة'}</strong><small>{vehicle?.plate}</small></td><td>{numberFormat.format(entry.startOdometer)}</td><td>{numberFormat.format(entry.currentOdometer)}</td><td>{numberFormat.format(distance)} كم</td><td>{numberFormat.format(expected)} لتر</td><td>{numberFormat.format(liters)} لتر</td><td className={deviationClass(deviation)}>{currencyFormat.format(deviation)}</td><td>{currencyFormat.format(entryTotalCost(entry))}</td><td className="ledger-actions"><button type="button" className="entry-edit-button fleet-icon-action" aria-label="تعديل حركة الوقود" title="تعديل" onClick={() => editEntry(entry)}><Pencil size={15} aria-hidden="true" /></button><button type="button" className="entry-delete-button fleet-icon-action" aria-label="حذف حركة الوقود" title="حذف" onClick={() => openDeleteConfirmation(entry)}><Trash2 size={15} aria-hidden="true" /></button></td></tr>; })}
        </tbody></table></div>
      </section>}

      <footer className="fleet-footer"><span>مراقبة سيارات شركة أبو غالي</span></footer>

      {vehicleDialog && <div className="fleet-modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setVehicleDialog(false)}><form className="fleet-modal" onSubmit={saveVehicle}><div className="fleet-modal-title"><div><span className="fleet-overline">بيانات المركبة</span><h2>{editingVehicleId ? 'تعديل بيانات السيارة' : 'إضافة سيارة'}</h2></div><button type="button" aria-label="إغلاق" onClick={() => { setVehicleDialog(false); setEditingVehicleId(null); }}>×</button></div><label>اسم السيارة<input required value={vehicleForm.name} onChange={(event) => setVehicleForm({ ...vehicleForm, name: event.target.value })} placeholder="مثال: تويوتا هايلوكس" /></label><label>رقم اللوحة<input required value={vehicleForm.plate} onChange={(event) => setVehicleForm({ ...vehicleForm, plate: event.target.value })} /></label><label>اسم السائق<input value={vehicleForm.driver} onChange={(event) => setVehicleForm({ ...vehicleForm, driver: event.target.value })} /></label><div className="fleet-form-row"><label>بداية العداد<input required type="number" min="0" value={vehicleForm.openingOdometer} onChange={(event) => setVehicleForm({ ...vehicleForm, openingOdometer: event.target.value })} /></label><label>اللترات المحددة<input required type="number" min="0.1" step="0.1" value={vehicleForm.kmPerLiter} onChange={(event) => setVehicleForm({ ...vehicleForm, kmPerLiter: event.target.value })} /></label></div><div className="fleet-form-row"><label>مسافة الأساس بالكيلومتر<input required type="number" min="1" step="1" value={vehicleForm.distanceBaseKm} onChange={(event) => setVehicleForm({ ...vehicleForm, distanceBaseKm: event.target.value })} /></label><span className="fleet-rate-help">المعادلة: الفجوة ÷ مسافة الأساس × اللترات</span></div><button className="fleet-button fleet-button-primary fleet-submit" type="submit">{editingVehicleId ? 'حفظ التعديل' : 'حفظ السيارة'}</button></form></div>}

      {maintenanceDialog && <div className="fleet-modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setMaintenanceDialog(false)}>
        <form className="fleet-modal" onSubmit={saveMaintenance}>
          <div className="fleet-modal-title"><div><span className="fleet-overline">مصروف مستقل</span><h2>إضافة صيانة</h2></div><button type="button" aria-label="إغلاق" onClick={() => setMaintenanceDialog(false)}>×</button></div>
          <label>السيارة<select required value={maintenanceForm.vehicleId} onChange={(event) => setMaintenanceForm((current) => ({ ...current, vehicleId: event.target.value }))}><option value="" disabled>اختر سيارة</option>{data.vehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</select></label>
          <div className="maintenance-amount-field">
            <label className="entry-field-label" htmlFor="maintenance-amount-input">مبلغ الصيانة بالجنيه</label>
            <div className="entry-amount-total"><span>الإجمالي</span><output aria-live="polite">{numberFormat.format(parseAmountExpression(maintenanceForm.amount) ?? 0)}</output></div>
            <div className="entry-amount-entry"><input id="maintenance-amount-input" dir="ltr" type="text" inputMode="decimal" value={maintenanceForm.amountDraft} onChange={(event) => setMaintenanceForm((current) => ({ ...current, amountDraft: event.target.value }))} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); addMaintenanceAmount(); } }} placeholder="أدخل مبلغًا" /><button type="button" className="entry-amount-add" onClick={addMaintenanceAmount}>إضافة</button></div>
            <div className="entry-amount-list" aria-label="مبالغ الصيانة المضافة">{splitAmountExpression(maintenanceForm.amount).map((amount, index) => <span className="entry-amount-chip" key={`${amount}-${index}`}><span dir="ltr">{amount}</span><button type="button" aria-label={`حذف مبلغ الصيانة ${amount}`} onClick={() => removeMaintenanceAmount(index)}>×</button></span>)}</div>
          </div>
          <p className="maintenance-entry-help">سيُضاف المبلغ إلى بند الصيانة في آخر سجل للسيارة المختارة.</p>
          <button className="fleet-button fleet-button-primary fleet-submit" type="submit" disabled={!data.vehicles.length}>حفظ الصيانة</button>
        </form>
      </div>}

      {entryDialog && <section className="fleet-entry-page" aria-label="صفحة تسجيل حركة المصروفات">
        <form className="fleet-modal fleet-entry-modal" onSubmit={saveEntry}>
          <header className="entry-form-header">
            <div><span>{entryVehicle ? `${entryVehicle.plate} / سجل المصاريف` : 'إدارة الأسطول / سجل المصاريف'}</span><h2>{editingEntryId ? 'تعديل حركة مسجلة' : 'إدخال حركة جديدة'}</h2></div>
            <button type="button" className="entry-page-back" onClick={closeEntryPage}>العودة</button>
          </header>
          <div className="entry-form-body">
            <section className="entry-form-section">
              <div className="entry-section-title"><span>01</span><div><h3>بيانات الرحلة</h3><p>السيارة والتاريخ والسائق والمسار</p></div></div>
              <div className="entry-form-grid entry-trip-grid">
                <label className="entry-field"><span>اختيار السيارة برقم اللوحة</span><input type="search" inputMode="numeric" value={entryPlateSearch} onChange={(event) => setEntryPlateSearch(event.target.value)} placeholder="اكتب رقم اللوحة هنا" /><select required value={entryForm.vehicleId} onChange={(event) => { const next = event.target.value; const vehicleSummary = summaries.find((item) => item.vehicle.id === next); const latest = vehicleSummary?.lastOdometer ?? 0; setEntryPlateSearch(vehicleSummary?.vehicle.plate || ''); setEntryForm({ ...entryForm, vehicleId: next, driver: vehicleSummary?.vehicle.driver || '', startOdometer: String(latest), currentOdometer: '', liters: '0.0' }); }}>{data.vehicles.filter((vehicle) => !entryPlateSearch.trim() || vehicle.plate.includes(entryPlateSearch.trim())).map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</select><small className="entry-field-help">اكتب رقم اللوحة للبحث ثم اختر النتيجة باسم السيارة الكامل</small></label>
                <label className="entry-field entry-date-field">
                  <span>التاريخ</span>
                  <input required type="date" lang="ar" dir="rtl" value={entryForm.date} onChange={(event) => setEntryForm((current) => ({ ...current, date: event.target.value }))} />
                  <small className="entry-field-help">اكتب التاريخ أو اختره من التقويم</small>
                </label>
                <label className="entry-field"><span>اسم السائق</span><input value={entryForm.driver} onChange={(event) => setEntryForm({ ...entryForm, driver: event.target.value })} /></label>
                <label className="entry-field"><span>خط السير</span><input value={entryForm.route} onChange={(event) => setEntryForm({ ...entryForm, route: event.target.value })} placeholder="نقطة الانطلاق والوجهة" /></label>
              </div>
            </section>

            <section className="entry-form-section entry-fuel-section">
              <button type="button" className="entry-section-title entry-section-toggle" aria-expanded={isFuelSectionExpanded} aria-controls="entry-fuel-content" onClick={() => setIsFuelSectionExpanded((expanded) => !expanded)}>
                <span>02</span><div><h3>العداد والتموين</h3><p>يُحسب فرق العداد تلقائيًا</p></div><ChevronDown aria-hidden="true" />
              </button>
              {isFuelSectionExpanded && <div className="entry-form-grid entry-fuel-grid" id="entry-fuel-content">
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
              </div>}
            </section>

            <section className="entry-form-section">
              <div className="entry-section-title"><span>03</span><div><h3>المصاريف الإضافية</h3><p>كل المبالغ بالجنيه · أدخل المبلغ واضغط إضافة ليُحسب في الإجمالي</p></div></div>
              <div className="entry-expense-grid">{expenseFields.map((field) => {
                const amount = expenseAmountPreviews[field.key];
                const isExpanded = expandedExpenseKey === field.key;
                return <div className={`entry-field entry-expense-item${isExpanded ? ' is-expanded' : ''}`} key={field.key}>
                  <button
                    type="button"
                    className="entry-expense-toggle"
                    aria-expanded={isExpanded}
                    aria-controls={`expense-panel-${field.key}`}
                    onClick={() => setExpandedExpenseKey((current) => current === field.key ? null : field.key)}
                  >
                    <span className="entry-expense-name">{field.label}</span>
                    <span className="entry-expense-total">{amount === null ? 'مبلغ غير صالح' : numberFormat.format(amount)}</span>
                    <span className="entry-expense-chevron" aria-hidden="true">{isExpanded ? '−' : '+'}</span>
                  </button>
                  {isExpanded && <div className="entry-expense-panel" id={`expense-panel-${field.key}`}>
                    <div className="entry-amount-entry"><input id={`expense-amount-${field.key}`} dir="ltr" type="text" inputMode="decimal" aria-label={`${field.label} بالجنيه`} value={expenseAmountDrafts[field.key]} onChange={(event) => setExpenseAmountDrafts((current) => ({ ...current, [field.key]: event.target.value }))} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); addExpenseAmount(field.key); } }} /><button type="button" className="entry-amount-add" onClick={() => addExpenseAmount(field.key)}>إضافة</button></div>
                    <div className="entry-amount-list" aria-label={`المبالغ المضافة: ${field.label}`}>{splitAmountExpression(entryForm.expenses[field.key]).map((part, index) => <span className="entry-amount-chip" key={`${part}-${index}`}><span dir="ltr">{part}</span><button type="button" aria-label={`حذف ${part} من ${field.label}`} onClick={() => removeExpenseAmount(field.key, index)}>×</button></span>)}</div>
                    {amount === null && <small className="entry-field-help" role="alert">راجع المبالغ المضافة</small>}
                  </div>}
                </div>;
              })}</div>
            </section>
          </div>
          <footer className="entry-form-actions"><div><span>إجمالي الحركة</span><strong>{hasInvalidAmount ? 'راجع المبالغ' : hasUnaddedAmounts ? 'اضغط إضافة أولًا' : currencyFormat.format(entryCostPreview)}</strong></div><button type="button" className="entry-cancel-button" onClick={closeEntryPage}>العودة</button><button className="fleet-button fleet-button-primary entry-save-button" type="submit">{editingEntryId ? 'حفظ التعديل' : 'حفظ الحركة'}</button></footer>
        </form>
      </section>}
      {pendingDeleteEntry && <div className="fleet-confirm-backdrop" role="presentation"><div className="fleet-confirm-dialog" role="alertdialog" aria-modal="true" aria-labelledby="delete-entry-title"><span className="fleet-confirm-icon">!</span><h2 id="delete-entry-title">تأكيد حذف السجل</h2><p>لحذف حركة التعبئة بتاريخ <strong>{formatDateArabic(pendingDeleteEntry.date)}</strong> اكتب كلمة <strong>موافق</strong> ثم اضغط تأكيد.</p><input className="fleet-confirm-input" aria-label="اكتب موافق للتأكيد" value={deleteConfirmationText} onChange={(event) => setDeleteConfirmationText(event.target.value)} autoFocus placeholder="اكتب: موافق" /><div><button type="button" className="fleet-confirm-cancel" onClick={cancelDeleteConfirmation}>إلغاء</button><button type="button" className="fleet-confirm-ok" disabled={deleteConfirmationText.trim() !== 'موافق'} onClick={confirmDeleteEntry}>تأكيد الحذف</button></div></div></div>}
    </main>
  );
};

export default CarFuelPage;