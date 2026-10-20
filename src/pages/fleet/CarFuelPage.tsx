import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { onAuthStateChanged } from 'firebase/auth';
import type ExcelJS from 'exceljs';
import { ArrowLeft, Bell, BusFront, CalendarDays, Calculator, Car, CarFront, CarTaxiFront, Check, CheckCheck, ChevronDown, ClipboardList, Download, Fuel, Gauge, Home, LayoutDashboard, LogOut, MapPin, MapPinned, MoreVertical, Pencil, Plus, RefreshCw, RotateCcw, Settings, Trash2, TrendingDown, TrendingUp, Truck, Upload, Wrench } from 'lucide-react';
import { addGuest, deleteGuest as deleteGuestFromCloud, setGuestActive, subscribeToFleetUsers, updateGuest } from '../../services/fleetUserStore';
import type { FleetGuest, FleetUser } from '../../services/fleetUserStore';
import { firestore, isFirebaseConfigured } from '../../services/firebase';
import { readCloudFleetData, readCloudFleetDataStrict, readCloudRouteEntries, saveCloudFleetDataStrict, saveCloudRouteEntries } from '../../services/fleetCloudStore';
import { firebaseAuth } from '../../services/firebase';
import { publishDeviationNotification, subscribeToDeviationNotifications } from '../../services/deviationNotificationStore';
import type { DeviationNotification, FleetNotification, FleetUpdateNotification } from '../../services/deviationNotificationStore';
import { recordFleetAuditEntry, subscribeToFleetAuditLog } from '../../services/fleetAuditLog';
import type { FleetAuditAction, FleetAuditEntry } from '../../services/fleetAuditLog';
import { applyUpdate } from '../../services/serviceWorker';
import { clearDeferredPrompt, getDeferredPrompt } from '../../services/pwaInstall';
import { loadMoreGpsLogs, subscribeToGpsLogs } from '../../services/gpsStore';
import type { GpsLog, GpsLogCursor } from '../../services/gpsStore';
import './car-fuel.css';

type VehicleType = 'dmax' | 'chevrolet' | 'bus';
type VehicleTypeColors = Record<VehicleType, string>;
const vehicleTypeOptions: { value: VehicleType; label: string }[] = [
  { value: 'dmax', label: 'ديمكس' },
  { value: 'chevrolet', label: 'شفر' },
  { value: 'bus', label: 'منى باص' },
];
const isVehicleType = (value: string): value is VehicleType => vehicleTypeOptions.some((option) => option.value === value);
const defaultVehicleTypeColors: VehicleTypeColors = {
  dmax: '#19776c',
  chevrolet: '#8a3f70',
  bus: '#426f91',
};
const normalizeVehicleTypeColors = (colors?: Partial<VehicleTypeColors>): VehicleTypeColors => ({
  dmax: typeof colors?.dmax === 'string' && /^#[0-9a-f]{6}$/i.test(colors.dmax) ? colors.dmax : defaultVehicleTypeColors.dmax,
  chevrolet: typeof colors?.chevrolet === 'string' && /^#[0-9a-f]{6}$/i.test(colors.chevrolet) ? colors.chevrolet : defaultVehicleTypeColors.chevrolet,
  bus: typeof colors?.bus === 'string' && /^#[0-9a-f]{6}$/i.test(colors.bus) ? colors.bus : defaultVehicleTypeColors.bus,
});
const vehicleTypeLabel = (vehicle: Vehicle) => {
  const type = vehicle.vehicleType || inferVehicleType(vehicle.name);
  return vehicleTypeOptions.find((option) => option.value === type)?.label || 'نوع غير محدد';
};
const inferVehicleType = (name: string): VehicleType | undefined => {
  const normalizedName = name.toLocaleLowerCase('ar');
  if (/ديمكس|ديماكس|d-?max/.test(normalizedName)) return 'dmax';
  if (/شفر|شيفر|شفروليه|chevrolet/.test(normalizedName)) return 'chevrolet';
  if (/باص|أتوبيس|اوتوبيس|bus/.test(normalizedName)) return 'bus';
  return undefined;
};

type Vehicle = {
  id: string;
  name: string;
  vehicleType?: VehicleType;
  plate: string;
  driver: string;
  openingOdometer: number;
  kmPerLiter: number;
  fuelLitersPer105Km?: number;
  fuelDistanceBaseKm?: number;
};

type FuelEntry = {
  id: string;
  createdAt?: string;
  documentNumber?: string;
  vehicleId: string;
  date: string;
  startOdometer: number;
  currentOdometer: number;
  liters: number;
  pricePerLiter: number;
  allowancePricePerLiter?: number;
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
  isDailyVehicle?: boolean;
  dailyVehicleId?: string;
  vehicleName: string;
  plate: string;
  departure: string;
  route: string;
  representative: string;
  driver: string;
};

type RouteDailyDraft = Pick<RouteEntry, 'departure' | 'route' | 'representative' | 'driver'>;
type RouteBoardVehicle = { id: string; fleetVehicleId: string | null; name: string; plate: string; isDailyOnly: boolean };
type DeviationNote = { text: string; createdAt: string };
type DeviationNotesByMonth = Record<string, Record<string, DeviationNote>>;

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
  documentNumber: string;
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

type EntryColumnKey = 'documentNumber' | 'date' | 'vehicle' | 'startOdometer' | 'currentOdometer' | 'distance' | 'expectedFuel' | 'fuel' | 'deviation' | 'totalCost';
type DetailColumnKey = 'documentNumber' | 'date' | 'driver' | 'route' | 'startOdometer' | 'currentOdometer' | 'distance' | 'liters' | 'allowance' | 'actual' | 'deviation' | 'rate' | `expense:${keyof ExpenseBreakdown}` | 'totalCost';
const entryColumns: { key: EntryColumnKey; label: string }[] = [
  { key: 'documentNumber', label: 'رقم المستند' },
  { key: 'date', label: 'التاريخ' },
  { key: 'vehicle', label: 'السيارة' },
  { key: 'startOdometer', label: 'سابقة' },
  { key: 'currentOdometer', label: 'حالية' },
  { key: 'distance', label: 'الفجوة (كم)' },
  { key: 'expectedFuel', label: 'الوقود المتوقع' },
  { key: 'fuel', label: 'المصروف' },
  { key: 'deviation', label: 'انحراف التموين (جنيه)' },
  { key: 'totalCost', label: 'الإجمالي' },
];
const defaultEntryColumnVisibility = Object.fromEntries(entryColumns.map(({ key }) => [key, true])) as Record<EntryColumnKey, boolean>;
const readDetailColumnVisibility = (storageKey: string, legacyStorageKey: string): Partial<Record<DetailColumnKey, boolean>> => {
  try {
    const stored = JSON.parse(localStorage.getItem(storageKey) || 'null') as unknown;
    const legacy = JSON.parse(localStorage.getItem(legacyStorageKey) || 'null') as unknown;
    const storedPreferences = stored && typeof stored === 'object' && !Array.isArray(stored)
      ? stored as Record<string, unknown>
      : legacy && typeof legacy === 'object' && !Array.isArray(legacy)
        ? Object.values(legacy as Record<string, unknown>).find((value) => value && typeof value === 'object' && !Array.isArray(value)) as Record<string, unknown> | undefined
        : undefined;
    if (!storedPreferences) return {};
    return Object.fromEntries(
      Object.entries(storedPreferences).filter((entry): entry is [DetailColumnKey, boolean] => typeof entry[1] === 'boolean'),
    ) as Partial<Record<DetailColumnKey, boolean>>;
  } catch (error) {
    console.error('Failed to read vehicle detail column preferences', error);
    return {};
  }
};

type FleetData = { vehicles: Vehicle[]; entries: FuelEntry[]; allowancePricePerLiter?: number; vehicleTypeColors?: VehicleTypeColors; hiddenDeviationVehicleIds?: string[] };
const STORAGE_KEY = 'company-fleet-fuel-v1';
const ROUTE_STORAGE_KEY = 'company-fleet-routes-v1';
const DEFAULT_ALLOWANCE_PRICE = 15.75;
const themePresets = {
  'dark-premium': {
    label: 'داكن',
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
    fontHeading: 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
  },
  'emerald-light': {
    label: 'دافئ',
    emoji: '✨',
    description: 'كريمي + بني ذهبي',
    background: '#fdfbf7',
    surface: '#fffbf0',
    soft: '#f5ead5',
    ink: '#44403c',
    inkSoft: '#57534e',
    inkMuted: '#78716c',
    line: '#e7e5e4',
    accent: '#b45309',
    accentDark: '#92400e',
    accentSoft: 'rgba(180,83,9,0.12)',
    heading: '#78350f',
    fontHeading: 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
  },
  'warm-luxury': {
    label: 'برتقالي',
    emoji: '🚚',
    description: 'برتقالي حيوي + رمادي ثلجي',
    background: '#f8fafc',
    surface: '#ffffff',
    soft: '#fff7ed',
    ink: '#1e293b',
    inkSoft: '#334155',
    inkMuted: '#64748b',
    line: '#ffedd5',
    accent: '#ea580c',
    accentDark: '#c2410c',
    accentSoft: 'rgba(234,88,12,0.12)',
    heading: '#9a3412',
    fontHeading: 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
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
const vehicleToneStyle = (vehicleId: string, vehicles: Vehicle[], vehicleTypeColors: VehicleTypeColors = defaultVehicleTypeColors): React.CSSProperties => {
  const vehicle = vehicles.find((item) => item.id === vehicleId);
  const vehicleType = vehicle?.vehicleType || (vehicle ? inferVehicleType(vehicle.name) : undefined);
  const accent = vehicleType ? vehicleTypeColors[vehicleType] : undefined;
  const index = vehicles.findIndex((item) => item.id === vehicleId);
  const palette = [
    ['#19776c', '#edf6f3', '#cfe4dc'],
    ['#426f91', '#eef4f8', '#d3e0e9'],
    ['#a66e2d', '#faf4e9', '#e8dbc2'],
    ['#a85851', '#faf0ef', '#ead7d4'],
    ['#6d7d45', '#f3f5ec', '#dfe4d0'],
  ];
  const [fallbackAccent, tint, line] = index < palette.length
    ? palette[Math.max(0, index)]
    : (() => {
      const hue = (index * 137.508 + 29) % 360;
      return [`hsl(${hue} 36% 38%)`, `hsl(${hue} 32% 96%)`, `hsl(${hue} 25% 84%)`];
    })();
  const effectiveAccent = accent || fallbackAccent;
  return {
    '--vehicle-accent': effectiveAccent,
    '--vehicle-tint': accent ? `color-mix(in srgb, ${accent} 8%, white)` : tint,
    '--vehicle-line': accent ? `color-mix(in srgb, ${accent} 25%, white)` : line,
  } as React.CSSProperties;
};

const initialData: FleetData = {
  vehicleTypeColors: defaultVehicleTypeColors,
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
    if (!stored) return initialData;
    const parsed = JSON.parse(stored) as FleetData;
    return {
      ...parsed,
      vehicleTypeColors: normalizeVehicleTypeColors(parsed.vehicleTypeColors),
      hiddenDeviationVehicleIds: Array.isArray(parsed.hiddenDeviationVehicleIds)
        ? [...new Set(parsed.hiddenDeviationVehicleIds.filter((id): id is string => typeof id === 'string'))]
        : [],
    };
  } catch {
    return initialData;
  }
};

const normalizeFleetData = (data: FleetData): FleetData => migrateLegacyFuelEntries({
  ...data,
  vehicleTypeColors: normalizeVehicleTypeColors(data.vehicleTypeColors),
  hiddenDeviationVehicleIds: Array.isArray(data.hiddenDeviationVehicleIds)
    ? [...new Set(data.hiddenDeviationVehicleIds.filter((id): id is string => typeof id === 'string'))]
    : [],
});
const fleetDataFingerprint = (data: FleetData) => JSON.stringify({
  vehicles: data.vehicles,
  entries: data.entries,
  allowancePricePerLiter: data.allowancePricePerLiter,
  vehicleTypeColors: normalizeVehicleTypeColors(data.vehicleTypeColors),
  hiddenDeviationVehicleIds: data.hiddenDeviationVehicleIds || [],
});
const isFleetUpdateNotification = (notification: FleetNotification): notification is FleetUpdateNotification => notification.kind === 'fleet-update';
const normalizeVehicleSearchText = (value: string) => value
  .toLocaleLowerCase('ar')
  .normalize('NFKC')
  .replace(/[أإآٱ]/g, 'ا')
  .replace(/ؤ/g, 'و')
  .replace(/ئ/g, 'ي')
  .replace(/ء/g, '')
  .normalize('NFKD')
  .replace(/[\u064B-\u065F\u0670\u06D6-\u06ED]/g, '')
  .replace(/[٠-٩]/g, (digit) => String(digit.charCodeAt(0) - 0x0660))
  .replace(/[۰-۹]/g, (digit) => String(digit.charCodeAt(0) - 0x06f0))
  .replace(/\s+/g, ' ')
  .trim();
const normalizeVehicleSearchKey = (value: string) => normalizeVehicleSearchText(value).replace(/\s/g, '');
const filterVehicles = (query: string, vehicles: Vehicle[]) => {
  const searchKey = normalizeVehicleSearchKey(query);
  if (!searchKey) return vehicles;
  return vehicles.filter((vehicle) => (
    normalizeVehicleSearchKey(vehicle.plate).includes(searchKey)
    || normalizeVehicleSearchKey(vehicle.name).includes(searchKey)
  ));
};
const findUniqueVehicleMatch = (query: string, vehicles: Vehicle[]) => {
  const matches = filterVehicles(query, vehicles);
  return normalizeVehicleSearchKey(query) && matches.length === 1 ? matches[0] : undefined;
};
const matchesVehicleSearchText = (query: string, value: string) => {
  const searchKey = normalizeVehicleSearchKey(query);
  return !searchKey || normalizeVehicleSearchKey(value).includes(searchKey);
};
type NotificationListItem = { kind: 'single'; notification: FleetNotification } | {
  kind: 'deviation-group';
  id: string;
  notifications: DeviationNotification[];
  plates: string[];
  createdAt: string;
};
const formatNotificationTimestamp = (value: string) => {
  const timestamp = new Date(value);
  return Number.isNaN(timestamp.getTime())
    ? ''
    : new Intl.DateTimeFormat('ar-EG', { dateStyle: 'short', timeStyle: 'short' }).format(timestamp);
};

const isRouteEntry = (value: unknown): value is RouteEntry => {
  if (!value || typeof value !== 'object') return false;
  const entry = value as Partial<RouteEntry>;
  return typeof entry.id === 'string'
    && typeof entry.date === 'string'
    && (typeof entry.vehicleId === 'string' || entry.vehicleId === null)
    && (entry.isDailyVehicle === undefined || typeof entry.isDailyVehicle === 'boolean')
    && (entry.dailyVehicleId === undefined || typeof entry.dailyVehicleId === 'string')
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
const formatDateFilter = (value: string) => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return '';
  const [year, month, day] = value.split('-');
  return `${day}/${month}/${year}`;
};
const parseDateFilter = (value: string): string | null => {
  const normalized = normalizeVehicleSearchText(value).replace(/[.-]/g, '/');
  if (!normalized) return '';
  const match = normalized.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;
  const [, dayText, monthText, yearText] = match;
  const day = Number(dayText);
  const month = Number(monthText);
  const year = Number(yearText);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return `${yearText}-${monthText.padStart(2, '0')}-${dayText.padStart(2, '0')}`;
};
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
const compactExpenseHeaders: Record<keyof ExpenseBreakdown, string> = {
  mealAllowance: 'وجبات',
  extra: 'إضافي',
  maintenance: 'صيانة',
  tolls: 'كارتة',
  traffic: 'مرور',
  violations: 'مخالفات',
  carWash: 'غسيل',
  tireExpenses: 'كاوتش',
  oilChange: 'زيت',
  transport: 'نقل',
  other: 'أخرى',
};
const templateHeaderAliases: Record<string, string> = {
  'اللوحة': 'رقم اللوحة',
  'السيارة': 'اسم السيارة',
  'السائق': 'اسم السائق',
  'خط سير': 'خط السير',
  'المسار': 'خط السير',
  'سابقة': 'القراءة السابقة',
  'حالية': 'القراءة الحالية',
  'قراءة سابقة': 'القراءة السابقة',
  'قراءة حالية': 'القراءة الحالية',
  'تموين فعلي': 'مصروف الوقود الفعلي (جنيه)',
  'التموين الفعلي': 'مصروف الوقود الفعلي (جنيه)',
  'مصروف الوقود': 'مصروف الوقود الفعلي (جنيه)',
  'مصروف الوقود الفعلي': 'مصروف الوقود الفعلي (جنيه)',
  'قيمة التموين الفعلية': 'مصروف الوقود الفعلي (جنيه)',
  'قيمة التموين الفعلية (جنيه)': 'مصروف الوقود الفعلي (جنيه)',
  'اللترات': 'التموين الفعلي (لتر - محسوب)',
  'عدد اللترات': 'التموين الفعلي (لتر - محسوب)',
  'تموين قراءة': 'قيمة التموين المقررة (جنيه - محسوبة)',
  'التموين المقرر': 'قيمة التموين المقررة (جنيه - محسوبة)',
  'قيمة التموين المقررة': 'قيمة التموين المقررة (جنيه - محسوبة)',
  'الانحراف': 'الانحراف (جنيه - محسوب)',
  'الإجمالي': 'إجمالي مصاريف السيارة (جنيه - محسوب)',
  'إجمالي المصروف': 'إجمالي مصاريف السيارة (جنيه - محسوب)',
  'ملاحظات': 'التفصيل',
  ...Object.fromEntries(expenseFields.map(({ key, label }) => [
    compactExpenseHeaders[key],
    `${label} (جنيه)`,
  ])),
};
const normalizeTemplateHeader = (header: string) => {
  const normalized = header.trim();
  return Object.prototype.hasOwnProperty.call(templateHeaderAliases, normalized)
    ? templateHeaderAliases[normalized]
    : normalized;
};
const normalizeTemplateHeaders = (headers: string[]) => {
  const normalizedHeaders: string[] = [];
  const seenHeaders = new Set<string>();

  headers.forEach((headerValue) => {
    const header = normalizeTemplateHeader((headerValue ?? '').trim());
    if (!header) return;
    const trimmed = header.trim();
    const isWorksheetTitle = /^(سجل السيارة|قالب مصروفات السيارة|كشف الحركة|قالب|بيانات السيارة)/i.test(trimmed)
      && /[-–—]/.test(trimmed);
    if (isWorksheetTitle) return;
    if (seenHeaders.has(trimmed)) return;
    seenHeaders.add(trimmed);
    normalizedHeaders.push(trimmed);
  });

  return normalizedHeaders;
};
const detailColumns: { key: DetailColumnKey; label: string }[] = [
  { key: 'documentNumber', label: 'رقم المستند' },
  { key: 'date', label: 'التاريخ' },
  { key: 'driver', label: 'السائق' },
  { key: 'route', label: 'خط سير' },
  { key: 'startOdometer', label: 'سابقة' },
  { key: 'currentOdometer', label: 'حالية' },
  { key: 'distance', label: 'الفجوة (كم)' },
  { key: 'liters', label: 'عدد اللترات' },
  { key: 'allowance', label: 'تموين قراءة' },
  { key: 'actual', label: 'تموين فعلي' },
  { key: 'deviation', label: 'الانحراف' },
  { key: 'rate', label: 'كم/لتر' },
  ...expenseFields.map(({ key }) => ({ key: `expense:${key}` as const, label: compactExpenseHeaders[key] })),
  { key: 'totalCost', label: 'الإجمالي' },
];
const defaultDetailColumnVisibility = Object.fromEntries(detailColumns.map(({ key }) => [key, true])) as Record<DetailColumnKey, boolean>;
const emptyExpenses = (): ExpenseBreakdown => ({ mealAllowance: 0, tolls: 0, traffic: 0, violations: 0, maintenance: 0, extra: 0, carWash: 0, tireExpenses: 0, oilChange: 0, transport: 0, other: 0 });
const emptyExpenseInputs = (): ExpenseInputs => Object.fromEntries(expenseFields.map(({ key }) => [key, ''])) as ExpenseInputs;
const createEmptyEntryForm = (): EntryForm => ({ documentNumber: '', vehicleId: '', driver: '', startOdometer: '', currentOdometer: '', liters: '', pricePerLiter: '2.18', fuelActualValue: '', allowanceValue: '', date: '', route: '', note: '', expenses: emptyExpenseInputs() });
const DEVIATION_NOTES_STORAGE_KEY = 'cars-site-deviation-notes-v1';
const readDeviationNotes = (): DeviationNotesByMonth => {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(DEVIATION_NOTES_STORAGE_KEY) || '{}');
    if (!stored || typeof stored !== 'object' || Array.isArray(stored)) return {};
    return stored as DeviationNotesByMonth;
  } catch (error) {
    console.error('Failed to read monthly deviation notes', error);
    return {};
  }
};
const createVehicleWorksheetName = (plate: string, usedNames: Set<string>): string => {
  const baseName = plate.trim()
    .replace(/\s+/g, '-')
    .replace(/[\\/?*:\[\]]/g, '')
    .replace(/^'+|'+$/g, '')
    .replace(/-+/g, '-')
    .slice(0, 31)
    .replace(/'+$/g, '') || 'سيارة';
  let worksheetName = baseName;
  let suffix = 2;
  while (usedNames.has(worksheetName.toLocaleLowerCase('ar-EG'))) {
    const suffixText = `-${suffix}`;
    worksheetName = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`;
    suffix += 1;
  }
  usedNames.add(worksheetName.toLocaleLowerCase('ar-EG'));
  return worksheetName;
};
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
const entryAllowancePrice = (entry: FuelEntry, fallbackRate: number) => entry.allowancePricePerLiter ?? fallbackRate;
const entryAllowanceValue = (entry: FuelEntry, fallbackRate: number, liters = entry.liters) => (
  entry.allowanceValue ?? allowanceCost(liters, entryAllowancePrice(entry, fallbackRate))
);
const entryDeviation = (entry: FuelEntry, rate: number, liters = entry.liters) => entryFuelCost(entry) - entryAllowanceValue(entry, rate, liters);
const entryTotalCost = (entry: FuelEntry) => Math.round(entryFuelCost(entry) + sumExpenses(entry.expenses));
const migrateLegacyFuelEntries = (data: FleetData): FleetData => {
  const fallbackRate = fuelAllowanceRate(data);
  let hasLegacyEntries = false;
  const entries = data.entries.map((entry) => {
    if (entry.allowancePricePerLiter !== undefined && entry.allowanceValue !== undefined) return entry;
    hasLegacyEntries = true;
    const allowancePricePerLiter = entry.allowancePricePerLiter ?? fallbackRate;
    return {
      ...entry,
      allowancePricePerLiter,
      allowanceValue: entry.allowanceValue ?? allowanceCost(entry.liters, allowancePricePerLiter),
    };
  });
  return hasLegacyEntries ? { ...data, entries } : data;
};

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
  { key: 'vehicles', label: 'منظومة سيارات الشركة', description: 'بيانات ومصروفات المركبات', Icon: CarFront },
  { key: 'entries', label: 'سجل التعبئة', description: 'حركات الوقود المسجلة', Icon: Fuel },
  { key: 'routes', label: 'GPS', description: 'تسجيل ومتابعة تحركات السيارات', Icon: MapPinned },
  { key: 'audit', label: 'سجل العمليات', description: 'تعديلات وحذف البيانات', Icon: ClipboardList },
] as const;
const fleetHomeActions = [
  { key: 'users', label: 'إدارة الزوار', description: 'إضافة الزوار وتحديث صلاحيات مشاهدتهم', Icon: Settings },
] as const;
type CarFuelPageProps = { onLogout?: () => void; currentUser: FleetUser };
type LedgerPeriod = 'all' | 'month' | 'custom';
type FleetRoute = { tab: 'home' | 'overview' | 'vehicles' | 'entries' | 'routes' | 'users' | 'audit'; selectedVehicleId: string | null; entryPage?: boolean };
type FleetHistoryState = { fleetRoute?: FleetRoute };

const dateInputValue = (date: Date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
const formatEntryDateInput = (value: string) => {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : value;
};
const parseEntryDateInput = (value: string): string | null => {
  const match = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?$/.exec(value.trim());
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3] || new Date().getFullYear());
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month - 1, day);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
};
const parseDriverSettlementDate = (value: string) => parseEntryDateInput(value
  .replace(/[٠-٩]/g, (digit) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)))
  .replace(/[۰-۹]/g, (digit) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit))));
const deviationClass = (value: number) => value < 0 ? 'ledger-ok' : value > 250 ? 'ledger-warning' : 'ledger-neutral';
const CLEAR_ENTRIES_CONFIRMATION_PASSWORD = '1020350';
const monthNames = Array.from({ length: 12 }, (_, index) => new Intl.DateTimeFormat('ar-EG', { month: 'long' }).format(new Date(2020, index, 1)));
const deviationMonthDisplayName = (month: string): string => {
  const [year, monthNumber] = month.split('-').map(Number);
  return `${monthNames[monthNumber - 1]} ${year}`;
};
type MonthYearPickerProps = {
  selectedMonth: string | null;
  onSelectMonth: (monthYear: string) => void;
  ariaLabel: string;
  isOpen: boolean;
  onOpenChange: (open: boolean) => void;
  placeholder?: string;
  isActive?: boolean;
  className?: string;
  includeAllMonths?: boolean;
  onSelectAllMonths?: () => void;
  portalStyle: React.CSSProperties;
};
const MonthYearPicker: React.FC<MonthYearPickerProps> = ({
  selectedMonth,
  onSelectMonth,
  ariaLabel,
  isOpen,
  onOpenChange,
  placeholder = 'اختر الشهر والسنة',
  isActive = true,
  className = '',
  includeAllMonths = false,
  onSelectAllMonths,
  portalStyle,
}) => {
  const [year, setYear] = useState(() => selectedMonth ? Number(selectedMonth.slice(0, 4)) : new Date().getFullYear());
  const selectedLabel = selectedMonth
    ? new Intl.DateTimeFormat('ar-EG', { month: 'long', year: 'numeric' }).format(new Date(`${selectedMonth}-01T00:00:00`))
    : placeholder;

  return <div className={`ledger-month-picker ${className}`.trim()}>
    <button
      type="button"
      className={`ledger-month-select${isActive ? ' is-active' : ''}`}
      aria-label={ariaLabel}
      aria-haspopup="dialog"
      aria-expanded={isOpen}
      onClick={() => {
        if (!isOpen) setYear(selectedMonth ? Number(selectedMonth.slice(0, 4)) : new Date().getFullYear());
        onOpenChange(!isOpen);
      }}
    >
      {selectedLabel}
    </button>
    {isOpen && createPortal(<div
      className="ledger-month-picker-modal-backdrop"
      style={portalStyle}
      dir="rtl"
      role="presentation"
      onMouseDown={(event) => event.target === event.currentTarget && onOpenChange(false)}
    >
      <div className="ledger-month-picker-panel" role="dialog" aria-modal="true" aria-label={ariaLabel}>
        {includeAllMonths && <button
          type="button"
          className={`deviation-all-months-option${selectedMonth === null ? ' is-selected' : ''}`}
          onClick={() => {
            onSelectAllMonths?.();
            onOpenChange(false);
          }}
        >
          كل الأشهر
        </button>}
        <div className="ledger-month-picker-header">
          <button type="button" aria-label="السنة السابقة" onClick={() => setYear((currentYear) => currentYear - 1)}>‹</button>
          <strong>{year}</strong>
          <button type="button" aria-label="السنة التالية" onClick={() => setYear((currentYear) => currentYear + 1)}>›</button>
        </div>
        <div className="ledger-month-picker-months">
          {monthNames.map((month, index) => {
            const value = `${year}-${String(index + 1).padStart(2, '0')}`;
            return <button
              key={month}
              type="button"
              className={selectedMonth === value ? 'is-selected' : ''}
              onClick={() => {
                onSelectMonth(value);
                onOpenChange(false);
              }}
            >
              {month}
            </button>;
          })}
        </div>
      </div>
    </div>, document.body)}
  </div>;
};
const CarFuelPage: React.FC<CarFuelPageProps> = ({ onLogout, currentUser }) => {
  const entryDraftStorageKey = `company-fleet-entry-draft-v1-${encodeURIComponent(currentUser.email || currentUser.username)}`;
  const [initialEntryDraft] = useState(() => readEntryDraft(entryDraftStorageKey));
  const entryDraftSuppressedRef = useRef(false);
  const [data, setData] = useState<FleetData>(() => normalizeFleetData(readData()));
  const latestFleetDataRef = useRef(data);
  latestFleetDataRef.current = data;
  const lastSyncedFleetFingerprintRef = useRef<string | null>(null);
  const [routeEntries, setRouteEntries] = useState<RouteEntry[]>(readRouteEntries);
  const latestRouteEntriesRef = useRef(routeEntries);
  latestRouteEntriesRef.current = routeEntries;
  const [gpsLogs, setGpsLogs] = useState<GpsLog[]>([]);
  const [gpsLogCursor, setGpsLogCursor] = useState<GpsLogCursor | null>(null);
  const [hasMoreGpsLogs, setHasMoreGpsLogs] = useState(false);
  const [isLoadingGpsLogs, setIsLoadingGpsLogs] = useState(false);
  const [isLoadingMoreGpsLogs, setIsLoadingMoreGpsLogs] = useState(false);
  const [gpsLogsError, setGpsLogsError] = useState('');
  const [cloudLoaded, setCloudLoaded] = useState(!isFirebaseConfigured);
  const [routesCloudLoaded, setRoutesCloudLoaded] = useState(!isFirebaseConfigured);
  const [routesCloudAvailable, setRoutesCloudAvailable] = useState(!isFirebaseConfigured);
  const [tab, setTab] = useState<'home' | 'overview' | 'vehicles' | 'entries' | 'routes' | 'users' | 'audit'>('home');
  const routeHistoryRef = useRef<FleetRoute | null>(null);
  const skipRoutePushRef = useRef(false);
  const exitConfirmedRef = useRef(false);
  const currentTabRef = useRef(tab);
  currentTabRef.current = tab;
  const [guests, setGuests] = useState<FleetGuest[]>([]);
  const [auditEntries, setAuditEntries] = useState<FleetAuditEntry[]>([]);
  const [auditLogError, setAuditLogError] = useState('');
  const [guestDraft, setGuestDraft] = useState({ password: '', employeeName: '' });
  const [editingGuest, setEditingGuest] = useState<FleetGuest | null>(null);
  const [showGuestDialog, setShowGuestDialog] = useState(false);
  const [isSavingGuest, setIsSavingGuest] = useState(false);
  const themeStorageKey = `cars-site-theme-${currentUser.email || currentUser.username}`;
  const dismissedDeviationNotificationsKey = `cars-site-dismissed-deviation-notifications-${encodeURIComponent(currentUser.email || currentUser.username)}`;
  const entryColumnVisibilityKey = `cars-site-entry-columns-${encodeURIComponent(currentUser.id || currentUser.email || currentUser.username)}`;
  const detailColumnVisibilityKey = 'fleet-columns-global';
  const legacyDetailColumnVisibilityKey = `cars-site-vehicle-detail-columns-${encodeURIComponent(currentUser.id || currentUser.email || currentUser.username)}`;
  const [themeKey, setThemeKey] = useState<ThemeKey>(() => resolveTheme(localStorage.getItem(themeStorageKey)));
  const [fontSizeKey, setFontSizeKey] = useState<FontSizeKey>(() => (localStorage.getItem(`${themeStorageKey}-font`) as FontSizeKey) || 'medium');
  const [openingOdometerDrafts, setOpeningOdometerDrafts] = useState<Record<string, string>>({});
  const [isThemePickerOpen, setIsThemePickerOpen] = useState(false);
  const [isMobileActionsOpen, setIsMobileActionsOpen] = useState(false);
  const mobileActionsRef = useRef<HTMLDivElement>(null);
  const theme = themePresets[themeKey];
  const monthPickerPortalStyle = {
    '--ink': theme.ink,
    '--surface': theme.surface,
    '--surface-soft': theme.soft,
    '--line': theme.line,
    '--teal': theme.accent,
    '--fleet-font-scale': fontSizeKey === 'small' ? '.88' : fontSizeKey === 'large' ? '1.12' : '1',
  } as React.CSSProperties;
  const [selectedVehicleId, setSelectedVehicleId] = useState<string | null>(null);
  const [isRateSummaryExpanded, setIsRateSummaryExpanded] = useState(true);
  const [isQuickNavOpen, setIsQuickNavOpen] = useState(false);
  const quickNavRef = useRef<HTMLDivElement>(null);
  const [expandedVehicleCardId, setExpandedVehicleCardId] = useState<string | null>(null);
  const [expandedMovementCards, setExpandedMovementCards] = useState<Set<string>>(() => new Set());
  const [search, setSearch] = useState('');
  const [ledgerSummaryVehicleId, setLedgerSummaryVehicleId] = useState(() => data.vehicles[0]?.id ?? '');
  const [entryColumnVisibility, setEntryColumnVisibility] = useState<Record<EntryColumnKey, boolean>>(() => {
    try {
      const storedVisibility = JSON.parse(localStorage.getItem(entryColumnVisibilityKey) || 'null') as Partial<Record<EntryColumnKey, unknown>> | null;
      if (!storedVisibility || typeof storedVisibility !== 'object') return defaultEntryColumnVisibility;
      return Object.fromEntries(entryColumns.map(({ key }) => [
        key,
        typeof storedVisibility[key] === 'boolean' ? storedVisibility[key] : true,
      ])) as Record<EntryColumnKey, boolean>;
    } catch (error) {
      console.error('Failed to read entry table column preferences', error);
      return defaultEntryColumnVisibility;
    }
  });
  const [isEntryColumnsPanelOpen, setIsEntryColumnsPanelOpen] = useState(false);
  const [detailColumnVisibility, setDetailColumnVisibility] = useState(() => readDetailColumnVisibility(detailColumnVisibilityKey, legacyDetailColumnVisibilityKey));
  const [isDetailColumnsPanelOpen, setIsDetailColumnsPanelOpen] = useState(false);
  const [ledgerFilters, setLedgerFilters] = useState({ vehicleId: 'all', from: '', to: '', search: '' });
  const [ledgerFromDateDraft, setLedgerFromDateDraft] = useState('');
  const [ledgerToDateDraft, setLedgerToDateDraft] = useState('');
  const [ledgerPeriod, setLedgerPeriod] = useState<LedgerPeriod>('all');
  const [ledgerMonthYear, setLedgerMonthYear] = useState('');
  const [isLedgerMonthPickerOpen, setIsLedgerMonthPickerOpen] = useState(false);
  const [isDriverSettlementOpen, setIsDriverSettlementOpen] = useState(false);
  const [driverSettlementDriver, setDriverSettlementDriver] = useState('');
  const [driverSettlementPeriodMode, setDriverSettlementPeriodMode] = useState<'month' | 'date'>('date');
  const [driverSettlementMonth, setDriverSettlementMonth] = useState(() => dateInputValue(new Date()).slice(0, 7));
  const [isDriverSettlementMonthPickerOpen, setIsDriverSettlementMonthPickerOpen] = useState(false);
  const [driverSettlementFrom, setDriverSettlementFrom] = useState('');
  const [driverSettlementTo, setDriverSettlementTo] = useState('');
  const [driverSettlementFromDraft, setDriverSettlementFromDraft] = useState('');
  const [driverSettlementToDraft, setDriverSettlementToDraft] = useState('');
  const driverSettlementFromPickerRef = useRef<HTMLInputElement>(null);
  const driverSettlementToPickerRef = useRef<HTMLInputElement>(null);
  const [monthlySummaryMonth, setMonthlySummaryMonth] = useState(() => dateInputValue(new Date()).slice(0, 7));
  const [isMonthlySummaryCollapsed, setIsMonthlySummaryCollapsed] = useState(true);
  const [isLedgerSheetCollapsed, setIsLedgerSheetCollapsed] = useState(true);
  const [isDeviationSummaryCollapsed, setIsDeviationSummaryCollapsed] = useState(true);
  const [isFollowupSummaryCollapsed, setIsFollowupSummaryCollapsed] = useState(true);
  const [followupSummaryMonth, setFollowupSummaryMonth] = useState(() => dateInputValue(new Date()).slice(0, 7));
  const [isFollowupSummaryMonthPickerOpen, setIsFollowupSummaryMonthPickerOpen] = useState(false);
  const [deviationSummaryVehicleId, setDeviationSummaryVehicleId] = useState('all');
  const [selectedDeviationMonths, setSelectedDeviationMonths] = useState<string[]>([]);
  const [isDeviationSummaryPickerOpen, setIsDeviationSummaryPickerOpen] = useState(false);
  const [isMonthlySummaryPickerOpen, setIsMonthlySummaryPickerOpen] = useState(false);
  const monthlySummaryMonthSelectedRef = useRef(false);
  const [vehicleDialog, setVehicleDialog] = useState(false);
  const [editingVehicleId, setEditingVehicleId] = useState<string | null>(null);
  const vehicleImportRef = useRef<HTMLInputElement>(null);
  const allVehiclesImportRef = useRef<HTMLInputElement>(null);
  const [isImportingVehicles, setIsImportingVehicles] = useState(false);
  const [isImportingAllVehicleData, setIsImportingAllVehicleData] = useState(false);
  const vehicleHistoryImportRef = useRef<HTMLInputElement>(null);
  const [isImportingVehicleHistory, setIsImportingVehicleHistory] = useState(false);
  const [entryDialog, setEntryDialog] = useState(false);
  const [entryFormError, setEntryFormError] = useState('');
  const entryDatePickerRef = useRef<HTMLInputElement>(null);
  const [maintenanceDialog, setMaintenanceDialog] = useState(false);
  const [maintenanceForm, setMaintenanceForm] = useState({ vehicleId: '', amount: '', amountDraft: '' });
  const [maintenanceVehicleSearch, setMaintenanceVehicleSearch] = useState('');
  const currentEntryDialogRef = useRef(entryDialog);
  currentEntryDialogRef.current = entryDialog;
  const [isFuelSectionExpanded, setIsFuelSectionExpanded] = useState(false);
  const [isExpenseSectionExpanded, setIsExpenseSectionExpanded] = useState(false);
  const [editingEntryId, setEditingEntryId] = useState<string | null>(null);
  const [allowanceRateDraft, setAllowanceRateDraft] = useState(String(DEFAULT_ALLOWANCE_PRICE));
  const [isInstallAvailable, setIsInstallAvailable] = useState(() => Boolean(getDeferredPrompt()));
  const [isAppInstalled, setIsAppInstalled] = useState(() => window.matchMedia('(display-mode: standalone)').matches);
  const [notice, setNotice] = useState('');
  const [noticeIsError, setNoticeIsError] = useState(false);
  useEffect(() => {
    if (!notice) return undefined;
    const timeout = window.setTimeout(() => {
      setNotice('');
      setNoticeIsError(false);
    }, 3000);
    return () => window.clearTimeout(timeout);
  }, [notice]);
  const [deviationNotifications, setDeviationNotifications] = useState<FleetNotification[]>([]);
  const [isDeviationNotificationPanelOpen, setIsDeviationNotificationPanelOpen] = useState(false);
  const [expandedNotificationIds, setExpandedNotificationIds] = useState<Set<string>>(() => new Set());
  const [isMarkingAllNotificationsRead, setIsMarkingAllNotificationsRead] = useState(false);
  const [refreshingNotificationId, setRefreshingNotificationId] = useState<string | null>(null);
  const dismissedDeviationNotificationsRef = useRef(new Set<string>());
  const knownNotificationIdsRef = useRef(new Set<string>());
  const hasLoadedNotificationSnapshotRef = useRef(false);
  const notificationAudioContextRef = useRef<AudioContext | null>(null);
  const notificationSoundEnabledRef = useRef(false);
  const pendingDeviationNavigationRef = useRef<DeviationNotification | null>(null);
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [pendingDeleteEntry, setPendingDeleteEntry] = useState<FuelEntry | null>(null);
  const [deleteConfirmationText, setDeleteConfirmationText] = useState('');
  const [isClearEntriesConfirmationOpen, setIsClearEntriesConfirmationOpen] = useState(false);
  const [clearEntriesConfirmationPassword, setClearEntriesConfirmationPassword] = useState('');
  const [clearEntriesMode, setClearEntriesMode] = useState<'all' | 'month'>('all');
  const [clearEntriesMonth, setClearEntriesMonth] = useState(() => dateInputValue(new Date()).slice(0, 7));
  const [isClearEntriesMonthPickerOpen, setIsClearEntriesMonthPickerOpen] = useState(false);
  const [isClearingAllEntries, setIsClearingAllEntries] = useState(false);
  const [isRateEditorOpen, setIsRateEditorOpen] = useState(false);
  const [isRateCardExpanded, setIsRateCardExpanded] = useState(false);
  const [deviationNotes, setDeviationNotes] = useState<DeviationNotesByMonth>(readDeviationNotes);
  const [editingDeviationNote, setEditingDeviationNote] = useState<{ vehicleId: string; month: string } | null>(null);
  const [deviationNoteDraft, setDeviationNoteDraft] = useState('');
  const [vehicleForm, setVehicleForm] = useState({ name: '', vehicleType: 'dmax' as VehicleType, plate: '', driver: '', openingOdometer: '', kmPerLiter: '20', distanceBaseKm: '105' });
  const [entryForm, setEntryForm] = useState<EntryForm>(() => initialEntryDraft?.form ?? createEmptyEntryForm());
  const [routeDateFilter, setRouteDateFilter] = useState(today);
  const [routeDateFilterDraft, setRouteDateFilterDraft] = useState(() => formatEntryDateInput(today()));
  const routeDatePickerRef = useRef<HTMLInputElement>(null);
  const [routeDailyDrafts, setRouteDailyDrafts] = useState<Record<string, RouteDailyDraft>>({});
  const [expandedRouteVehicleIds, setExpandedRouteVehicleIds] = useState<Set<string>>(() => new Set());
  const [savingRouteVehicleIds, setSavingRouteVehicleIds] = useState<Set<string>>(() => new Set());
  const [routeVehicleDialog, setRouteVehicleDialog] = useState<{ vehicleId: string | null; isDailyOnly: boolean; plate: string; name: string } | null>(null);
  const [routeVehicleDialogError, setRouteVehicleDialogError] = useState('');
  const [isExportingRoutes, setIsExportingRoutes] = useState(false);
  const [entryPlateSearch, setEntryPlateSearch] = useState(initialEntryDraft?.entryPlateSearch ?? '');
  const [fuelAmountDraft, setFuelAmountDraft] = useState(initialEntryDraft?.fuelAmountDraft ?? '');
  const [expenseAmountDrafts, setExpenseAmountDrafts] = useState<ExpenseInputs>(() => initialEntryDraft?.expenseAmountDrafts ?? emptyExpenseInputs());
  const [recentlyAddedAmount, setRecentlyAddedAmount] = useState<{ target: string; token: number } | null>(null);
  const amountFlashTokenRef = useRef(0);
  const amountFlashTimeoutRef = useRef<number | null>(null);
  const [expandedExpenseKey, setExpandedExpenseKey] = useState<keyof ExpenseBreakdown | null>(null);
  const canManageUsers = currentUser.role === 'admin' && Boolean(firebaseAuth?.currentUser);
  const canEditFleet = currentUser.role === 'admin' && Boolean(firebaseAuth?.currentUser);
  const canAccessEntries = currentUser.role !== 'viewer';
  const isClearEntriesPasswordValid = clearEntriesConfirmationPassword.trim().toLowerCase() === CLEAR_ENTRIES_CONFIRMATION_PASSWORD;
  const recordAudit = (action: FleetAuditAction, itemName: string) => {
    if (!canEditFleet) return;
    void recordFleetAuditEntry(currentUser, action, itemName).catch((error: unknown) => {
      console.error('Could not record fleet audit entry', error);
      setNoticeIsError(true);
      setNotice(`تم تنفيذ العملية، لكن تعذر تسجيلها في سجل العمليات: ${error instanceof Error ? error.message : 'خطأ غير معروف'}`);
    });
  };
  const notifyAccessDenied = () => {
    setNotice('هذا القسم مخصص للمسؤولين فقط. ليس لديك صلاحية للوصول');
    setNoticeIsError(true);
  };

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
    if (!isMobileActionsOpen) return undefined;
    const closeOnOutsideClick = (event: PointerEvent) => {
      if (event.target instanceof Node && !mobileActionsRef.current?.contains(event.target)) {
        setIsMobileActionsOpen(false);
      }
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setIsMobileActionsOpen(false);
    };
    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [isMobileActionsOpen]);

  useEffect(() => {
    setLedgerFromDateDraft(formatDateFilter(ledgerFilters.from));
  }, [ledgerFilters.from]);

  useEffect(() => {
    setLedgerToDateDraft(formatDateFilter(ledgerFilters.to));
  }, [ledgerFilters.to]);

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

  useEffect(() => () => {
    if (amountFlashTimeoutRef.current !== null) window.clearTimeout(amountFlashTimeoutRef.current);
  }, []);

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
    if (canAccessEntries || (tab !== 'entries' && !entryDialog)) return;
    setTab('home');
    setEntryDialog(false);
    notifyAccessDenied();
  }, [canAccessEntries, entryDialog, tab]);

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
          const denied = (!canAccessEntries && (previousRoute.tab === 'entries' || previousRoute.entryPage))
            || (!canManageUsers && previousRoute.tab === 'audit');
          setTab(denied ? 'home' : previousRoute.tab);
          setSelectedVehicleId(previousRoute.selectedVehicleId);
          setEntryDialog(canAccessEntries && (previousRoute.entryPage ?? false));
          if (denied) notifyAccessDenied();
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
      const denied = (!canAccessEntries && (previousRoute.tab === 'entries' || previousRoute.entryPage))
        || (!canManageUsers && previousRoute.tab === 'audit');
      setTab(denied ? 'home' : previousRoute.tab);
      setSelectedVehicleId(previousRoute.selectedVehicleId);
      setEntryDialog(canAccessEntries && (previousRoute.entryPage ?? false));
      if (denied) notifyAccessDenied();
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, [canAccessEntries, canManageUsers]);

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
    try {
      localStorage.setItem(entryColumnVisibilityKey, JSON.stringify(entryColumnVisibility));
    } catch (error) {
      console.error('Failed to save entry table column preferences', error);
      setNoticeIsError(true);
      setNotice('تعذر حفظ إعدادات أعمدة سجل التعبئة على هذا الجهاز.');
    }
  }, [entryColumnVisibility, entryColumnVisibilityKey]);

  useEffect(() => {
    try {
      localStorage.setItem(detailColumnVisibilityKey, JSON.stringify(detailColumnVisibility));
    } catch (error) {
      console.error('Failed to save vehicle detail column preferences', error);
      setNoticeIsError(true);
      setNotice('تعذر حفظ إعدادات أعمدة تفاصيل السيارات على هذا الجهاز.');
    }
  }, [detailColumnVisibility, detailColumnVisibilityKey]);

  useEffect(() => {
    if (!cloudLoaded || monthlySummaryMonthSelectedRef.current || data.entries.length === 0) return;
    const latestEntryMonth = data.entries.reduce((latest, entry) => entry.date.slice(0, 7) > latest ? entry.date.slice(0, 7) : latest, '');
    if (!latestEntryMonth) return;
    setMonthlySummaryMonth(latestEntryMonth);
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
      setSelectedDeviationMonths([pendingNotification.month]);
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
    setIsThemePickerOpen(false);
  };

  const chooseFontSize = (nextFontSize: FontSizeKey) => {
    setFontSizeKey(nextFontSize);
    localStorage.setItem(`${themeStorageKey}-font`, nextFontSize);
    setIsThemePickerOpen(false);
  };

  const renderThemeOptions = () => themeOptions.map((key) => {
    const preset = themePresets[key];
    return <button
      key={key}
      type="button"
      className={`fleet-theme-option${themeKey === key ? ' is-active' : ''}`}
      aria-pressed={themeKey === key}
      onClick={() => chooseTheme(key)}
    >{preset.label}</button>;
  });

  const renderFontSizePicker = () => <div className="fleet-font-picker">
    <strong>حجم الخط</strong>
    <div className="fleet-font-options">{(Object.keys(fontSizePresets) as FontSizeKey[]).map((key) => <button
      key={key}
      type="button"
      className={fontSizeKey === key ? 'is-active' : ''}
      aria-pressed={fontSizeKey === key}
      onClick={() => chooseFontSize(key)}
    >{fontSizePresets[key]}</button>)}</div>
  </div>;

  const loadOlderGpsLogs = async () => {
    if (!gpsLogCursor || isLoadingMoreGpsLogs) return;
    setIsLoadingMoreGpsLogs(true);
    setGpsLogsError('');
    try {
      const page = await loadMoreGpsLogs(gpsLogCursor);
      setGpsLogs((current) => {
        const logsById = new Map(current.map((log) => [log.id, log]));
        page.logs.forEach((log) => logsById.set(log.id, log));
        return [...logsById.values()].sort((first, second) => second.timestamp.localeCompare(first.timestamp));
      });
      setGpsLogCursor(page.cursor);
      setHasMoreGpsLogs(page.hasMore);
    } catch (error) {
      console.error('Could not load older GPS logs', error);
      setGpsLogsError(error instanceof Error ? error.message : 'تعذر تحميل سجلات GPS الأقدم');
    } finally {
      setIsLoadingMoreGpsLogs(false);
    }
  };

  const routeBoardVehicles: RouteBoardVehicle[] = [
    ...data.vehicles.map((vehicle) => {
      const savedDescription = routeEntries.find((entry) => entry.isDailyVehicle && entry.vehicleId === vehicle.id);
      return {
        id: vehicle.id,
        fleetVehicleId: vehicle.id,
        name: savedDescription?.vehicleName || vehicle.name,
        plate: savedDescription?.plate || vehicle.plate,
        isDailyOnly: false,
      };
    }),
    ...routeEntries.filter((entry) => entry.isDailyVehicle && entry.vehicleId === null).map((entry) => ({
      id: entry.id,
      fleetVehicleId: null,
      name: entry.vehicleName,
      plate: entry.plate,
      isDailyOnly: true,
    })),
  ];

  const saveRouteVehicleMetadata = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!canEditFleet || !routeVehicleDialog) return;
    const plate = routeVehicleDialog.plate.trim();
    const name = routeVehicleDialog.name.trim();
    if (!plate || !name) {
      setRouteVehicleDialogError('أدخل رقم السيارة والوصف.');
      return;
    }
    const duplicateVehicle = routeBoardVehicles.some((vehicle) => (
      vehicle.id !== routeVehicleDialog.vehicleId
      && vehicle.plate.trim().toLocaleLowerCase('ar-EG') === plate.toLocaleLowerCase('ar-EG')
    ));
    if (duplicateVehicle) {
      setRouteVehicleDialogError('رقم السيارة موجود بالفعل في بيان الحركة.');
      return;
    }

    const dailyOnly = routeVehicleDialog.isDailyOnly;
    const markerId = dailyOnly
      ? routeVehicleDialog.vehicleId || `route-vehicle:${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
      : `route-vehicle-meta:${routeVehicleDialog.vehicleId}`;
    const marker: RouteEntry = {
      id: dailyOnly ? markerId : `route-vehicle-meta:${routeVehicleDialog.vehicleId}`,
      date: '__vehicle__',
      vehicleId: dailyOnly ? null : routeVehicleDialog.vehicleId,
      isDailyVehicle: true,
      vehicleName: name,
      plate,
      departure: '',
      route: '',
      representative: '',
      driver: '',
    };
    setRouteEntries((current) => {
      const existing = current.findIndex((entry) => (
        dailyOnly ? entry.id === marker.id : entry.isDailyVehicle && entry.vehicleId === routeVehicleDialog.vehicleId
      ));
      return existing >= 0
        ? current.map((entry, index) => index === existing ? marker : entry)
        : [...current, marker];
    });
    setRouteVehicleDialog(null);
    setRouteVehicleDialogError('');
    setNoticeIsError(false);
    setNotice(dailyOnly ? 'تمت إضافة السيارة إلى بيان الحركة فقط.' : 'تم تحديث وصف السيارة في بيان الحركة فقط.');
  };

  const saveRouteVehicle = async (vehicle: RouteBoardVehicle) => {
    const parsedDate = parseDriverSettlementDate(routeDateFilterDraft);
    if (!canEditFleet || !routesCloudLoaded || !parsedDate || parsedDate !== routeDateFilter || savingRouteVehicleIds.has(vehicle.id)) return;
    setSavingRouteVehicleIds((current) => new Set(current).add(vehicle.id));
    const key = `${routeDateFilter}|${vehicle.id}`;
    const draft = routeDailyDrafts[key] || { departure: '', route: '', representative: '', driver: '' };
    const existing = latestRouteEntriesRef.current.find((entry) => !entry.isDailyVehicle
      && entry.date === routeDateFilter
      && (vehicle.fleetVehicleId
        ? entry.vehicleId === vehicle.fleetVehicleId
        : entry.dailyVehicleId === vehicle.id || (!entry.vehicleId && entry.plate.trim() === vehicle.plate.trim())));
    const savedEntry: RouteEntry = {
      id: existing?.id || `r${routeDateFilter}-${vehicle.id}`,
      date: routeDateFilter,
      vehicleId: vehicle.fleetVehicleId,
      isDailyVehicle: false,
      dailyVehicleId: vehicle.isDailyOnly ? vehicle.id : undefined,
      vehicleName: vehicle.name,
      plate: vehicle.plate,
      departure: draft.departure.trim(),
      route: draft.route.trim(),
      representative: draft.representative.trim(),
      driver: draft.driver.trim(),
    };
    const nextEntries = existing
      ? latestRouteEntriesRef.current.map((entry) => entry.id === existing.id ? savedEntry : entry)
      : [...latestRouteEntriesRef.current, savedEntry];
    latestRouteEntriesRef.current = nextEntries;
    setRouteEntries(nextEntries);
    setNoticeIsError(false);
    try {
      localStorage.setItem(ROUTE_STORAGE_KEY, JSON.stringify(nextEntries));
      if (isFirebaseConfigured) {
        if (!routesCloudAvailable) throw new Error('لا يتوفر اتصال Firestore حاليًا');
        await saveCloudRouteEntries(nextEntries);
      }
      setNotice(`${isFirebaseConfigured ? 'تم حفظ' : 'تم حفظ محليًا'} بيان حركة سيارة ${vehicle.plate} بتاريخ ${formatDateArabic(routeDateFilter)}`);
    } catch (error) {
      console.error('Could not save daily vehicle route to Firestore', error);
      setNoticeIsError(true);
      setNotice(`تعذر حفظ بيان حركة السيارة في Firestore: ${error instanceof Error ? error.message : 'خطأ غير معروف'}`);
    } finally {
      setSavingRouteVehicleIds((current) => {
        const next = new Set(current);
        next.delete(vehicle.id);
        return next;
      });
    }
  };

  const exportRouteEntries = async () => {
    setIsExportingRoutes(true);
    try {
      const ExcelJS = (await import('exceljs')).default;
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'مراقبة سيارات شركة أبو غالي';
      workbook.subject = 'خط سير السيارات اليومي';
      const worksheet = workbook.addWorksheet(routeDateFilter, { views: [{ rightToLeft: true, showGridLines: false, state: 'frozen', ySplit: 2 }] });
      worksheet.columns = [
        { width: 8 }, { width: 16 }, { width: 19 }, { width: 34 },
        { width: 36 }, { width: 24 }, { width: 24 },
      ];
      worksheet.mergeCells('A1:G1');
      worksheet.getCell('A1').value = `خط سير السيارات اليومي - ${formatDateArabic(routeDateFilter)}`;
      worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      worksheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9A682C' } };
      worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(1).height = 32;
      const header = ['م', 'الوصف', 'رقم السيارة', 'نقطة الإنطلاق', 'خط السير', 'المندوب', 'السائق'];
      const headerRow = worksheet.addRow(header);
      headerRow.height = 26;
      headerRow.eachCell((cell) => {
        cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5D4025' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      routeBoardVehicles.forEach((vehicle, index) => {
        const entry = routeEntriesForDate.find((route) => vehicle.fleetVehicleId
          ? route.vehicleId === vehicle.fleetVehicleId
          : route.dailyVehicleId === vehicle.id || (!route.vehicleId && route.plate.trim() === vehicle.plate.trim()));
        const row = worksheet.addRow([index + 1, vehicle.name, vehicle.plate, entry?.departure || '', entry?.route || '', entry?.representative || '', entry?.driver || '']);
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
    .sort((first, second) => first.plate.localeCompare(second.plate, 'ar'));
  const savedRouteVehicleCount = routeBoardVehicles.filter((vehicle) => routeEntriesForDate.some((entry) => (
    vehicle.fleetVehicleId
      ? entry.vehicleId === vehicle.fleetVehicleId
      : entry.dailyVehicleId === vehicle.id || (!entry.vehicleId && entry.plate.trim() === vehicle.plate.trim())
  ))).length;
  const routeDateFilterInvalid = !parseDriverSettlementDate(routeDateFilterDraft);

  const saveOpeningOdometer = (vehicle: Vehicle, entriesCount: number) => {
    if (!canEditFleet) return;
    const value = Number(openingOdometerDrafts[vehicle.id] ?? vehicle.openingOdometer);
    if (!Number.isFinite(value) || value < 0) {
      setNotice('أدخل قراءة سابقة صحيحة');
      return;
    }
    setData((current) => ({ ...current, vehicles: current.vehicles.map((item) => item.id === vehicle.id ? { ...item, openingOdometer: value } : item) }));
    setOpeningOdometerDrafts((current) => ({ ...current, [vehicle.id]: String(value) }));
    recordAudit('edit', `القراءة السابقة لسيارة ${vehicle.plate}`);
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
        recordAudit('edit', `بيانات الزائر ${guestDraft.employeeName.trim()}`);
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
      recordAudit('edit', `حالة الزائر ${guest.employeeName}`);
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
      recordAudit('delete', `الزائر ${guest.employeeName}`);
      setNotice('تم حذف الزائر');
      setNoticeIsError(false);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'تعذر حذف الزائر');
      setNoticeIsError(true);
    }
  };

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    if (cloudLoaded && canEditFleet) {
      const nextFingerprint = fleetDataFingerprint(data);
      const previousFingerprint = lastSyncedFleetFingerprintRef.current;
      if (previousFingerprint === nextFingerprint) return;
      void saveCloudFleetDataStrict(data).then(() => {
        lastSyncedFleetFingerprintRef.current = nextFingerprint;
      }).catch((error: unknown) => {
        console.error('Could not save fleet data to the cloud', error);
        setNoticeIsError(true);
        setNotice(error instanceof Error && error.message.includes('Firestore')
          ? 'تم حفظ البيانات محليًا، لكن تعذرت مزامنتها.'
          : 'تعذرت مزامنة البيانات.');
      });
    }
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
      const resolvedData = cloudData ? normalizeFleetData(cloudData) : latestFleetDataRef.current;
      lastSyncedFleetFingerprintRef.current = fleetDataFingerprint(cloudData ?? resolvedData);
      if (cloudData) setData(resolvedData);
      setCloudLoaded(true);
    });
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!canManageUsers) {
      setAuditEntries([]);
      setAuditLogError('');
      return undefined;
    }
    return subscribeToFleetAuditLog(
      setAuditEntries,
      (error) => {
        setAuditLogError(error.message);
        setNoticeIsError(true);
        setNotice(`تعذر تحميل سجل العمليات: ${error.message}`);
      },
    );
  }, [canManageUsers]);

  useEffect(() => {
    if (tab !== 'audit' || canManageUsers) return;
    setTab('home');
    notifyAccessDenied();
  }, [canManageUsers, tab]);

  useEffect(() => {
    if (data.vehicles.some((vehicle) => vehicle.id === ledgerSummaryVehicleId)) return;
    setLedgerSummaryVehicleId(data.vehicles[0]?.id ?? '');
  }, [data.vehicles, ledgerSummaryVehicleId]);

  useEffect(() => {
    dismissedDeviationNotificationsRef.current = new Set<string>();
    knownNotificationIdsRef.current = new Set<string>();
    hasLoadedNotificationSnapshotRef.current = false;
    setDeviationNotifications([]);
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
      const validNotifications = notifications.filter((notification) => (
        notification
        && typeof notification.id === 'string'
        && (isFleetUpdateNotification(notification)
          ? notification.source === 'fuel-entry' && typeof notification.title === 'string' && typeof notification.summary === 'string'
          : typeof notification.vehicleId === 'string' && typeof notification.month === 'string' && typeof notification.deviation === 'number')
      ));
      const isInitialSnapshot = !hasLoadedNotificationSnapshotRef.current;
      hasLoadedNotificationSnapshotRef.current = true;
      const freshNotifications = validNotifications.filter(({ id }) => !knownNotificationIdsRef.current.has(id));
      validNotifications.forEach(({ id }) => knownNotificationIdsRef.current.add(id));
      if (validNotifications.length) {
        setDeviationNotifications((current) => {
          const existingIds = new Set(current.map(({ id }) => id));
          return [...current, ...validNotifications.filter(({ id }) => !existingIds.has(id))];
        });
      }
      if (isInitialSnapshot || !freshNotifications.length || !notificationSoundEnabledRef.current) return;
      const newDeviations = freshNotifications.filter((notification): notification is DeviationNotification => !isFleetUpdateNotification(notification));
      const highestDeviation = Math.max(0, ...newDeviations.map(({ deviation }) => deviation));
      if (highestDeviation <= 100) return;
      const severity = highestDeviation > 250 ? 'danger' : 'warning';
      const audioContext = notificationAudioContextRef.current;
      if (audioContext?.state === 'running') {
        const oscillator = audioContext.createOscillator();
        const gain = audioContext.createGain();
        oscillator.type = 'sine';
        oscillator.frequency.value = severity === 'danger' ? 880 : 660;
        gain.gain.setValueAtTime(severity === 'danger' ? 0.16 : 0.08, audioContext.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, audioContext.currentTime + (severity === 'danger' ? 0.45 : 0.2));
        oscillator.connect(gain);
        gain.connect(audioContext.destination);
        oscillator.start();
        oscillator.stop(audioContext.currentTime + (severity === 'danger' ? 0.45 : 0.2));
      }
      if (severity === 'danger' && typeof navigator.vibrate === 'function') navigator.vibrate([250, 120, 250]);
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
    if (tab !== 'routes') return undefined;
    if (!isFirebaseConfigured || !firebaseAuth) {
      setGpsLogsError('خدمة Firebase غير مهيأة لتحميل سجلات GPS');
      setIsLoadingGpsLogs(false);
      return undefined;
    }

    let unsubscribeLogs: (() => void) | undefined;
    setGpsLogs([]);
    setGpsLogCursor(null);
    setHasMoreGpsLogs(false);
    setGpsLogsError('');
    setIsLoadingGpsLogs(true);
    const unsubscribeAuth = onAuthStateChanged(firebaseAuth, (user) => {
      unsubscribeLogs?.();
      if (!user) {
        setIsLoadingGpsLogs(false);
        setGpsLogsError('سجّل الدخول إلى Firebase لعرض سجلات GPS');
        return;
      }
      unsubscribeLogs = subscribeToGpsLogs(
        (page) => {
          const firstPageIds = new Set(page.logs.map((log) => log.id));
          setGpsLogs((current) => {
            const olderPages = current.filter((log) => !firstPageIds.has(log.id));
            const logsById = new Map([...page.logs, ...olderPages].map((log) => [log.id, log]));
            return [...logsById.values()].sort((first, second) => second.timestamp.localeCompare(first.timestamp));
          });
          setGpsLogCursor(page.cursor);
          setHasMoreGpsLogs(page.hasMore);
          setIsLoadingGpsLogs(false);
          setGpsLogsError('');
        },
        (error) => {
          console.error('Could not subscribe to GPS logs', error);
          setIsLoadingGpsLogs(false);
          setGpsLogsError(error.message || 'تعذر تحميل سجلات GPS');
        },
      );
    });
    return () => {
      unsubscribeAuth();
      unsubscribeLogs?.();
    };
  }, [tab]);

  useEffect(() => {
    setAllowanceRateDraft(String(fuelAllowanceRate(data)));
  }, [data.allowancePricePerLiter]);

  useEffect(() => {
    setIsRateSummaryExpanded(true);
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
    const actualLiters = entries.reduce((total, entry) => {
      const rate = entryAllowancePrice(entry, fuelAllowanceRate(data));
      return total + (rate > 0 ? entryFuelCost(entry) / rate : 0);
    }, 0);
    const kmPerLiter = actualLiters > 0 ? distance / actualLiters : 0;
    const lastOdometer = entries.reduce((latest, entry) => Math.max(latest, entry.currentOdometer), vehicle.openingOdometer);
    return { vehicle, entries, distance, liters, expected, deviation, cost, kmPerLiter, lastOdometer };
  }), [data]);

  const calculateFuelLiters = (vehicleId: string, startOdometer: number, currentOdometer: number) => {
    const vehicle = data.vehicles.find((item) => item.id === vehicleId);
    if (!vehicle) return 0;
    return litersForDistance(vehicle, currentOdometer - startOdometer);
  };
  const calculateAllowanceValue = (vehicleId: string, liters: number, rate = fuelAllowanceRate(data)) => {
    return allowanceCost(liters, rate);
  };
  const allowanceValueForEntry = (entry: FuelEntry) => entryAllowanceValue(entry, fuelAllowanceRate(data));

  const totals = summaries.reduce((total, summary) => ({
    distance: total.distance + summary.distance,
    liters: total.liters + summary.liters,
    expected: total.expected + summary.expected,
    deviation: total.deviation + summary.deviation,
    cost: total.cost + summary.cost,
  }), { distance: 0, liters: 0, expected: 0, deviation: 0, cost: 0 });

  const selectedSummary = summaries.find((summary) => summary.vehicle.id === selectedVehicleId);
  const selectedDetailColumnVisibility = { ...defaultDetailColumnVisibility, ...detailColumnVisibility };
  const visibleDetailColumns = detailColumns.filter(({ key }) => selectedDetailColumnVisibility[key] !== false);
  const selectedEntries = data.entries
    .filter((entry) => entry.vehicleId === selectedVehicleId)
    .sort((a, b) => a.date.localeCompare(b.date));
  const selectedRateMovements = selectedEntries.map((entry) => {
    const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
    const price = entryAllowancePrice(entry, fuelAllowanceRate(data));
    const actualLiters = price > 0 ? entryFuelCost(entry) / price : 0;
    return { entry, distance, actualLiters, rate: actualLiters > 0 ? distance / actualLiters : null };
  });
  const selectedRateTotals = selectedRateMovements.reduce((totals, movement) => ({
    distance: totals.distance + movement.distance,
    actualLiters: totals.actualLiters + movement.actualLiters,
  }), { distance: 0, actualLiters: 0 });
  const selectedMovementRates = selectedRateMovements
    .map((movement) => movement.rate)
    .filter((rate): rate is number => rate !== null && Number.isFinite(rate));
  const selectedAverageRate = selectedRateTotals.actualLiters > 0
    ? selectedRateTotals.distance / selectedRateTotals.actualLiters
    : 0;
  const selectedCalculatedBaseDistance = selectedSummary
    ? selectedAverageRate * vehicleFuelRate(selectedSummary.vehicle)
    : 0;
  const selectedBestRate = selectedMovementRates.reduce((best, rate) => Math.max(best, rate), 0);
  const selectedWorstRate = selectedMovementRates.length
    ? selectedMovementRates.reduce((worst, rate) => Math.min(worst, rate), Number.POSITIVE_INFINITY)
    : 0;
  const comparisonToday = new Date();
  const comparisonMonthStart = new Date(comparisonToday.getFullYear(), comparisonToday.getMonth(), 1);
  const lastMonthStart = new Date(comparisonToday.getFullYear(), comparisonToday.getMonth() - 1, 1);
  const monthBeforeLastStart = new Date(comparisonToday.getFullYear(), comparisonToday.getMonth() - 2, 1);
  const lastMonthPeriodStart = dateInputValue(lastMonthStart);
  const lastMonthPeriodEnd = dateInputValue(new Date(comparisonMonthStart.getFullYear(), comparisonMonthStart.getMonth(), 0));
  const monthBeforeLastPeriodStart = dateInputValue(monthBeforeLastStart);
  const monthBeforeLastPeriodEnd = dateInputValue(new Date(lastMonthStart.getFullYear(), lastMonthStart.getMonth(), 0));
  const comparisonMonthLabel = (date: Date) => new Intl.DateTimeFormat('ar-EG', { month: 'long', year: 'numeric' }).format(date);
  const calculatePeriodFuelRate = (startDate: string, endDate: string) => {
    const monthMovements = selectedRateMovements.filter(({ entry }) => entry.date >= startDate && entry.date <= endDate);
    const totals = monthMovements.reduce((sum, movement) => ({
      distance: sum.distance + movement.distance,
      actualLiters: sum.actualLiters + movement.actualLiters,
    }), { distance: 0, actualLiters: 0 });
    return { ...totals, rate: totals.actualLiters > 0 ? totals.distance / totals.actualLiters : 0 };
  };
  const lastMonthRateSummary = calculatePeriodFuelRate(lastMonthPeriodStart, lastMonthPeriodEnd);
  const monthBeforeLastRateSummary = calculatePeriodFuelRate(monthBeforeLastPeriodStart, monthBeforeLastPeriodEnd);
  const monthlyRateChange = monthBeforeLastRateSummary.rate > 0
    ? ((lastMonthRateSummary.rate - monthBeforeLastRateSummary.rate) / monthBeforeLastRateSummary.rate) * 100
    : null;
  const chartRateMovements = selectedRateMovements
    .filter((movement): movement is typeof movement & { rate: number } => movement.rate !== null && Number.isFinite(movement.rate))
    .slice(-12);
  const chartAverageRate = chartRateMovements.length
    ? chartRateMovements.reduce((total, movement) => total + movement.rate, 0) / chartRateMovements.length
    : 0;
  const chartMaximumRate = Math.max(12, ...chartRateMovements.map(({ rate }) => rate));
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
    const pricePerLiter = entryAllowancePrice(entry, fuelAllowanceRate(data));
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
      allowance: totals.allowance + entryAllowanceValue(row.entry, fuelAllowanceRate(data), liters),
      actual: totals.actual + row.actualFuelValue,
      deviation: totals.deviation + row.deviation,
      actualLiters: totals.actualLiters + (entryAllowancePrice(row.entry, fuelAllowanceRate(data)) > 0
        ? row.actualFuelValue / entryAllowancePrice(row.entry, fuelAllowanceRate(data))
        : 0),
    };
  }, { allowance: 0, actual: 0, deviation: 0, actualLiters: 0 });
  const totalKmPerLiter = selectedFuelTotals.actualLiters > 0
    ? selectedSummary.distance / selectedFuelTotals.actualLiters
    : 0;

  const normalizedLedgerSearchKey = normalizeVehicleSearchKey(search);
  const matchingLedgerVehicleIds = new Set(filterVehicles(search, data.vehicles).map(({ id }) => id));
  const filteredEntries = data.entries
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date))
    .filter((entry) => {
      const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
      const matchesVehicle = matchingLedgerVehicleIds.has(entry.vehicleId);
      const searchable = `${entry.documentNumber ?? ''} ${vehicle?.driver ?? ''} ${entry.driver ?? ''} ${entry.note}`;
      return (ledgerFilters.vehicleId === 'all' || entry.vehicleId === ledgerFilters.vehicleId)
        && (!ledgerFilters.from || entry.date >= ledgerFilters.from)
        && (!ledgerFilters.to || entry.date <= ledgerFilters.to)
        && (!normalizedLedgerSearchKey || matchesVehicle || matchesVehicleSearchText(search, searchable));
    });
  const visibleEntryColumns = entryColumns.filter(({ key }) => entryColumnVisibility[key]);
  const latestEntries = data.entries
    .slice()
    .sort((a, b) => {
      const firstTimestamp = Date.parse(a.createdAt || `${a.date}T00:00:00.000Z`);
      const secondTimestamp = Date.parse(b.createdAt || `${b.date}T00:00:00.000Z`);
      return (Number.isFinite(secondTimestamp) ? secondTimestamp : 0) - (Number.isFinite(firstTimestamp) ? firstTimestamp : 0);
    })
    .slice(0, 5);

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
    setLedgerSummaryVehicleId(data.vehicles[0]?.id ?? '');
    setLedgerPeriod('all');
    setLedgerMonthYear('');
    setIsLedgerMonthPickerOpen(false);
  };

  const ledgerEntries = useMemo(() => {
    const matchingVehicles = filterVehicles(ledgerFilters.search, data.vehicles);
    return data.entries
      .slice()
      .sort((a, b) => a.date.localeCompare(b.date))
      .filter((entry) => {
        const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
        const matchesVehicle = matchingVehicles.some((item) => item.id === entry.vehicleId);
        const searchable = `${entry.documentNumber ?? ''} ${vehicle?.driver ?? ''} ${entry.driver ?? ''} ${entry.note}`;
        return entry.vehicleId === ledgerSummaryVehicleId
          && (!ledgerFilters.from || entry.date >= ledgerFilters.from)
          && (!ledgerFilters.to || entry.date <= ledgerFilters.to)
          && (!normalizeVehicleSearchKey(ledgerFilters.search) || matchesVehicle || matchesVehicleSearchText(ledgerFilters.search, searchable));
      });
  }, [data.entries, data.vehicles, ledgerFilters, ledgerSummaryVehicleId]);

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
  const followupSummaryEntries = useMemo(
    () => data.entries.filter((entry) => entry.date.slice(0, 7) === followupSummaryMonth),
    [data.entries, followupSummaryMonth],
  );
  const followupSummaryTotals = followupSummaryEntries.reduce((total, entry) => {
    const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
    const liters = litersForEntry(entry, vehicle);
    const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
    return {
      distance: total.distance + distance,
      liters: total.liters + liters,
      cost: total.cost + entryTotalCost(entry),
      deviation: total.deviation + entryDeviation(entry, fuelAllowanceRate(data), liters),
    };
  }, { distance: 0, liters: 0, cost: 0, deviation: 0 });
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
      allowance: entries.reduce((total, entry) => total + entryAllowanceValue(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)), 0),
      deviation: entries.reduce((total, entry) => total + entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)), 0),
      distance: entries.reduce((total, entry) => total + Math.max(0, entry.currentOdometer - entry.startOdometer), 0),
      liters: entries.reduce((total, entry) => total + litersForEntry(entry, vehicle), 0),
      actualLiters: entries.reduce((total, entry) => {
        const rate = entryAllowancePrice(entry, fuelAllowanceRate(data));
        return total + (rate > 0 ? entryFuelCost(entry) / rate : 0);
      }, 0),
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
    const entries = data.entries.filter((entry) => entry.vehicleId === vehicle.id
      && (!selectedDeviationMonths.length || selectedDeviationMonths.some((month) => entry.date.startsWith(month))));
    const deviation = entries.reduce((total, entry) => total + entryDeviation(entry, fuelAllowanceRate(data), litersForEntry(entry, vehicle)), 0);
    return { vehicle, entries, deviation };
  }).filter((summary) => summary.deviation > 250), [data, selectedDeviationMonths]);
  const hiddenDeviationVehicleIds = new Set(data.hiddenDeviationVehicleIds || []);
  const displayedDeviationSummaries = deviationSummaries.filter((summary) => canEditFleet || !hiddenDeviationVehicleIds.has(summary.vehicle.id));
  const visibleDeviationSummaries = displayedDeviationSummaries.filter((summary) => deviationSummaryVehicleId === 'all' || summary.vehicle.id === deviationSummaryVehicleId);
  const visibleDeviationTotal = visibleDeviationSummaries.reduce((total, summary) => total + summary.deviation, 0);
  const driverSettlementDrivers = useMemo(() => [...new Set(data.entries
    .map((entry) => entry.driver?.trim() || data.vehicles.find((vehicle) => vehicle.id === entry.vehicleId)?.driver.trim() || '')
    .filter(Boolean))].sort((left, right) => left.localeCompare(right, 'ar')), [data.entries, data.vehicles]);
  const parsedDriverSettlementFrom = parseDriverSettlementDate(driverSettlementFromDraft);
  const parsedDriverSettlementTo = parseDriverSettlementDate(driverSettlementToDraft);
  const driverSettlementDateInputInvalid = driverSettlementPeriodMode === 'date' && Boolean(
    (driverSettlementFromDraft.trim() && !parsedDriverSettlementFrom)
    || (driverSettlementToDraft.trim() && !parsedDriverSettlementTo),
  );
  const driverSettlementEntries = useMemo(() => data.entries.filter((entry) => {
    if (driverSettlementDateInputInvalid) return false;
    const driver = entry.driver?.trim() || data.vehicles.find((vehicle) => vehicle.id === entry.vehicleId)?.driver.trim() || '';
    if (driver !== driverSettlementDriver) return false;
    if (driverSettlementPeriodMode === 'month') return entry.date.slice(0, 7) === driverSettlementMonth;
    return (!driverSettlementFrom || entry.date >= driverSettlementFrom)
      && (!driverSettlementTo || entry.date <= driverSettlementTo);
  }), [data.entries, data.vehicles, driverSettlementDriver, driverSettlementPeriodMode, driverSettlementMonth, driverSettlementFrom, driverSettlementTo, driverSettlementDateInputInvalid]);
  const driverSettlementTotals = driverSettlementEntries.reduce((totals, entry) => {
    const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
    const liters = litersForEntry(entry, vehicle);
    return {
      distance: totals.distance + Math.max(0, entry.currentOdometer - entry.startOdometer),
      actual: totals.actual + entryFuelCost(entry),
      allowance: totals.allowance + entryAllowanceValue(entry, fuelAllowanceRate(data), liters),
      deviation: totals.deviation + entryDeviation(entry, fuelAllowanceRate(data), liters),
    };
  }, { distance: 0, actual: 0, allowance: 0, deviation: 0 });
  const driverSettlementDateRangeInvalid = driverSettlementPeriodMode === 'date' && (
    driverSettlementDateInputInvalid
    || Boolean(driverSettlementFrom && driverSettlementTo && driverSettlementFrom > driverSettlementTo)
  );
  const availableDeviationMonths = [...new Set(data.entries.map((entry) => entry.date.slice(0, 7)).filter((month) => /^\d{4}-\d{2}$/.test(month)))].sort((left, right) => right.localeCompare(left));
  const deviationMonthLabel = selectedDeviationMonths.length === 0
    ? 'كل الأشهر'
    : selectedDeviationMonths.length === 1
      ? deviationMonthDisplayName(selectedDeviationMonths[0])
      : `${selectedDeviationMonths.length} أشهر محددة`;
  const editingDeviationVehicle = editingDeviationNote
    ? data.vehicles.find((vehicle) => vehicle.id === editingDeviationNote.vehicleId)
    : undefined;
  const unreadNotificationCount = deviationNotifications.filter(({ id }) => !dismissedDeviationNotificationsRef.current.has(id)).length;
  const notificationListItems = useMemo<NotificationListItem[]>(() => {
    const dangerousNotifications = deviationNotifications.filter((notification): notification is DeviationNotification => (
      !isFleetUpdateNotification(notification) && notification.deviation > 250
    ));
    const plates = [...new Set(dangerousNotifications.map(({ plate }) => plate))];
    const groupedNotifications = plates.length > 1
      ? [...dangerousNotifications].sort((left, right) => right.createdAt.localeCompare(left.createdAt))
      : [];
    const groupIds = new Set(groupedNotifications.map(({ id }) => id));
    const latestGroupedNotification = groupedNotifications.reduce<DeviationNotification | null>((latest, notification) => (
      !latest || notification.createdAt > latest.createdAt ? notification : latest
    ), null);
    const items: NotificationListItem[] = [];
    let groupAdded = false;
    [...deviationNotifications]
      .sort((left, right) => right.createdAt.localeCompare(left.createdAt))
      .forEach((notification) => {
        if (groupIds.has(notification.id)) {
          if (!groupAdded && latestGroupedNotification) {
            items.push({
              kind: 'deviation-group',
              id: `group-${groupedNotifications.map(({ id }) => id).sort().join('-')}`,
              notifications: groupedNotifications,
              plates,
              createdAt: latestGroupedNotification.createdAt,
            });
            groupAdded = true;
          }
          return;
        }
        items.push({ kind: 'single', notification });
      });
    return items;
  }, [deviationNotifications]);

  const enableNotificationAlerts = () => {
    const AudioContextConstructor = window.AudioContext;
    try {
      if (AudioContextConstructor && !notificationAudioContextRef.current) {
        notificationAudioContextRef.current = new AudioContextConstructor();
      }
    } catch (error) {
      console.warn('Could not initialize notification audio', error);
    }
    const audioContext = notificationAudioContextRef.current;
    if (audioContext?.state === 'suspended') {
      void audioContext.resume().catch((error: unknown) => {
        console.warn('Could not enable notification audio', error);
      });
    }
    if (!('Notification' in window)) return;
    if (Notification.permission === 'granted') {
      notificationSoundEnabledRef.current = true;
      return;
    }
    if (Notification.permission === 'default') {
      void Notification.requestPermission().then((permission) => {
        notificationSoundEnabledRef.current = permission === 'granted';
      }).catch((error: unknown) => {
        console.warn('Could not request notification permission', error);
      });
    }
  };

  const toggleNotificationDetails = (notificationId: string) => {
    setExpandedNotificationIds((current) => {
      const next = new Set(current);
      if (next.has(notificationId)) next.delete(notificationId);
      else next.add(notificationId);
      return next;
    });
  };

  const dismissFleetNotification = (notification: FleetNotification) => {
    const dismissedIds = new Set(dismissedDeviationNotificationsRef.current);
    dismissedIds.add(notification.id);
    dismissedDeviationNotificationsRef.current = dismissedIds;
    try {
      localStorage.setItem(dismissedDeviationNotificationsKey, JSON.stringify([...dismissedIds]));
    } catch (error) {
      console.error('Failed to save dismissed deviation notification', error);
      setNoticeIsError(true);
      setNotice('تعذر حفظ حالة الإشعار؛ قد يظهر مجددًا عند إعادة فتح التطبيق.');
    }
  };

  const markAllFleetNotificationsRead = async () => {
    if (isMarkingAllNotificationsRead || !deviationNotifications.length) return;
    setIsMarkingAllNotificationsRead(true);
    try {
      const unreadNotifications = deviationNotifications.filter(({ id }) => !dismissedDeviationNotificationsRef.current.has(id));
      const dataUpdate = unreadNotifications.find(isFleetUpdateNotification);
      if (dataUpdate) {
        const cloudData = await readCloudFleetDataStrict<FleetData>();
        if (!cloudData) throw new Error('لم يتم العثور على نسخة حديثة من البيانات في السحابة.');
        const refreshedData = normalizeFleetData(cloudData);
        lastSyncedFleetFingerprintRef.current = fleetDataFingerprint(cloudData);
        setData(refreshedData);
      }

      const dismissedIds = new Set(dismissedDeviationNotificationsRef.current);
      unreadNotifications.forEach(({ id }) => dismissedIds.add(id));
      localStorage.setItem(dismissedDeviationNotificationsKey, JSON.stringify([...dismissedIds]));
      dismissedDeviationNotificationsRef.current = dismissedIds;
      setNotice(dataUpdate ? 'تم تحديث البيانات وقراءة كل الإشعارات.' : 'تمت قراءة كل الإشعارات.');
      setNoticeIsError(false);
    } catch (error) {
      console.error('Failed to mark all notifications as read', error);
      setNotice(error instanceof Error ? `تعذر قراءة كل الإشعارات: ${error.message}` : 'تعذر قراءة كل الإشعارات. أعد المحاولة.');
      setNoticeIsError(true);
    } finally {
      setIsMarkingAllNotificationsRead(false);
    }
  };

  const refreshFleetFromNotification = async (notification: FleetUpdateNotification) => {
    setRefreshingNotificationId(notification.id);
    try {
      const cloudData = await readCloudFleetDataStrict<FleetData>();
      if (!cloudData) throw new Error('لم يتم العثور على نسخة حديثة من البيانات في السحابة.');
      const refreshedData = normalizeFleetData(cloudData);
      lastSyncedFleetFingerprintRef.current = fleetDataFingerprint(cloudData);
      setData(refreshedData);
      dismissFleetNotification(notification);
      setNotice('تم تحديث البيانات بنجاح.');
      setNoticeIsError(false);
    } catch (error) {
      console.error('Failed to refresh fleet data from notification', error);
      setNotice(error instanceof Error ? `تعذر تحديث البيانات: ${error.message}` : 'تعذر تحديث البيانات. أعد المحاولة.');
      setNoticeIsError(true);
    } finally {
      setRefreshingNotificationId(null);
    }
  };

  const openDeviationNotification = (notification: DeviationNotification) => {
    dismissFleetNotification(notification);
    setSelectedDeviationMonths([notification.month]);
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

  const startEntry = () => {
    if (!canEditFleet) {
      notifyAccessDenied();
      return;
    }
    const savedDraft = readEntryDraft(entryDraftStorageKey);
    if (savedDraft) {
      entryDraftSuppressedRef.current = false;
      setEditingEntryId(null);
      setIsFuelSectionExpanded(false);
      setIsExpenseSectionExpanded(false);
      setExpandedExpenseKey(null);
      setEntryForm({ ...savedDraft.form, date: formatEntryDateInput(savedDraft.form.date) });
      setEntryPlateSearch(savedDraft.entryPlateSearch);
      setFuelAmountDraft(savedDraft.fuelAmountDraft);
      setExpenseAmountDrafts(savedDraft.expenseAmountDrafts);
      setEntryFormError('');
      setEntryDialog(true);
      setNotice('تم استعادة مسودة الحركة السابقة');
      setNoticeIsError(false);
      return;
    }
    entryDraftSuppressedRef.current = false;
    setEditingEntryId(null);
    setIsFuelSectionExpanded(false);
    setFuelAmountDraft('');
    setEntryPlateSearch('');
    setExpenseAmountDrafts(emptyExpenseInputs());
    setExpandedExpenseKey(null);
    setIsExpenseSectionExpanded(false);
    setEntryFormError('');
    setEntryForm({ ...createEmptyEntryForm(), vehicleId: '', driver: '', startOdometer: '', liters: '0.0', pricePerLiter: '0' });
    setEntryDialog(true);
  };

  const clearEntryForm = () => {
    if (editingEntryId) return;
    const vehicle = data.vehicles.find((item) => item.id === entryForm.vehicleId);
    const summary = summaries.find((item) => item.vehicle.id === entryForm.vehicleId);
    entryDraftSuppressedRef.current = false;
    setEntryForm({
      ...createEmptyEntryForm(),
      vehicleId: vehicle?.id ?? '',
      driver: vehicle?.driver ?? '',
      startOdometer: String(summary?.lastOdometer ?? vehicle?.openingOdometer ?? 0),
      liters: '0.0',
      pricePerLiter: '0',
      date: '',
    });
    setEntryPlateSearch('');
    setFuelAmountDraft('');
    setExpenseAmountDrafts(emptyExpenseInputs());
    setExpandedExpenseKey(null);
    setIsExpenseSectionExpanded(false);
    setNotice('تم مسح البيانات المدخلة');
    setNoticeIsError(false);
  };

  const selectEntryVehicle = (vehicleId: string) => {
    const vehicleSummary = summaries.find((item) => item.vehicle.id === vehicleId);
    if (!vehicleSummary) return;
    setEntryForm((current) => current.vehicleId === vehicleId ? current : ({
      ...current,
      vehicleId,
      driver: vehicleSummary.vehicle.driver || '',
      startOdometer: String(vehicleSummary.lastOdometer ?? 0),
      currentOdometer: '',
      liters: '0.0',
    }));
  };

  const updateEntryVehicleSearch = (query: string) => {
    const matches = filterVehicles(query, data.vehicles);
    const uniqueMatch = findUniqueVehicleMatch(query, data.vehicles);
    setEntryPlateSearch(query);
    if (!normalizeVehicleSearchKey(query)) {
      setEntryForm((current) => ({ ...current, vehicleId: '', driver: '', startOdometer: '' }));
      return;
    }
    if (uniqueMatch) {
      selectEntryVehicle(uniqueMatch.id);
      return;
    }
    setEntryForm((current) => matches.some(({ id }) => id === current.vehicleId)
      ? current
      : { ...current, vehicleId: '' });
  };

  const updateMaintenanceVehicleSearch = (query: string) => {
    const matches = filterVehicles(query, data.vehicles);
    const uniqueMatch = findUniqueVehicleMatch(query, data.vehicles);
    setMaintenanceVehicleSearch(query);
    setMaintenanceForm((current) => {
      if (!normalizeVehicleSearchKey(query)) {
        return current.vehicleId ? { ...current, vehicleId: '' } : current;
      }
      const vehicleId = uniqueMatch?.id
        ?? (matches.some(({ id }) => id === current.vehicleId) ? current.vehicleId : '');
      return current.vehicleId === vehicleId ? current : { ...current, vehicleId };
    });
  };

  const openMaintenanceDialog = () => {
    if (!canEditFleet) {
      notifyAccessDenied();
      return;
    }
    setMaintenanceVehicleSearch('');
    setMaintenanceForm({
      vehicleId: '',
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
    recordAudit('edit', `صيانة سيارة ${vehicle.plate} - تعبئة ${formatDateArabic(latestEntry.date)}`);
    setMaintenanceDialog(false);
    setNotice(`تمت إضافة الصيانة إلى آخر سجل لسيارة ${vehicle.name}`);
    setNoticeIsError(false);
  };

  const closeEntryPage = () => {
    setEntryFormError('');
    setEditingEntryId(null);
    const currentRoute = (window.history.state as FleetHistoryState | null)?.fleetRoute;
    if (currentRoute?.entryPage) {
      window.history.back();
      return;
    }
    setEntryDialog(false);
  };

  const editEntry = (entry: FuelEntry) => {
    if (!canEditFleet) {
      notifyAccessDenied();
      return;
    }
    entryDraftSuppressedRef.current = true;
    setEntryFormError('');
    const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
    setEditingEntryId(entry.id);
    setIsFuelSectionExpanded(false);
    setFuelAmountDraft('');
    setExpenseAmountDrafts(emptyExpenseInputs());
    setExpandedExpenseKey(null);
    setIsExpenseSectionExpanded(false);
    setEntryForm({
      documentNumber: entry.documentNumber || '',
      vehicleId: entry.vehicleId,
      driver: entry.driver || vehicle?.driver || '',
      startOdometer: String(entry.startOdometer),
      currentOdometer: String(entry.currentOdometer),
      liters: calculateFuelLiters(entry.vehicleId, entry.startOdometer, entry.currentOdometer).toFixed(1),
      pricePerLiter: String(entry.pricePerLiter),
      fuelActualValue: entry.fuelActualExpression ?? String(entryFuelCost(entry)),
      allowanceValue: String(allowanceValueForEntry(entry)),
      date: formatEntryDateInput(entry.date),
      route: entry.route || '',
      note: entry.note || '',
      expenses: Object.fromEntries(expenseFields.map(({ key }) => [key, entry.expenseExpressions?.[key] ?? (entry.expenses?.[key] ? String(entry.expenses[key]) : '')])) as ExpenseInputs,
    });
    setEntryDialog(true);
  };

  const flashAddedAmount = (target: string) => {
    if (amountFlashTimeoutRef.current !== null) window.clearTimeout(amountFlashTimeoutRef.current);
    setRecentlyAddedAmount({ target, token: ++amountFlashTokenRef.current });
    amountFlashTimeoutRef.current = window.setTimeout(() => {
      setRecentlyAddedAmount(null);
      amountFlashTimeoutRef.current = null;
    }, 1200);
  };

  const addFuelAmount = () => {
    const amount = parseSingleAmount(fuelAmountDraft.trim());
    if (amount === null || !fuelAmountDraft.trim()) {
      setNotice('أدخل مبلغًا واحدًا صحيحًا ثم اضغط إضافة');
      return;
    }
    setEntryForm((current) => ({ ...current, fuelActualValue: [...splitAmountExpression(current.fuelActualValue), String(amount)].join('+') }));
    flashAddedAmount(`fuel:${splitAmountExpression(entryForm.fuelActualValue).length}`);
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
    flashAddedAmount(`expense:${key}:${splitAmountExpression(entryForm.expenses[key]).length}`);
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
      vehicleType: vehicleForm.vehicleType,
      plate: vehicleForm.plate.trim(),
      driver: vehicleForm.driver.trim(),
      openingOdometer: Number(vehicleForm.openingOdometer),
      kmPerLiter: Number(vehicleForm.kmPerLiter) > 0 ? 105 / Number(vehicleForm.kmPerLiter) : 5.25,
      fuelLitersPer105Km: Number(vehicleForm.kmPerLiter),
      fuelDistanceBaseKm: Number(vehicleForm.distanceBaseKm) > 0 ? Number(vehicleForm.distanceBaseKm) : 105,
    };
    setData((current) => ({ ...current, vehicles: editingVehicleId ? current.vehicles.map((item) => item.id === editingVehicleId ? vehicle : item) : [...current.vehicles, vehicle] }));
    if (editingVehicleId) recordAudit('edit', `السيارة ${vehicle.plate} (${vehicle.name})`);
    setVehicleDialog(false);
    setEditingVehicleId(null);
    setVehicleForm({ name: '', vehicleType: 'dmax', plate: '', driver: '', openingOdometer: '', kmPerLiter: '20', distanceBaseKm: '105' });
    setNotice(editingVehicleId ? 'تم تحديث بيانات السيارة' : 'تمت إضافة السيارة');
  };

  const openVehicleEditor = (vehicle: Vehicle) => {
    if (!canEditFleet) {
      notifyAccessDenied();
      return;
    }
    setEditingVehicleId(vehicle.id);
    setVehicleForm({ name: vehicle.name, vehicleType: vehicle.vehicleType || inferVehicleType(vehicle.name) || 'dmax', plate: vehicle.plate, driver: vehicle.driver, openingOdometer: String(vehicle.openingOdometer), kmPerLiter: String(vehicleFuelRate(vehicle)), distanceBaseKm: String(vehicleDistanceBase(vehicle)) });
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
          else vehicles.push({ id: `v${Date.now()}-${index}`, ...vehicle, vehicleType: inferVehicleType(vehicle.name) });
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
    const vehicleHeaders = ['معرف السيارة', 'اللوحة', 'اسم السيارة', 'السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)', 'القراءة السابقة الافتراضية'];
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
    for (let rowNumber = 2; rowNumber <= vehicleSheet.rowCount; rowNumber += 1) {
      [4, 5, 6].forEach((columnNumber) => {
        const cell = vehicleSheet.getRow(rowNumber).getCell(columnNumber);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        cell.border = { top: { style: 'thin', color: { argb: 'FFB38F00' } }, bottom: { style: 'thin', color: { argb: 'FFB38F00' } }, left: { style: 'thin', color: { argb: 'FFB38F00' } }, right: { style: 'thin', color: { argb: 'FFB38F00' } } };
      });
    }
    [1, 2, 3].forEach((columnNumber) => {
      const column = vehicleSheet.getColumn(columnNumber);
      column.hidden = true;
      column.width = 0.01;
    });
    vehicleSheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: data.vehicles.length + 1, column: vehicleHeaders.length } };
    workbook.definedNames.add(`'بيانات السيارات'!$B$2:$B$${data.vehicles.length + 1}`, 'VehiclePlates');

    const entryHeaders = [
      'معرف السيارة', 'اللوحة', 'معرف الحركة', 'رقم المستند', 'التاريخ', 'السائق', 'المسار',
      'قراءة سابقة', 'قراءة حالية', 'التموين الفعلي',
      ...expenseFields.map((field) => compactExpenseHeaders[field.key]), 'ملاحظات',
    ];
    const entryInputColumns = new Set([4, 5, 6, 7, 9]);
    for (let columnNumber = 10; columnNumber <= entryHeaders.length; columnNumber += 1) {
      entryInputColumns.add(columnNumber);
    }
    const entryColumnWidths = [24, 20, 25, 18, 16, 22, 24, 18, 18, 24, ...expenseFields.map(() => 18), 28];
    const usedVehicleSheetNames = new Set(['بيانات السيارات'.toLocaleLowerCase('ar-EG')]);
    data.vehicles.forEach((vehicle) => {
      const entrySheet = workbook.addWorksheet(createVehicleWorksheetName(vehicle.plate, usedVehicleSheetNames), {
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
          vehicle.id, vehicle.plate, entry.id, entry.documentNumber || '', new Date(`${entry.date}T00:00:00`),
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
            ? `IFERROR(LOOKUP(2,1/(($B$5:B${previousRowsEnd}=B${rowNumber})*($I$5:I${previousRowsEnd}<>"")),$I$5:I${previousRowsEnd}),VLOOKUP(B${rowNumber},'بيانات السيارات'!$B$2:$G$${data.vehicles.length + 1},6,FALSE))`
            : `VLOOKUP(B${rowNumber},'بيانات السيارات'!$B$2:$G$${data.vehicles.length + 1},6,FALSE)`;
          row.getCell(8).value = { formula: `IF(B${rowNumber}="","",IFERROR(${previousReadingFormula},""))` };
        }
        entryInputColumns.forEach((columnNumber) => {
          const cell = row.getCell(columnNumber);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
          cell.border = { top: { style: 'thin', color: { argb: 'FFB38F00' } }, bottom: { style: 'thin', color: { argb: 'FFB38F00' } }, left: { style: 'thin', color: { argb: 'FFB38F00' } }, right: { style: 'thin', color: { argb: 'FFB38F00' } } };
        });
      }
      [1, 2, 3].forEach((columnNumber) => {
        const column = entrySheet.getColumn(columnNumber);
        column.hidden = true;
        column.width = 0.01;
      });
      entrySheet.getColumn(5).numFmt = 'yyyy-mm-dd';
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
      const vehicleHeaders = normalizeTemplateHeaders(vehicleRows[0] || []);
      const requiredVehicleHeaders = ['رقم اللوحة', 'اسم السيارة', 'اسم السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)'];
      const missingVehicleHeaders = requiredVehicleHeaders.filter((header) => !vehicleHeaders.includes(header));
      if (missingVehicleHeaders.length) throw new Error(`أعمدة بيانات السيارات الناقصة: ${missingVehicleHeaders.join('، ')}`);
      const requiredEntryHeaders = [
        'رقم اللوحة', 'معرف الحركة', 'التاريخ', 'اسم السائق', 'خط السير',
        'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)', 'التفصيل',
        ...expenseFields.map((field) => `${field.label} (جنيه)`),
      ];

      const normalizePlate = (plate: string) => plate.trim().toLocaleLowerCase('ar-EG');
      const vehiclesByPlate = new Map(data.vehicles.map((vehicle) => [normalizePlate(vehicle.plate), vehicle]));
      const importedVehicles = new Map<string, Vehicle>();
      const seenVehicleIds = new Set<string>();
      for (const [index, row] of vehicleRows.slice(1).entries()) {
        const value = (header: string) => (row[vehicleHeaders.indexOf(header)] || '').trim();
        const plate = value('رقم اللوحة');
        if (!plate && !row.some((cell) => cell.trim())) continue;
        const lineNumber = index + 2;
        const existing = vehiclesByPlate.get(normalizePlate(plate));
        if (!existing) {
          throw new Error(`ورقة السيارات، صف ${lineNumber}: رقم اللوحة غير موجود`);
        }
        if (seenVehicleIds.has(existing.id)) {
          throw new Error(`ورقة السيارات، صف ${lineNumber}: رقم اللوحة مكرر`);
        }
        seenVehicleIds.add(existing.id);
        const name = value('اسم السيارة');
        const driver = value('اسم السائق');
        const openingText = value('بداية العداد');
        const rateText = value('معدل الاستهلاك (كم/لتر)');
        const openingOdometer = parseTemplateNumber(openingText);
        const kmPerLiter = parseTemplateNumber(rateText);
        const errors: string[] = [];
        if (!name) errors.push('اسم السيارة مطلوب');
        if (!openingText || !Number.isFinite(openingOdometer) || openingOdometer < 0) errors.push('بداية العداد غير صحيحة');
        if (!rateText || !Number.isFinite(kmPerLiter) || kmPerLiter <= 0) errors.push('معدل الاستهلاك يجب أن يكون أكبر من صفر');
        if (errors.length) {
          throw new Error(`ورقة السيارات، صف ${lineNumber}: ${errors.join('، ')}`);
        }
        importedVehicles.set(existing.id, {
          ...existing, plate, name, driver, openingOdometer, kmPerLiter,
        });
      }
      const missingVehicle = data.vehicles.find((vehicle) => !importedVehicles.has(vehicle.id));
      if (missingVehicle) throw new Error(`القالب لا يحتوي على بيانات السيارة ذات اللوحة ${missingVehicle.plate}`);

      const existingEntriesById = new Map(data.entries.map((entry) => [entry.id, entry]));
      const seenEntryIds = new Set<string>();
      const importedEntries: FuelEntry[] = [];
      let skippedRows = 0;
      const vehicleSheets = workbook.worksheets.filter((sheet) => sheet !== vehicleSheet);
      const seenSheetVehicleIds = new Set<string>();
      for (const sheet of vehicleSheets) {
        const rows = worksheetToRows(sheet);
        const metadata = rows[1] || [];
        const plate = metadata[3]?.trim() || '';
        const existingVehicle = vehiclesByPlate.get(normalizePlate(plate));
        const vehicleId = existingVehicle?.id || '';
        const vehicle = vehicleId ? importedVehicles.get(vehicleId) : undefined;
        if (normalizeTemplateHeader(metadata[2] || '') !== 'رقم اللوحة' || !plate || !vehicle || !existingVehicle) {
          throw new Error(`ورقة «${sheet.name}»: رقم اللوحة غير صحيح أو لا يطابق ورقة بيانات السيارات`);
        }
        if (seenSheetVehicleIds.has(vehicleId)) {
          throw new Error(`ورقة «${sheet.name}»: يوجد أكثر من ورقة للسيارة نفسها`);
        }
        seenSheetVehicleIds.add(vehicleId);

        const entryHeaderIndex = rows.findIndex((row) => {
          const headers = normalizeTemplateHeaders(row);
          return headers.includes('معرف الحركة') && headers.includes('التاريخ');
        });
        if (entryHeaderIndex < 0) {
          throw new Error(`ورقة «${sheet.name}»: لم أجد عناوين سجل الحركات`);
        }
        const entryHeaders = normalizeTemplateHeaders(rows[entryHeaderIndex]);
        const missingEntryHeaders = requiredEntryHeaders.filter((header) => !entryHeaders.includes(header));
        if (missingEntryHeaders.length) {
          throw new Error(`ورقة «${sheet.name}»: الأعمدة الناقصة ${missingEntryHeaders.join('، ')}`);
        }
        const column = (header: string) => entryHeaders.indexOf(header);
        for (const [index, row] of rows.slice(entryHeaderIndex + 1).entries()) {
          const value = (header: string) => (row[column(header)] || '').trim();
          const dateText = value('التاريخ');
          const startText = value('القراءة السابقة');
          const currentText = value('القراءة الحالية');
          const hasMovementData = [
            'معرف الحركة', 'رقم المستند', 'التاريخ', 'اسم السائق', 'خط السير', 'القراءة السابقة',
            'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)', 'التفصيل',
            ...expenseFields.map((field) => `${field.label} (جنيه)`),
          ].some((header) => value(header) !== '');
          if (!hasMovementData) continue;
          if (!dateText || !startText || !currentText) {
            skippedRows += 1;
            continue;
          }
          const lineNumber = index + entryHeaderIndex + 2;
          const rowVehicleId = vehicleId;
          const rowVehicle = vehicle;
          const rowPlate = value('رقم اللوحة');
          const errors: string[] = [];
          if (rowPlate && normalizePlate(rowPlate) !== normalizePlate(rowVehicle.plate)) errors.push('رقم اللوحة لا تطابق هذه الورقة');
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
          if (!value('رقم المستند') && !oldEntry?.documentNumber) errors.push('رقم المستند مطلوب');
          if (suppliedId && seenEntryIds.has(suppliedId)) errors.push('معرف الحركة مكرر داخل الملف');
          if (errors.length) {
            throw new Error(`ورقة «${sheet.name}»، صف ${lineNumber}: ${errors.join('، ')}`);
          }
          if (suppliedId) seenEntryIds.add(suppliedId);
          const id = oldEntry ? suppliedId : `e${Date.now()}-${lineNumber}-${rowVehicleId}`;
          const liters = litersForDistance(rowVehicle, current - start);
          const allowancePricePerLiter = oldEntry?.allowancePricePerLiter ?? fuelAllowanceRate(data);
          importedEntries.push({
            id, vehicleId: rowVehicleId, date: date || '', startOdometer: start, currentOdometer: current, liters,
            allowancePricePerLiter,
            pricePerLiter: liters > 0 ? actualFuelValue / liters : 0, fuelActualValue: actualFuelValue,
            driver: value('اسم السائق'), route: value('خط السير'),
            allowanceValue: allowanceCost(liters, allowancePricePerLiter), expenses, note: value('التفصيل'),
            documentNumber: value('رقم المستند') || oldEntry?.documentNumber || '',
          });
        }
      }
      const missingVehicleSheet = data.vehicles.find((vehicle) => !seenSheetVehicleIds.has(vehicle.id));
      if (missingVehicleSheet) throw new Error(`القالب لا يحتوي على ورقة مستقلة للسيارة ذات اللوحة ${missingVehicleSheet.plate}`);

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
    console.info('Exporting Single Vehicle Template');
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
      ['اللوحة', selectedSummary.vehicle.plate, 'اسم السيارة', selectedSummary.vehicle.name, 'السائق', selectedSummary.vehicle.driver],
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
      'معرف الحركة', 'رقم اللوحة', 'اسم السيارة', 'رقم المستند', 'التاريخ', 'اسم السائق', 'خط السير',
      'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)',
      ...expenseFields.map((field) => `${field.label} (جنيه)`), 'التفصيل',
    ];
    const hiddenColumns = new Set([1, 2, 3, 10, 11, 12, 14, 15, headers.length]);
    const inputColumns = new Set([4, 5, 6, 7, 8, 9, 13, 16 + expenseFields.length]);
    expenseFields.forEach((_, index) => inputColumns.add(16 + index));
    sheet.columns = [12, 17, 22, 18, 15, 20, 25, 17, 17, 15, 20, 24, 24, 20, 19, ...expenseFields.map(() => 17), 26, 26]
      .map((width, index) => ({ width, hidden: hiddenColumns.has(index + 1) }));
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
        entry?.documentNumber || '', entry ? new Date(`${entry.date}T00:00:00`) : null,
        entry?.driver || selectedSummary.vehicle.driver, entry?.route || '',
        entry?.startOdometer ?? (index === 0 ? selectedSummary.lastOdometer : null), entry?.currentOdometer ?? null,
        { formula: `IF(OR(H${rowNumber}="",I${rowNumber}=""),"",MAX(0,I${rowNumber}-H${rowNumber}))` },
        { formula: `IF(OR(J${rowNumber}="",'بيانات السيارة'!$D$4<=0),"",J${rowNumber}/'بيانات السيارة'!$D$4)` },
        { formula: `IF(K${rowNumber}="","",ROUND(K${rowNumber}*'بيانات السيارة'!$F$4,0))` },
        entry ? entryFuelCost(entry) : null,
        { formula: `IF(M${rowNumber}="","",M${rowNumber}-L${rowNumber})` },
        { formula: `IF(N${rowNumber}="","",SUM($N$5:N${rowNumber}))` },
        ...expenseFields.map((field) => entry ? Math.round(expenses[field.key] || 0) : null),
        entry?.note || '',
        { formula: `IF(COUNTA(M${rowNumber},P${rowNumber}:AA${rowNumber})=0,"",ROUND(SUM(M${rowNumber},P${rowNumber}:AA${rowNumber}),0))` },
      ];
      const row = sheet.addRow(values);
      row.height = 25;
      row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
        cell.font = { name: 'Arial', size: 10, color: { argb: 'FF263A34' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { bottom: { style: 'hair', color: { argb: 'FFD7E0DA' } }, left: { style: 'hair', color: { argb: 'FFE0E6E1' } }, right: { style: 'hair', color: { argb: 'FFE0E6E1' } } };
        const calculated = [10, 11, 12, 14, 15, headers.length].includes(columnNumber);
        if (inputColumns.has(columnNumber)) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
          cell.border = { top: { style: 'thin', color: { argb: 'FFB38F00' } }, bottom: { style: 'thin', color: { argb: 'FFB38F00' } }, left: { style: 'thin', color: { argb: 'FFB38F00' } }, right: { style: 'thin', color: { argb: 'FFB38F00' } } };
        } else if (calculated) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7EEE9' } };
        }
        if (columnNumber === 5) cell.numFmt = 'yyyy-mm-dd';
        if ([12, 13, 14, 15, headers.length].includes(columnNumber)) cell.numFmt = '#,##0 "جنيه"';
        if ([8, 9, 10].includes(columnNumber)) cell.numFmt = '#,##0';
        if (columnNumber === 11) cell.numFmt = '0.0';
      });
    }
    for (let rowNumber = 5; rowNumber <= sheet.rowCount; rowNumber += 1) {
      inputColumns.forEach((columnNumber) => {
        const cell = sheet.getRow(rowNumber).getCell(columnNumber);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        cell.border = { top: { style: 'thin', color: { argb: 'FFB38F00' } }, bottom: { style: 'thin', color: { argb: 'FFB38F00' } }, left: { style: 'thin', color: { argb: 'FFB38F00' } }, right: { style: 'thin', color: { argb: 'FFB38F00' } } };
      });
    }
    hiddenColumns.forEach((columnNumber) => {
      const column = sheet.getColumn(columnNumber);
      column.hidden = true;
      column.width = 0.01;
    });
    [1, 2, 3].forEach((columnNumber) => {
      const column = sheet.getColumn(columnNumber);
      column.hidden = true;
      column.width = 0.01;
    });
    sheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4, column: headers.length } };
    sheet.getColumn(5).numFmt = 'yyyy-mm-dd';

    const buffer = await workbook.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buffer as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = `سجل السيارة ${selectedSummary.vehicle.plate.replace(/[\\/:*?"<>|\s]+/g, '-')}.xlsx`;
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
      const headerRowIndex = rows.findIndex((row) => {
        const normalized = normalizeTemplateHeaders(row);
        return normalized.includes('معرف الحركة')
          && normalized.includes('رقم اللوحة')
          && normalized.includes('التاريخ');
      });
      if (headerRowIndex < 0) throw new Error('لم أجد صف عناوين قالب الحركات.');
      const headers = normalizeTemplateHeaders(rows[headerRowIndex]);
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
        if (normalizeTemplateHeader(info[0] || '') !== 'رقم اللوحة' || info[2]?.trim() !== 'اسم السيارة' || normalizeTemplateHeader(info[4] || '') !== 'اسم السائق'
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
        };
      }

      rows.slice(headerRowIndex + 1).forEach((row, index) => {
        const value = (header: string) => (row[column(header)] || '').trim();
        const editableHeaders = [
          'معرف الحركة', 'رقم المستند', 'التاريخ', 'اسم السائق', 'خط السير',
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
        const previousEntry = suppliedId ? oldEntries.find((entry) => entry.id === suppliedId) : undefined;
        if (!value('رقم المستند') && !previousEntry?.documentNumber) {
          errors.push(`صف ${lineNumber}: رقم المستند مطلوب`);
          return;
        }
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
        const liters = calculateFuelLiters(vehicle.id, start, current);
        const allowancePricePerLiter = previousEntry?.allowancePricePerLiter ?? fuelAllowanceRate(data);
        imported.push({
          id, vehicleId: vehicle.id, date, startOdometer: start, currentOdometer: current,
          allowancePricePerLiter,
          liters, pricePerLiter: liters > 0 ? actualFuelValue / liters : 0, fuelActualValue: actualFuelValue,
          driver: value('اسم السائق'), route: value('خط السير'), allowanceValue: allowanceCost(liters, allowancePricePerLiter),
          expenses, note: value('التفصيل'), documentNumber: value('رقم المستند') || previousEntry?.documentNumber || '',
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

  const openDeviationNoteEditor = (vehicleId: string) => {
    if (!canEditFleet || selectedDeviationMonths.length !== 1) return;
    const month = selectedDeviationMonths[0];
    setDeviationNoteDraft(deviationNotes[month]?.[vehicleId]?.text ?? '');
    setEditingDeviationNote({ vehicleId, month });
  };

  const toggleDeviationVehicleVisibility = (vehicleId: string) => {
    if (!canEditFleet) return;
    const isHidden = (latestFleetDataRef.current.hiddenDeviationVehicleIds || []).includes(vehicleId);
    setData((current) => {
      const hiddenVehicleIds = new Set(current.hiddenDeviationVehicleIds || []);
      if (hiddenVehicleIds.has(vehicleId)) hiddenVehicleIds.delete(vehicleId);
      else hiddenVehicleIds.add(vehicleId);
      return { ...current, hiddenDeviationVehicleIds: [...hiddenVehicleIds] };
    });
    setNotice(isHidden ? 'تم إظهار انحراف السيارة للمستخدمين' : 'تم إخفاء انحراف السيارة عن المستخدمين');
    setNoticeIsError(false);
  };

  const saveDeviationNote = () => {
    if (!editingDeviationNote || !canEditFleet) return;
    const { vehicleId, month } = editingDeviationNote;
    const text = deviationNoteDraft.trim();
    const nextMonthNotes = { ...deviationNotes[month] };
    if (text) {
      nextMonthNotes[vehicleId] = {
        text,
        createdAt: nextMonthNotes[vehicleId]?.createdAt ?? new Date().toISOString(),
      };
    } else {
      delete nextMonthNotes[vehicleId];
    }
    const nextNotes = { ...deviationNotes };
    if (Object.keys(nextMonthNotes).length) nextNotes[month] = nextMonthNotes;
    else delete nextNotes[month];
    try {
      localStorage.setItem(DEVIATION_NOTES_STORAGE_KEY, JSON.stringify(nextNotes));
      setDeviationNotes(nextNotes);
      setEditingDeviationNote(null);
      setDeviationNoteDraft('');
      setNotice('تم حفظ ملاحظة الانحراف');
      setNoticeIsError(false);
    } catch (error) {
      console.error('Failed to save monthly deviation note', error);
      setNotice('تعذر حفظ ملاحظة الانحراف على هذا الجهاز.');
      setNoticeIsError(true);
    }
  };

  const saveEntry = (event: React.FormEvent) => {
    event.preventDefault();
    if (!canEditFleet) return;
    setEntryFormError('');
    if (!entryForm.vehicleId || !data.vehicles.some((vehicle) => vehicle.id === entryForm.vehicleId)) {
      setEntryFormError('اختر سيارة من القائمة قبل حفظ الحركة');
      return;
    }
    const documentNumber = entryForm.documentNumber.trim();
    if (!documentNumber) {
      setEntryFormError('رقم المستند مطلوب قبل حفظ حركة التعبئة');
      return;
    }
    if (!entryForm.date.trim()) {
      setEntryFormError('التاريخ مطلوب قبل حفظ الحركة');
      return;
    }
    const entryDate = parseEntryDateInput(entryForm.date);
    if (!entryDate) {
      setEntryFormError('أدخل تاريخًا صحيحًا بصيغة يوم/شهر/سنة، مثل 18/11/2026');
      return;
    }
    if (fuelAmountDraft.trim() || (Object.values(expenseAmountDrafts) as string[]).some((amount) => amount.trim())) {
      setEntryFormError('اضغط «إضافة» بجوار كل مبلغ لإدخاله في الإجمالي قبل الحفظ');
      return;
    }
    if (!entryForm.fuelActualValue.trim()) {
      setEntryFormError('أضف مبلغ مصروف الوقود قبل الحفظ');
      return;
    }
    const start = Number(entryForm.startOdometer);
    const current = Number(entryForm.currentOdometer);
    if (!Number.isFinite(start) || start < 0 || !Number.isFinite(current) || current < 0) {
      setEntryFormError('أدخل قراءتي عداد صحيحتين غير سالبتين');
      return;
    }
    if (current < start) {
      setEntryFormError('العداد الحالي يجب أن يكون أكبر من أو يساوي بداية العداد');
      return;
    }
    const fuelActualValue = parseAmountExpression(entryForm.fuelActualValue);
    const parsedExpenses = Object.fromEntries(expenseFields.map(({ key }) => [key, parseAmountExpression(entryForm.expenses[key])])) as Record<keyof ExpenseBreakdown, number | null>;
    if (fuelActualValue === null || Object.values(parsedExpenses).some((amount) => amount === null)) {
      setEntryFormError('راجع المبالغ المضافة وتأكد من صحتها');
      return;
    }
    const calculatedLiters = calculateFuelLiters(entryForm.vehicleId, start, current);
    const previousEntry = editingEntryId ? data.entries.find((item) => item.id === editingEntryId) : undefined;
    const allowancePricePerLiter = previousEntry?.allowancePricePerLiter ?? fuelAllowanceRate(data);
    const entry: FuelEntry = {
      id: editingEntryId || `e${Date.now()}`,
      createdAt: previousEntry?.createdAt ?? new Date().toISOString(),
      documentNumber,
      vehicleId: entryForm.vehicleId,
      date: entryDate,
      startOdometer: start,
      currentOdometer: current,
      liters: calculatedLiters,
      allowancePricePerLiter,
      pricePerLiter: calculatedLiters > 0 ? fuelActualValue / calculatedLiters : 0,
      fuelActualValue,
      fuelActualExpression: entryForm.fuelActualValue.trim(),
      driver: entryForm.driver.trim(),
      allowanceValue: allowanceCost(calculatedLiters, allowancePricePerLiter),
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
    const crossedWarningThreshold = previousMonthlyDeviation <= 100 && nextMonthlyDeviation > 100;
    const crossedDangerThreshold = previousMonthlyDeviation <= 250 && nextMonthlyDeviation > 250;
    const newDeviationNotification = !editingEntryId && (crossedWarningThreshold || crossedDangerThreshold) && vehicle
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
        setEntryFormError('تعذر مسح المسودة المحفوظة. لم يتم حفظ الحركة');
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
    if (!editingEntryId) {
      entryDraftSuppressedRef.current = true;
      setEntryForm({ ...createEmptyEntryForm(), vehicleId: '', driver: '', startOdometer: '', liters: '0.0', pricePerLiter: '0' });
      setEntryPlateSearch('');
      setFuelAmountDraft('');
      setExpenseAmountDrafts(emptyExpenseInputs());
      setIsFuelSectionExpanded(false);
      setIsExpenseSectionExpanded(false);
      setExpandedExpenseKey(null);
      setRecentlyAddedAmount(null);
    }
    if (editingEntryId) closeEntryPage();
    setEntryFormError('');
    setNotice(editingEntryId ? 'تم تحديث الحركة' : 'تم حفظ الحركة بنجاح');
    setNoticeIsError(false);
    if (editingEntryId) recordAudit('edit', `تعبئة ${entry.documentNumber} للسيارة ${vehicle?.plate ?? ''} بتاريخ ${formatDateArabic(entry.date)}`);
    if (newDeviationNotification && isFirebaseConfigured) {
      void Promise.resolve(saveCloudFleetDataStrict(nextData))
        .then(() => publishDeviationNotification(newDeviationNotification))
        .catch((error: unknown) => {
          console.error('Failed to sync fuel deviation notification', error);
          setNoticeIsError(true);
          setNotice(`تم حفظ التعبئة، لكن تعذر إرسال إشعار الانحراف: ${error instanceof Error ? error.message : 'خطأ غير معروف'}`);
        });
    }
  };

  const confirmDeleteEntry = () => {
    if (!canEditFleet) return;
    if (!pendingDeleteEntry || deleteConfirmationText.trim() !== 'موافق') return;
    setData((current) => ({ ...current, entries: current.entries.filter((entry) => entry.id !== pendingDeleteEntry.id) }));
    const deletedVehicle = data.vehicles.find((vehicle) => vehicle.id === pendingDeleteEntry.vehicleId);
    recordAudit('delete', `تعبئة ${pendingDeleteEntry.documentNumber} للسيارة ${deletedVehicle?.plate ?? ''} بتاريخ ${formatDateArabic(pendingDeleteEntry.date)}`);
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

  const clearAllEntries = async () => {
    if (!canEditFleet || !isClearEntriesPasswordValid || !cloudLoaded || isClearingAllEntries) return;
    if (clearEntriesMode === 'month' && !/^\d{4}-(0[1-9]|1[0-2])$/.test(clearEntriesMonth)) return;
    if (!isFirebaseConfigured || !firestore) {
      setNoticeIsError(true);
      setNotice('تعذر مسح الحركات من Firestore لأن الاتصال غير مُعد.');
      return;
    }
    const currentData = latestFleetDataRef.current;
    const entriesToDelete = clearEntriesMode === 'all'
      ? currentData.entries
      : currentData.entries.filter((entry) => entry.date.slice(0, 7) === clearEntriesMonth);
    if (entriesToDelete.length === 0) {
      setIsClearEntriesConfirmationOpen(false);
      setClearEntriesConfirmationPassword('');
      setIsClearEntriesMonthPickerOpen(false);
      setNoticeIsError(false);
      setNotice(clearEntriesMode === 'all' ? 'لا توجد حركات لمسحها.' : 'لا توجد حركات مسجلة في الشهر المحدد.');
      return;
    }

    setIsClearingAllEntries(true);
    const deletedIds = new Set(entriesToDelete.map((entry) => entry.id));
    const nextData: FleetData = {
      ...currentData,
      entries: currentData.entries.filter((entry) => !deletedIds.has(entry.id)),
    };
    let previousLocalData: string | null = null;
    let localDataUpdated = false;
    try {
      previousLocalData = localStorage.getItem(STORAGE_KEY);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(nextData));
      localDataUpdated = true;
      await saveCloudFleetDataStrict(nextData);
      lastSyncedFleetFingerprintRef.current = fleetDataFingerprint(nextData);
      setData(nextData);
      setPendingDeleteEntry(null);
      setDeleteConfirmationText('');
      setIsClearEntriesConfirmationOpen(false);
      setClearEntriesConfirmationPassword('');
      setIsClearEntriesMonthPickerOpen(false);
      recordAudit(
        'delete',
        clearEntriesMode === 'all'
          ? `مسح جميع سجلات التعبئة (${entriesToDelete.length} حركة)`
          : `مسح سجلات التعبئة لشهر ${deviationMonthDisplayName(clearEntriesMonth)} (${entriesToDelete.length} حركة)`,
      );
      setNotice(clearEntriesMode === 'all'
        ? 'تم مسح كل الحركات مع الاحتفاظ ببيانات السيارات وإعدادات السعر'
        : `تم مسح ${entriesToDelete.length} حركة لشهر ${deviationMonthDisplayName(clearEntriesMonth)} مع الاحتفاظ بباقي الحركات والسيارات وإعدادات السعر`);
      setNoticeIsError(false);
    } catch (error) {
      console.error('Could not clear all fuel entries', error);
      if (localDataUpdated) {
        try {
          if (previousLocalData === null) localStorage.removeItem(STORAGE_KEY);
          else localStorage.setItem(STORAGE_KEY, previousLocalData);
        } catch (rollbackError) {
          console.error('Could not restore local fleet data after failed entry deletion', rollbackError);
        }
      }
      setNoticeIsError(true);
      setNotice(`تعذر مسح الحركات من السحابة: ${error instanceof Error ? error.message : 'خطأ غير معروف'}`);
    } finally {
      setIsClearingAllEntries(false);
    }
  };

  const exportExcel = async () => {
    const rows = [
      ['رقم المستند', 'التاريخ', 'اسم السائق', 'خط السير', 'القراءة السابقة', 'القراءة الحالية', 'الفجوة كم', 'عدد اللترات', 'قيمة التموين وفقًا للتر (جنيه)', 'قيمة التموين الفعلية (جنيه)', 'الانحراف (جنيه)', 'كم/لتر', ...expenseFields.map((field) => `${field.label} (جنيه)`), 'إجمالي مصاريف السيارة (جنيه)', 'رقم اللوحة', 'ملاحظات'],
      ...filteredEntries.map((entry) => {
        const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
        const liters = litersForEntry(entry, vehicle);
        const allowanceValue = entryAllowanceValue(entry, fuelAllowanceRate(data), liters);
        const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters);
        const accumulated = data.entries
          .filter((item) => item.vehicleId === entry.vehicleId && item.date <= entry.date)
          .reduce((sum, item) => sum + entryDeviation(item, fuelAllowanceRate(data)), 0);
        return [entry.documentNumber || '', entry.date, entry.driver || vehicle?.driver || '', entry.route || '', entry.startOdometer, entry.currentOdometer, Math.max(0, entry.currentOdometer - entry.startOdometer), liters, allowanceValue, entryFuelCost(entry), deviation, accumulated, ...expenseFields.map((field) => Math.round(entry.expenses?.[field.key] || 0)), entryTotalCost(entry), vehicle?.plate || '', entry.note];
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
    recordAudit('edit', 'سعر التموين المعتمد');
    setAllowanceRateDraft(rate.toFixed(2));
    setIsRateEditorOpen(false);
    setNotice('تم تحديث سعر التموين لجميع السيارات');
  };

  const entryVehicle = data.vehicles.find((vehicle) => vehicle.id === entryForm.vehicleId);
  const editingEntryAllowancePrice = editingEntryId
    ? data.entries.find((entry) => entry.id === editingEntryId)?.allowancePricePerLiter ?? fuelAllowanceRate(data)
    : fuelAllowanceRate(data);
  const entryDistance = Math.max(0, Number(entryForm.currentOdometer) - Number(entryForm.startOdometer));
  const enteredOdometer = entryForm.currentOdometer.trim() ? Number(entryForm.currentOdometer) : null;
  const entryOdometerSummary = enteredOdometer !== null && Number.isFinite(enteredOdometer)
    ? `العداد: ${numberFormat.format(enteredOdometer)} كم${entryDistance > 0 ? ` · الفجوة: ${numberFormat.format(entryDistance)} كم` : ''}`
    : null;
  const fuelActualValuePreview = parseAmountExpression(entryForm.fuelActualValue);
  const expenseAmountPreviews = Object.fromEntries(expenseFields.map(({ key }) => [key, parseAmountExpression(entryForm.expenses[key])])) as Record<keyof ExpenseBreakdown, number | null>;
  const entryExpenseTotal = Math.round(Object.values(expenseAmountPreviews).reduce<number>((sum, amount) => sum + (amount ?? 0), 0));
  const hasInvalidAmount = fuelActualValuePreview === null || Object.values(expenseAmountPreviews).some((amount) => amount === null);
  const hasUnaddedAmounts = Boolean(fuelAmountDraft.trim() || (Object.values(expenseAmountDrafts) as string[]).some((amount) => amount.trim()));
  const defaultEntryDriver = data.vehicles.find((vehicle) => vehicle.id === entryForm.vehicleId)?.driver ?? '';
  const hasEnteredEntryData = Boolean(
    entryForm.documentNumber.trim()
    || entryForm.currentOdometer.trim()
    || entryForm.route.trim()
    || entryForm.note.trim()
    || entryForm.fuelActualValue.trim()
    || Object.values(entryForm.expenses).some((amount) => amount.trim())
    || (Object.values(expenseAmountDrafts) as string[]).some((amount) => amount.trim())
    || fuelAmountDraft.trim()
    || entryPlateSearch.trim()
    || entryForm.date.trim()
    || entryForm.driver !== defaultEntryDriver
  );
  const entryCostPreview = Math.round((fuelActualValuePreview ?? 0) + Object.values(expenseAmountPreviews).reduce<number>((sum, amount) => sum + (amount ?? 0), 0));
  const navigationTiles = fleetSections.filter(({ key }) => key !== 'audit' || canManageUsers);
  const homeTileOrder = ['overview', 'vehicles', 'entries', 'audit', 'routes'];
  const homeNavigationTiles = navigationTiles
    .filter(({ key }) => currentUser.role !== 'viewer' || key !== 'entries')
    .sort((first, second) => homeTileOrder.indexOf(first.key) - homeTileOrder.indexOf(second.key));
  const openSection = navigationTiles.find((section) => section.key === tab)
    || (tab === 'users' ? fleetHomeActions[0] : undefined);
  const navigateToSection = (key: typeof fleetSections[number]['key']) => {
    if (key === 'audit' && !canManageUsers) {
      notifyAccessDenied();
      setIsQuickNavOpen(false);
      return;
    }
    if (!canAccessEntries && key === 'entries') {
      notifyAccessDenied();
      setIsQuickNavOpen(false);
      return;
    }
    setTab(key);
    setSelectedVehicleId(null);
    setIsQuickNavOpen(false);
  };
  const normalizedVehicleSearch = normalizeVehicleSearchText(entryPlateSearch);
  const filteredEntryVehicles = filterVehicles(entryPlateSearch, data.vehicles);
  const normalizedMaintenanceSearch = normalizeVehicleSearchText(maintenanceVehicleSearch);
  const filteredMaintenanceVehicles = filterVehicles(maintenanceVehicleSearch, data.vehicles);
  const maintenanceDropdownVehicles = filteredMaintenanceVehicles;
  const entryDropdownVehicles = filteredEntryVehicles;
  const visibleMaintenanceVehicleId = maintenanceDropdownVehicles.some(({ id }) => id === maintenanceForm.vehicleId) ? maintenanceForm.vehicleId : '';
  const visibleEntryVehicleId = entryDropdownVehicles.some(({ id }) => id === entryForm.vehicleId) ? entryForm.vehicleId : '';

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
          <button className="fleet-brand fleet-brand-button" type="button" aria-expanded={window.matchMedia('(min-width: 769px)').matches ? isQuickNavOpen : undefined} aria-haspopup={window.matchMedia('(min-width: 769px)').matches ? 'menu' : undefined} onClick={() => { if (window.matchMedia('(min-width: 769px)').matches) setIsQuickNavOpen((open) => !open); }} aria-label="Abou Ghaly Muraqabat">
            <span className="fleet-brand-mark"><img src="/icon-speedometer.svg" alt="" /></span>
            <span className="fleet-brand-copy"><strong>Abou Ghaly</strong><small>Muraqabat</small></span>
          </button>
          {isQuickNavOpen && <nav className="fleet-quick-nav" aria-label="اختيار القسم" role="menu">
            <button type="button" className={tab === 'home' ? 'is-active' : ''} role="menuitem" onClick={() => { setTab('home'); setSelectedVehicleId(null); setIsQuickNavOpen(false); }}><Home aria-hidden="true" /><span>الرئيسية</span></button>
            {navigationTiles.map(({ key, label, Icon }) => <button key={key} type="button" className={tab === key ? 'is-active' : ''} role="menuitem" onClick={() => navigateToSection(key)}><Icon aria-hidden="true" /><span>{label}</span></button>)}
            {canManageUsers && <button type="button" className={tab === 'users' ? 'is-active' : ''} role="menuitem" onClick={() => { setTab('users'); setSelectedVehicleId(null); setIsQuickNavOpen(false); }}><Settings aria-hidden="true" /><span>إدارة الزوار</span></button>}
          </nav>}
        </div>
        <div className="fleet-header-actions">
          <div className="fleet-notification-anchor">
            <button
              className="fleet-header-icon-button fleet-notification-button"
              type="button"
              title="الإشعارات"
              aria-label={unreadNotificationCount ? `الإشعارات، ${unreadNotificationCount} جديد` : 'الإشعارات'}
              aria-expanded={isDeviationNotificationPanelOpen}
              aria-haspopup="dialog"
              onClick={() => {
                enableNotificationAlerts();
                setIsDeviationNotificationPanelOpen((open) => !open);
              }}
            >
              <Bell aria-hidden="true" />
              {unreadNotificationCount > 0 && <span className="fleet-notification-count">{unreadNotificationCount > 99 ? '99+' : unreadNotificationCount}</span>}
            </button>
            {isDeviationNotificationPanelOpen && <section className="fleet-notification-panel" role="dialog" aria-label="الإشعارات">
              <div className="fleet-notification-panel-heading">
                <strong>الإشعارات</strong>
                {unreadNotificationCount > 0 && <button className="fleet-notification-read-all" type="button" disabled={isMarkingAllNotificationsRead || refreshingNotificationId !== null} onClick={() => void markAllFleetNotificationsRead()}>
                  <CheckCheck size={15} aria-hidden="true" />
                  <span>{isMarkingAllNotificationsRead ? 'جارٍ القراءة...' : 'قراءة الكل'}</span>
                </button>}
                <span>{unreadNotificationCount ? `${unreadNotificationCount} جديد` : deviationNotifications.length ? `${deviationNotifications.length} في السجل` : 'لا توجد إشعارات جديدة'}</span>
              </div>
              {notificationListItems.length ? <div className="fleet-notification-list">
                {notificationListItems.map((item) => {
                  if (item.kind === 'deviation-group') {
                    const isRead = item.notifications.every(({ id }) => dismissedDeviationNotificationsRef.current.has(id));
                    const isExpanded = expandedNotificationIds.has(item.id);
                    return <article
                      key={item.id}
                      className={`fleet-notification-item is-deviation ${isRead ? 'is-read' : ''} ${isExpanded ? 'is-expanded' : ''}`}
                    >
                      <button
                        className="fleet-notification-item-header"
                        type="button"
                        aria-expanded={isExpanded}
                        onClick={() => toggleNotificationDetails(item.id)}
                      >
                        <span className="fleet-notification-item-icon"><Bell size={14} aria-hidden="true" /></span>
                        <strong>{item.plates.length} سيارات فيها انحراف كبير</strong>
                        <ChevronDown className="fleet-notification-chevron" size={16} aria-hidden="true" />
                      </button>
                      {isExpanded && <div className="fleet-notification-item-details">
                        {item.notifications.map((notification) => (
                          <div className="fleet-notification-detail" key={notification.id}>
                            <div className="fleet-notification-detail-copy">
                              <strong>سيارة <b className="fleet-notification-plate">{notification.plate}</b></strong>
                              <span>انحراف {notification.deviation.toLocaleString('ar-EG')} جنيه · {notification.month}</span>
                              <small className="fleet-notification-timestamp">{formatNotificationTimestamp(notification.createdAt)}</small>
                            </div>
                            <button className="fleet-notification-item-action" type="button" onClick={() => {
                              dismissFleetNotification(notification);
                              setIsDeviationNotificationPanelOpen(false);
                              openDeviationNotification(notification);
                            }}>عرض التفاصيل</button>
                          </div>
                        ))}
                        <small className="fleet-notification-timestamp fleet-notification-group-timestamp">{formatNotificationTimestamp(item.createdAt)}</small>
                      </div>}
                    </article>;
                  }
                  const { notification } = item;
                  const isRead = dismissedDeviationNotificationsRef.current.has(notification.id);
                  const isExpanded = expandedNotificationIds.has(notification.id);
                  const fleetUpdateSummaryParts = isFleetUpdateNotification(notification)
                    ? notification.summary.match(/^(.*السيارة\s+)(.+?)(\s+بتاريخ.*)$/)
                    : null;
                  return <article
                    key={notification.id}
                    className={`fleet-notification-item ${isFleetUpdateNotification(notification) ? 'is-data-update' : 'is-deviation'} ${isRead ? 'is-read' : ''} ${isExpanded ? 'is-expanded' : ''}`}
                  >
                    <button
                      className="fleet-notification-item-header"
                      type="button"
                      aria-expanded={isExpanded}
                      onClick={() => toggleNotificationDetails(notification.id)}
                    >
                      <span className="fleet-notification-item-icon">{refreshingNotificationId === notification.id ? <RefreshCw size={14} aria-hidden="true" /> : <Bell size={14} aria-hidden="true" />}</span>
                      <strong>{isFleetUpdateNotification(notification) ? notification.title : notification.deviation > 250 ? `انحراف خطر — السيارة ${notification.plate}` : `تحذير انحراف — السيارة ${notification.plate}`}</strong>
                      <ChevronDown className="fleet-notification-chevron" size={16} aria-hidden="true" />
                    </button>
                    {isExpanded && <div className="fleet-notification-item-details">
                      {isFleetUpdateNotification(notification)
                        ? <div className="fleet-notification-detail-copy">
                          <span>{fleetUpdateSummaryParts
                            ? <>{fleetUpdateSummaryParts[1]}<b className="fleet-notification-plate">{fleetUpdateSummaryParts[2]}</b>{fleetUpdateSummaryParts[3]}</>
                            : notification.summary}</span>
                          <small className="fleet-notification-timestamp">{formatNotificationTimestamp(notification.createdAt)}</small>
                        </div>
                        : <div className="fleet-notification-detail-copy">
                          <strong>سيارة <b className="fleet-notification-plate">{notification.plate}</b></strong>
                          <span>انحراف {notification.deviation.toLocaleString('ar-EG')} جنيه · {notification.month}</span>
                          <small className="fleet-notification-timestamp">{formatNotificationTimestamp(notification.createdAt)}</small>
                        </div>}
                      <button
                        className="fleet-notification-item-action"
                        type="button"
                        disabled={refreshingNotificationId === notification.id || isMarkingAllNotificationsRead}
                        onClick={() => {
                          if (isFleetUpdateNotification(notification)) {
                            void refreshFleetFromNotification(notification);
                            return;
                          }
                          setIsDeviationNotificationPanelOpen(false);
                          openDeviationNotification(notification);
                        }}
                      >
                        {refreshingNotificationId === notification.id ? 'جاري التحديث...' : isFleetUpdateNotification(notification) ? 'تحديث البيانات' : 'عرض التفاصيل'}
                      </button>
                    </div>}
                  </article>;
                })}
              </div> : <p className="fleet-notification-empty">ستظهر هنا إشعارات تحديث البيانات والانحرافات الجديدة.</p>}
            </section>}
          </div>
          {currentUser.isGuest && <span className="fleet-user-greeting">مرحباً، {currentUser.guestName}</span>}
          <span className="fleet-mobile-user" title={currentUser.employeeName || currentUser.guestName || currentUser.username}>
            {currentUser.employeeName || currentUser.guestName || currentUser.username}
          </span>
          <span className="fleet-live"><i /> متابعة مباشرة</span>
          <div className="fleet-header-desktop-actions">
            <div className="fleet-theme-dropdown-anchor">
              <button className="fleet-theme-button" type="button" aria-expanded={isThemePickerOpen} onClick={() => setIsThemePickerOpen((open) => !open)}>ألوان التطبيق</button>
              {isThemePickerOpen && <div className="fleet-theme-picker" role="group" aria-label="إعدادات المظهر">{renderThemeOptions()}{renderFontSizePicker()}</div>}
            </div>
            {canManageUsers && <button className="fleet-header-icon-button fleet-header-settings" type="button" title="إدارة الزوار" aria-label="إدارة الزوار" onClick={() => setTab('users')}><Settings aria-hidden="true" /><span>إدارة الزوار</span></button>}
            <button className="fleet-header-icon-button fleet-header-update" type="button" title="تحديث النظام" aria-label="تحديث النظام" onClick={() => void updateSystem()}><RefreshCw aria-hidden="true" /><span>تحديث النظام</span></button>
            {(isInstallAvailable || isAppInstalled) && <button className="fleet-header-icon-button fleet-header-install" type="button" title={isAppInstalled ? 'التطبيق مثبت' : 'تثبيت التطبيق'} aria-label={isAppInstalled ? 'التطبيق مثبت' : 'تثبيت التطبيق'} onClick={installApp} disabled={isAppInstalled}><Download aria-hidden="true" /><span>{isAppInstalled ? 'مثبت' : 'تثبيت التطبيق'}</span></button>}
            {onLogout && <button className="fleet-header-icon-button fleet-header-logout" type="button" title="تسجيل الخروج" aria-label="تسجيل الخروج" onClick={onLogout}><LogOut aria-hidden="true" /><span>تسجيل الخروج</span></button>}
          </div>
          <div className="fleet-mobile-actions" ref={mobileActionsRef}>
            <button
              className="fleet-mobile-actions-toggle"
              type="button"
              aria-label="المزيد من الخيارات"
              aria-expanded={isMobileActionsOpen}
              aria-haspopup="menu"
              onClick={() => setIsMobileActionsOpen((open) => !open)}
            ><MoreVertical aria-hidden="true" /></button>
            {isMobileActionsOpen && <div className="fleet-mobile-actions-menu" role="menu" aria-label="خيارات التطبيق">
              <div className="fleet-mobile-theme-anchor">
                <button type="button" role="menuitem" aria-expanded={isThemePickerOpen} onClick={() => setIsThemePickerOpen((open) => !open)}>ألوان التطبيق</button>
                {isThemePickerOpen && <div className="fleet-theme-picker" role="group" aria-label="إعدادات المظهر">{renderThemeOptions()}{renderFontSizePicker()}</div>}
              </div>
              {canManageUsers && <button type="button" role="menuitem" onClick={() => { setTab('users'); setIsMobileActionsOpen(false); }}><Settings aria-hidden="true" /><span>إدارة الزوار</span></button>}
              <button type="button" role="menuitem" onClick={() => { void updateSystem(); setIsMobileActionsOpen(false); }}><RefreshCw aria-hidden="true" /><span>تحديث النظام</span></button>
              {(isInstallAvailable || isAppInstalled) && <button type="button" role="menuitem" disabled={isAppInstalled} onClick={() => { void installApp(); setIsMobileActionsOpen(false); }}><Download aria-hidden="true" /><span>{isAppInstalled ? 'التطبيق مثبت' : 'تثبيت التطبيق'}</span></button>}
              {onLogout && <button type="button" role="menuitem" className="is-logout" onClick={() => { setIsMobileActionsOpen(false); onLogout(); }}><LogOut aria-hidden="true" /><span>تسجيل الخروج</span></button>}
            </div>}
          </div>
        </div>
      </header>

      {tab === 'home' ? <nav className={`fleet-tabs fleet-tabs-home${currentUser.role === 'viewer' ? ' fleet-tabs-home-viewer' : ''}`} aria-label="أقسام متابعة السيارات">
        {homeNavigationTiles.map(({ key, label, description, Icon }) => (
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
          <button type="button" className="fleet-section-title-button" onClick={() => { setTab('home'); setSelectedVehicleId(null); }}>{openSection.label}</button>
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

      {tab === 'audit' && canManageUsers && <section className="fleet-section fleet-audit-section">
        <div className="fleet-section-heading">
          <div><span className="fleet-overline">سجل التعديلات والحذف</span><h2>سجل العمليات</h2><p>أحدث 100 عملية تعديل أو حذف مسجلة.</p></div>
          <button type="button" className="fleet-link-button" onClick={() => setTab('home')}>العودة للرئيسية</button>
        </div>
        {auditLogError
          ? <div className="fleet-user-permission-note" role="alert">تعذر تحميل سجل العمليات: {auditLogError}</div>
          : auditEntries.length
            ? <div className="fleet-user-table-wrap"><table className="fleet-table fleet-user-table fleet-audit-table">
              <thead><tr><th>التاريخ والوقت</th><th>المستخدم</th><th>العملية</th><th>العنصر</th></tr></thead>
              <tbody>{auditEntries.map((entry) => <tr key={entry.id}>
                <td>{new Intl.DateTimeFormat('ar-EG', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(entry.createdAt))}</td>
                <td><strong>{entry.actorName}</strong>{entry.actorEmail && <small className="fleet-audit-email" dir="ltr">{entry.actorEmail}</small>}</td>
                <td><span className={`fleet-audit-action is-${entry.action}`}>{entry.action === 'edit' ? 'تعديل' : 'حذف'}</span></td>
                <td>{entry.itemName}</td>
              </tr>)}</tbody>
            </table></div>
            : <p className="fleet-empty-state">لا توجد عمليات مسجلة حتى الآن.</p>}
      </section>}

      {tab === 'routes' && <section className="fleet-section route-tracking-section">
        <div className="route-section-heading">
          <div><span className="fleet-overline">بيان حركة السيارات</span><h2>خط سير السيارات اليومي</h2></div>
          <div className="route-section-count"><strong>{numberFormat.format(savedRouteVehicleCount)} / {numberFormat.format(routeBoardVehicles.length)}</strong><span>سيارة محفوظة في هذا اليوم</span></div>
        </div>
        {!routesCloudLoaded
          ? <div className="route-storage-status" role="status">جاري تحميل سجلات خطوط السير...</div>
          : (!routesCloudAvailable || !isFirebaseConfigured) && <div className="route-storage-status" role="status">{routesCloudAvailable ? 'السجلات محفوظة على هذا الجهاز فقط؛ المزامنة السحابية غير مفعلة.' : 'تعذر الاتصال بسجل خط السير السحابي؛ السجلات الحالية محفوظة على هذا الجهاز فقط.'}</div>}
        {!canEditFleet && <div className="fleet-user-permission-note">يمكنك مشاهدة سجلات خط السير. التعديل والإضافة متاحان للمسؤول فقط.</div>}

        <section className="route-daily-sheet" aria-label="كشف خطوط السير">
          <div className="route-daily-toolbar">
            <div><span className="fleet-overline">بيان اليوم</span><h3>حركة السيارات المسجلة</h3></div>
            {canEditFleet && <button type="button" className="fleet-button fleet-button-primary route-add-vehicle-button" onClick={() => {
              setRouteVehicleDialog({ vehicleId: null, isDailyOnly: true, plate: '', name: '' });
              setRouteVehicleDialogError('');
            }}><Plus size={16} aria-hidden="true" /> إضافة سيارة للبيان</button>}
            <label className="route-date-filter">
              <span>تاريخ البيان</span>
              <div className="entry-date-input-wrap route-date-input-wrap">
                <input type="text" inputMode="numeric" dir="ltr" maxLength={10} placeholder="يوم/شهر/سنة" autoComplete="off" value={routeDateFilterDraft} aria-label="تاريخ البيان، يوم شهر سنة" onChange={(event) => {
                  const value = event.target.value;
                  const parsedDate = parseDriverSettlementDate(value);
                  setRouteDateFilterDraft(value);
                  if (parsedDate) setRouteDateFilter(parsedDate);
                }} onBlur={() => {
                  const parsedDate = parseDriverSettlementDate(routeDateFilterDraft);
                  if (parsedDate) setRouteDateFilterDraft(formatEntryDateInput(parsedDate));
                }} />
                <button type="button" className="entry-date-picker-button" aria-label="اختيار تاريخ البيان من التقويم" title="اختيار تاريخ البيان من التقويم" onClick={() => {
                  const picker = routeDatePickerRef.current;
                  if (!picker) return;
                  if (typeof picker.showPicker === 'function') picker.showPicker();
                  else {
                    picker.focus();
                    picker.click();
                  }
                }}><CalendarDays size={18} aria-hidden="true" /></button>
                <input ref={routeDatePickerRef} className="entry-date-native-picker" type="date" aria-label="تقويم تاريخ البيان" tabIndex={-1} value={routeDateFilter} onChange={(event) => {
                  if (!event.target.value) return;
                  setRouteDateFilter(event.target.value);
                  setRouteDateFilterDraft(formatEntryDateInput(event.target.value));
                }} />
              </div>
            </label>
          </div>
          {routeDateFilterInvalid && <p className="route-date-error" role="alert">أدخل تاريخًا صحيحًا بصيغة يوم/شهر/سنة.</p>}
          {routeBoardVehicles.length === 0 ? <div className="fleet-empty-state route-empty-state"><MapPinned aria-hidden="true" /><strong>لا توجد سيارات مسجلة لعرض البيان</strong></div> : <div className="route-daily-cards">
            {routeBoardVehicles.map((vehicle) => {
                const entry = routeEntriesForDate.find((item) => !item.isDailyVehicle && (
                  vehicle.fleetVehicleId
                    ? item.vehicleId === vehicle.fleetVehicleId
                    : item.dailyVehicleId === vehicle.id || (!item.vehicleId && item.plate.trim() === vehicle.plate.trim())
                ));
                const key = `${routeDateFilter}|${vehicle.id}`;
                const draft = routeDailyDrafts[key] || {
                  departure: entry?.departure || '',
                  route: entry?.route || '',
                  representative: entry?.representative || '',
                  driver: entry?.driver || '',
                };
                const updateDraft = (field: keyof RouteDailyDraft, value: string) => setRouteDailyDrafts((current) => ({
                  ...current,
                  [key]: { ...(current[key] || draft), [field]: value },
                }));
                const isExpanded = expandedRouteVehicleIds.has(vehicle.id);
                const panelId = `route-daily-fields-${vehicle.id}`;
                return <article className={`route-daily-card${isExpanded ? ' is-expanded' : ''}`} key={vehicle.id}>
                  <header className="route-daily-card-header">
                  <button type="button" className="route-daily-card-toggle" aria-expanded={isExpanded} aria-controls={panelId} onClick={() => setExpandedRouteVehicleIds((current) => {
                    const next = new Set(current);
                    if (next.has(vehicle.id)) next.delete(vehicle.id);
                    else next.add(vehicle.id);
                    return next;
                  })}>
                    <span className="route-daily-card-vehicle"><strong dir="auto">{vehicle.plate}</strong><span>{vehicle.name}</span></span>
                    <ChevronDown className="route-daily-card-chevron" size={19} aria-hidden="true" />
                  </button>
                  {canEditFleet && <button type="button" className="route-daily-card-edit" aria-label={`تعديل بيانات ${vehicle.plate} في البيان فقط`} title="تعديل الاسم والوصف في البيان فقط" onClick={() => {
                    setRouteVehicleDialog({ vehicleId: vehicle.id, isDailyOnly: vehicle.isDailyOnly, plate: vehicle.plate, name: vehicle.name });
                    setRouteVehicleDialogError('');
                  }}><Pencil size={14} aria-hidden="true" /></button>}
                  </header>
                  {isExpanded && <div className="route-daily-card-content" id={panelId}>
                    <div className="route-daily-card-fields">
                      <label><span>السائق</span>{canEditFleet ? <input className="route-daily-input" aria-label={`السائق لسيارة ${vehicle.plate}`} value={draft.driver} onChange={(event) => updateDraft('driver', event.target.value)} /> : <strong>{draft.driver || '—'}</strong>}</label>
                      <label><span>نقطة الانطلاق</span>{canEditFleet ? <div className="route-daily-location-field"><input className="route-daily-input" aria-label={`نقطة الانطلاق لسيارة ${vehicle.plate}`} value={draft.departure} onChange={(event) => updateDraft('departure', event.target.value)} /><button type="button" aria-label={`فتح الخريطة لنقطة الانطلاق ${vehicle.plate}`} title="فتح في خرائط Google" onClick={() => window.open(`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(draft.departure)}`, '_blank', 'noopener,noreferrer')}><MapPin size={16} aria-hidden="true" /></button></div> : <strong>{draft.departure || '—'}</strong>}</label>
                      <label><span>خط السير</span>{canEditFleet ? <div className="route-daily-location-field"><input className="route-daily-input" aria-label={`خط السير لسيارة ${vehicle.plate}`} value={draft.route} onChange={(event) => updateDraft('route', event.target.value)} /><button type="button" aria-label={`فتح الخريطة لخط سير ${vehicle.plate}`} title="فتح في خرائط Google" onClick={() => window.open(`https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(draft.route)}`, '_blank', 'noopener,noreferrer')}><MapPin size={16} aria-hidden="true" /></button></div> : <strong>{draft.route || '—'}</strong>}</label>
                      <label><span>المندوب</span>{canEditFleet ? <input className="route-daily-input" aria-label={`المندوب لسيارة ${vehicle.plate}`} value={draft.representative} onChange={(event) => updateDraft('representative', event.target.value)} /> : <strong>{draft.representative || '—'}</strong>}</label>
                    </div>
                    {canEditFleet && <button type="button" className="fleet-button fleet-button-primary route-daily-save" disabled={!routesCloudLoaded || routeDateFilterInvalid || savingRouteVehicleIds.has(vehicle.id)} onClick={() => void saveRouteVehicle(vehicle)}>{savingRouteVehicleIds.has(vehicle.id) ? 'جارٍ الحفظ...' : 'حفظ'}</button>}
                  </div>}
                </article>;
              })}
          </div>}
          <div className="route-export-footer">
            <button className="fleet-button fleet-button-light route-export-button" type="button" onClick={() => void exportRouteEntries()} disabled={isExportingRoutes || routeDateFilterInvalid || !routeBoardVehicles.length}><Download aria-hidden="true" />{isExportingRoutes ? 'جاري تجهيز Excel...' : 'تصدير إلى Excel'}</button>
          </div>
        </section>
        {routeVehicleDialog && canEditFleet && <div className="fleet-modal-backdrop route-vehicle-modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setRouteVehicleDialog(null)}>
          <form className="fleet-modal route-vehicle-modal" onSubmit={saveRouteVehicleMetadata} role="dialog" aria-modal="true" aria-labelledby="route-vehicle-dialog-title">
            <div className="fleet-modal-title">
              <div><span className="fleet-overline">بيان حركة السيارات</span><h2 id="route-vehicle-dialog-title">{routeVehicleDialog.vehicleId ? 'تعديل سيارة في البيان' : 'إضافة سيارة للبيان'}</h2></div>
              <button type="button" aria-label="إغلاق" onClick={() => setRouteVehicleDialog(null)}>×</button>
            </div>
            <p className="route-vehicle-modal-note">التعديلات هنا خاصة ببيان الحركة اليومي ولا تغيّر سجل السيارات في بقية التطبيق.</p>
            <label>رقم السيارة<input autoFocus required maxLength={80} value={routeVehicleDialog.plate} onChange={(event) => setRouteVehicleDialog((current) => current ? { ...current, plate: event.target.value } : current)} /></label>
            <label>الوصف / المنطقة<input required maxLength={100} value={routeVehicleDialog.name} onChange={(event) => setRouteVehicleDialog((current) => current ? { ...current, name: event.target.value } : current)} placeholder="مثال: دمياط" /></label>
            {routeVehicleDialogError && <p className="route-vehicle-modal-error" role="alert">{routeVehicleDialogError}</p>}
            <div className="fleet-user-edit-actions">
              <button className="fleet-button fleet-button-primary" type="submit">{routeVehicleDialog.vehicleId ? 'حفظ التعديل' : 'إضافة للبيان'}</button>
              <button className="fleet-button fleet-button-light" type="button" onClick={() => setRouteVehicleDialog(null)}>إلغاء</button>
            </div>
          </form>
        </div>}
        <section className="route-daily-sheet" aria-label="سجلات GPS">
          <div className="route-daily-toolbar">
            <div><span className="fleet-overline">سجلات الموقع</span><h3>سجلات GPS</h3></div>
            <span className="entries-pagination-status">{numberFormat.format(gpsLogs.length)} سجل محمّل</span>
          </div>
          {isLoadingGpsLogs
            ? <div className="route-storage-status" role="status">جاري تحميل أحدث سجلات GPS...</div>
            : gpsLogsError
              ? <div className="route-storage-status" role="alert">{gpsLogsError}</div>
              : gpsLogs.length === 0
                ? <div className="fleet-empty-state route-empty-state"><MapPinned aria-hidden="true" /><strong>لا توجد سجلات GPS</strong></div>
                : <div className="fleet-table-wrap route-table-wrap">
                  <table className="fleet-table route-table">
                    <thead><tr><th>التاريخ والوقت</th><th>السائق</th><th>السيارة</th><th>الموقع</th><th>الإحداثيات</th><th>الدقة</th></tr></thead>
                    <tbody>{gpsLogs.map((log) => {
                      const vehicle = data.vehicles.find((item) => item.id === log.vehicleId);
                      return <tr key={log.id}>
                        <td>{new Intl.DateTimeFormat('ar-EG', { dateStyle: 'short', timeStyle: 'short' }).format(new Date(log.timestamp))}</td>
                        <td>{log.driver}</td>
                        <td>{vehicle?.plate || '—'}</td>
                        <td>{log.note || '—'}</td>
                        <td dir="ltr">{log.latitude.toFixed(6)}, {log.longitude.toFixed(6)}</td>
                        <td>{log.accuracy === undefined ? '—' : `${Math.round(log.accuracy)} م`}</td>
                      </tr>;
                    })}</tbody>
                  </table>
                </div>}
          {!isLoadingGpsLogs && hasMoreGpsLogs && <div className="route-load-more">
            <button className="fleet-button fleet-button-light" type="button" onClick={() => void loadOlderGpsLogs()} disabled={isLoadingMoreGpsLogs}>
              {isLoadingMoreGpsLogs ? 'جاري تحميل السجلات...' : 'تحميل المزيد'}
            </button>
          </div>}
        </section>
      </section>}

      {tab === 'overview' && <>
        <section className={`ledger-sheet ledger-report-sheet ledger-summary-block${isFollowupSummaryCollapsed ? ' is-collapsed' : ''}`} aria-label="ملخص أداء كشف المتابعة">
          <div
            className={`ledger-titlebar ledger-report-titlebar ledger-summary-titlebar${isFollowupSummaryCollapsed ? ' is-collapsed' : ''}`}
            onClick={() => setIsFollowupSummaryCollapsed((collapsed) => !collapsed)}
          >
            <div className="ledger-sheet-title">
              <span>{`ملخص الأداء · ${new Intl.DateTimeFormat('ar-EG', { month: 'long', year: 'numeric' }).format(new Date(`${followupSummaryMonth}-01T00:00:00`))}`}</span>
              <button type="button" className="ledger-sheet-toggle" aria-expanded={!isFollowupSummaryCollapsed} aria-controls="ledger-followup-summary-content" onClick={(event) => { event.stopPropagation(); setIsFollowupSummaryCollapsed((collapsed) => !collapsed); }}>
                <strong>صورة سريعة للمتابعة</strong><ChevronDown aria-hidden="true" />
              </button>
            </div>
            <span className="ledger-period">{numberFormat.format(followupSummaryEntries.length)} حركة مسجلة</span>
          </div>
          {!isFollowupSummaryCollapsed && <>
            <div className="ledger-summary-heading-controls">
              <MonthYearPicker
                selectedMonth={followupSummaryMonth}
                onSelectMonth={(monthYear) => {
                  setFollowupSummaryMonth(monthYear);
                  setIsFollowupSummaryMonthPickerOpen(false);
                }}
                ariaLabel="اختيار شهر وسنة ملخص المتابعة"
                isOpen={isFollowupSummaryMonthPickerOpen}
                onOpenChange={setIsFollowupSummaryMonthPickerOpen}
                isActive
                className="ledger-summary-month-picker"
                portalStyle={monthPickerPortalStyle}
              />
            </div>
            <div id="ledger-followup-summary-content" className="ledger-summary ledger-expanded-content">
            <div><span>عدد الحركات المسجلة</span><strong>{numberFormat.format(followupSummaryEntries.length)}</strong><small>ضمن الشهر المحدد</small></div>
            <div><span>المسافة الكلية</span><strong>{numberFormat.format(followupSummaryTotals.distance)} <small>كم</small></strong><small>ضمن الشهر المحدد</small></div>
            <div><span>الوقود المصروف</span><strong>{numberFormat.format(followupSummaryTotals.liters)} <small>لتر</small></strong><small>ضمن الشهر المحدد</small></div>
            <div><span>إجمالي المصروف</span><strong>{currencyFormat.format(followupSummaryTotals.cost)}</strong><small>وقود ومصاريف إضافية</small></div>
            <div><span>انحراف التموين</span><strong className={deviationClass(followupSummaryTotals.deviation)}>{currencyFormat.format(followupSummaryTotals.deviation)}</strong><small>مقارنة بالسعر المعتمد</small></div>
            </div>
          </>}
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
          {!isLedgerSheetCollapsed && <div id="ledger-sheet-content" className="ledger-expanded-content">
            <section className="ledger-filters" aria-label="فلترة كشف الوقود">
              <div className="ledger-filters-heading"><h2>فلترة سجل الوقود</h2></div>
              <div className="ledger-period-controls" role="group" aria-label="الفترة الزمنية لسجل الوقود">
                <button type="button" data-period="all" className={ledgerPeriod === 'all' ? 'is-active' : ''} aria-pressed={ledgerPeriod === 'all'} onClick={() => selectLedgerPeriod('all')}>كل السجلات</button>
                <MonthYearPicker
                  selectedMonth={ledgerPeriod === 'month' ? ledgerMonthYear : null}
                  onSelectMonth={selectLedgerMonthYear}
                  ariaLabel="اختيار الشهر والسنة"
                  isOpen={isLedgerMonthPickerOpen}
                  onOpenChange={setIsLedgerMonthPickerOpen}
                  isActive={ledgerPeriod === 'month'}
                  portalStyle={monthPickerPortalStyle}
                />
                <button type="button" data-period="custom" className={ledgerPeriod === 'custom' ? 'is-active' : ''} aria-pressed={ledgerPeriod === 'custom'} onClick={() => selectLedgerPeriod('custom')}>تاريخ معين</button>
                <button className="ledger-filter-reset" type="button" onClick={resetLedgerFilters}><span aria-hidden="true">↶</span> إعادة تعيين</button>
              </div>
              <div className={`ledger-filter-grid${ledgerPeriod === 'custom' ? ' ledger-filter-grid-custom' : ''}`}>
                <label className="ledger-filter-vehicle"><span>السيارة</span><select required value={ledgerSummaryVehicleId} disabled={!data.vehicles.length} onChange={(event) => setLedgerSummaryVehicleId(event.target.value)} aria-label="اختر سيارة لكشف المتابعة"><option value="" disabled>{data.vehicles.length ? 'اختر سيارة' : 'لا توجد سيارات'}</option>{data.vehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</select></label>
                {ledgerPeriod === 'custom' && <>
                  <div className="ledger-filter-date-range" role="group" aria-label="نطاق التاريخ">
                    <label><span>بداية الفترة</span><input type="text" inputMode="numeric" dir="ltr" placeholder="يوم/شهر/سنة" aria-label="بداية الفترة، يوم شهر سنة" value={ledgerFromDateDraft} onChange={(event) => {
                      const value = event.target.value;
                      setLedgerFromDateDraft(value);
                      const parsedDate = parseDateFilter(value);
                      if (parsedDate !== null) setLedgerFilters((current) => ({ ...current, from: parsedDate }));
                    }} /></label>
                    <label><span>نهاية الفترة</span><input type="text" inputMode="numeric" dir="ltr" placeholder="يوم/شهر/سنة" aria-label="نهاية الفترة، يوم شهر سنة" value={ledgerToDateDraft} onChange={(event) => {
                      const value = event.target.value;
                      setLedgerToDateDraft(value);
                      const parsedDate = parseDateFilter(value);
                      if (parsedDate !== null) setLedgerFilters((current) => ({ ...current, to: parsedDate }));
                    }} /></label>
                  </div>
                </>}
                <label className="ledger-filter-search"><span>بحث في السجل</span><input type="search" value={ledgerFilters.search} onChange={(event) => {
                  const query = event.target.value;
                  setLedgerFilters((current) => ({ ...current, search: query }));
                  const uniqueMatch = findUniqueVehicleMatch(query, data.vehicles);
                  if (uniqueMatch) setLedgerSummaryVehicleId(uniqueMatch.id);
                }} placeholder="رقم المستند أو السيارة أو اللوحة" /></label>
              </div>
            </section>
            <div className="fleet-table-wrap">
              <table className="fleet-table ledger-table">
                <thead><tr><th>م</th><th>رقم المستند</th><th>رقم السيارة</th><th>السائق</th><th>تاريخ التعبئة</th><th>بداية العداد</th><th>قراءة العداد</th><th>الفجوة (كم)</th><th>الوقود المتوقع</th><th>الوقود المصروف</th><th>انحراف التموين (جنيه)</th><th>إجمالي المصروف</th></tr></thead>
                <tbody>{ledgerEntries.map((entry, index) => {
                  const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
                  const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
                  const expected = vehicle ? litersForDistance(vehicle, distance) : 0;
                  const liters = litersForEntry(entry, vehicle);
                  const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters);
                  return <tr key={entry.id} className="fleet-vehicle-tone" style={vehicleToneStyle(entry.vehicleId, data.vehicles, normalizeVehicleTypeColors(data.vehicleTypeColors))}>
                    <td className="ledger-index">{numberFormat.format(index + 1)}</td>
                    <td>{entry.documentNumber || '—'}</td>
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
          <div className={`ledger-titlebar monthly-summary-titlebar${isMonthlySummaryCollapsed ? ' is-collapsed' : ''}`}>
            <div className="monthly-summary-title">
              <button type="button" className="monthly-summary-toggle" aria-expanded={!isMonthlySummaryCollapsed} aria-controls="monthly-summary-content" onClick={() => setIsMonthlySummaryCollapsed((collapsed) => !collapsed)}>
                <span>ملخص شهري</span><ChevronDown aria-hidden="true" />
              </button>
              <span>تجميع مصروفات السيارات حسب الشهر</span>
            </div>
            <div className="monthly-summary-controls">
              <span className="ledger-period">{numberFormat.format(monthlyVehicleSummaries.reduce((count, summary) => count + summary.entries.length, 0))} حركة</span>
              {!isMonthlySummaryCollapsed && <MonthYearPicker
                selectedMonth={monthlySummaryMonth}
                isOpen={isMonthlySummaryPickerOpen}
                onOpenChange={setIsMonthlySummaryPickerOpen}
                onSelectMonth={(selectedMonth) => {
                  monthlySummaryMonthSelectedRef.current = true;
                  setMonthlySummaryMonth(selectedMonth);
                  setDeviationSummaryVehicleId('all');
                  setIsMonthlySummaryPickerOpen(false);
                }}
                ariaLabel="اختيار شهر الملخص"
                portalStyle={monthPickerPortalStyle}
              />}
            </div>
          </div>
          {!isMonthlySummaryCollapsed && <div id="monthly-summary-content" className="ledger-expanded-content">
            <div className="fleet-table-wrap monthly-summary-table-wrap">
              <table className="fleet-table ledger-table monthly-summary-table">
                <thead><tr><th title="رقم السيارة">رقم السيارة</th><th title="إجمالي التموين الفعلي">التموين الفعلي</th><th title="إجمالي المصروفات الإضافية">المصروفات الإضافية</th><th title="إجمالي مصروفات السيارة">إجمالي المصروفات</th><th title="إجمالي قيمة التموين المقررة">التموين المقرر</th><th title="إجمالي الانحراف">الانحراف</th><th title="إجمالي كم/لتر">كم/لتر</th></tr></thead>
                <tbody>{monthlyVehicleSummaries.map((summary) => (
                  <tr key={summary.vehicle.id} className="fleet-vehicle-tone" style={vehicleToneStyle(summary.vehicle.id, data.vehicles, normalizeVehicleTypeColors(data.vehicleTypeColors))}>
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
              <span>السيارات التي تجاوز انحرافها الإجمالي 250 خلال الفترة المختارة</span>
            </div>
            <div className="monthly-summary-controls deviation-summary-controls">
              <div className="deviation-summary-period-controls">
                <span className="ledger-period deviation-summary-count">{numberFormat.format(displayedDeviationSummaries.length)} سيارة</span>
              {!isDeviationSummaryCollapsed && <div className="deviation-month-picker deviation-month-multiselect">
                <button
                  type="button"
                  className="ledger-month-select is-active"
                  aria-label="اختيار شهور ملخص الانحرافات"
                  aria-haspopup="dialog"
                  aria-expanded={isDeviationSummaryPickerOpen}
                  onClick={() => setIsDeviationSummaryPickerOpen((open) => !open)}
                >
                  {deviationMonthLabel}
                </button>
                {isDeviationSummaryPickerOpen && createPortal(
                  <div
                    className={`deviation-month-modal-backdrop${themeKey === 'dark-premium' ? ' fleet-theme-dark' : ''}`}
                    dir="rtl"
                    style={{ '--ink': theme.ink, '--surface': theme.surface, '--line': theme.line, '--accent': theme.accent, '--fleet-font-scale': fontSizeKey === 'small' ? '.88' : fontSizeKey === 'large' ? '1.12' : '1' } as React.CSSProperties}
                    role="presentation"
                    onMouseDown={(event) => event.target === event.currentTarget && setIsDeviationSummaryPickerOpen(false)}
                  >
                    <div className="deviation-month-multiselect-panel" role="dialog" aria-modal="true" aria-labelledby="deviation-month-picker-title">
                      <div className="deviation-month-multiselect-heading">
                        <div>
                          <h2 id="deviation-month-picker-title">اختيار الشهور</h2>
                          <p>اختر حتى 3 أشهر لعرض إجمالي الانحرافات</p>
                        </div>
                        <button type="button" onClick={() => { setSelectedDeviationMonths([]); setDeviationSummaryVehicleId('all'); }}>كل الأشهر</button>
                      </div>
                      <div className="deviation-month-options">
                        {availableDeviationMonths.map((month) => (
                          <label key={month} className={`deviation-month-option${selectedDeviationMonths.includes(month) ? ' is-selected' : ''}`}>
                            <input
                              type="checkbox"
                              checked={selectedDeviationMonths.includes(month)}
                              disabled={!selectedDeviationMonths.includes(month) && selectedDeviationMonths.length >= 3}
                              onChange={(event) => {
                                setSelectedDeviationMonths((current) => {
                                  const next = event.target.checked
                                    ? current.length < 3 ? [...current, month] : current
                                    : current.filter((selected) => selected !== month);
                                  return next.sort((left, right) => right.localeCompare(left));
                                });
                                setDeviationSummaryVehicleId('all');
                              }}
                            />
                            <span>{deviationMonthDisplayName(month)}</span>
                          </label>
                        ))}
                        {availableDeviationMonths.length === 0 && <span className="deviation-month-empty">لا توجد أشهر مسجلة.</span>}
                      </div>
                      <div className="deviation-month-modal-actions">
                        <span>{selectedDeviationMonths.length ? `${selectedDeviationMonths.length} من 3 أشهر` : 'كل الأشهر محددة'}</span>
                        <button type="button" className="deviation-month-confirm" onClick={() => setIsDeviationSummaryPickerOpen(false)}>تم</button>
                      </div>
                    </div>
                  </div>,
                  document.body,
                )}
              </div>}
              </div>
              {!isDeviationSummaryCollapsed && <label className="deviation-summary-filter deviation-summary-vehicle-filter"><span>السيارة</span>
                <select value={deviationSummaryVehicleId} onChange={(event) => setDeviationSummaryVehicleId(event.target.value)} aria-label="تصفية ملخص الانحرافات حسب السيارة">
                  <option value="all">كل السيارات المتجاوزة</option>
                  {displayedDeviationSummaries.map(({ vehicle }) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate}</option>)}
                </select>
              </label>}
            </div>
          </div>
          {!isDeviationSummaryCollapsed && <div id="deviation-summary-content" className="ledger-expanded-content">
            {visibleDeviationSummaries.length > 0 ? <div className="fleet-table-wrap deviation-summary-table-wrap">
              <table className={`fleet-table ledger-table deviation-summary-table${selectedDeviationMonths.length > 0 ? ' has-notes' : ''}`}>
                <thead><tr><th>رقم السيارة</th><th>عدد الحركات</th><th>إجمالي الانحراف</th>{selectedDeviationMonths.length > 0 && <th>ملاحظات</th>}</tr></thead>
                <tbody>{visibleDeviationSummaries.map((summary) => {
                  const matchingNotes = selectedDeviationMonths.flatMap((month) => {
                    const note = deviationNotes[month]?.[summary.vehicle.id];
                    return note ? [{ month, note }] : [];
                  });
                  const isHiddenFromUsers = hiddenDeviationVehicleIds.has(summary.vehicle.id);
                  return <tr key={summary.vehicle.id} className={`fleet-vehicle-tone${isHiddenFromUsers ? ' is-hidden-from-users' : ''}`} style={vehicleToneStyle(summary.vehicle.id, data.vehicles, normalizeVehicleTypeColors(data.vehicleTypeColors))}>
                    <td className="ledger-plate">
                      <button type="button" className="ledger-vehicle-link" onClick={() => { setSelectedVehicleId(summary.vehicle.id); setTab('vehicles'); }}>{summary.vehicle.plate}</button>
                      {canEditFleet && <div className="deviation-visibility-control">
                        {isHiddenFromUsers && <span className="deviation-hidden-label">مخفي للمستخدمين</span>}
                        <button type="button" className="deviation-visibility-toggle" aria-label={`${isHiddenFromUsers ? 'إظهار' : 'إخفاء'} انحراف السيارة ${summary.vehicle.plate} ${isHiddenFromUsers ? 'للمستخدمين' : 'عن المستخدمين'}`} onClick={() => toggleDeviationVehicleVisibility(summary.vehicle.id)}>{isHiddenFromUsers ? 'إظهار' : 'إخفاء'}</button>
                      </div>}
                    </td>
                    <td>{numberFormat.format(summary.entries.length)}</td>
                    <td className="ledger-warning deviation-summary-amount">{currencyFormat.format(Math.round(summary.deviation))}</td>
                    {selectedDeviationMonths.length > 0 && <td className="deviation-note-cell">
                      <div className="deviation-note-content">
                        {matchingNotes.length > 0
                          ? matchingNotes.map(({ month, note }) => <div className="deviation-month-note" key={month}>
                            {selectedDeviationMonths.length > 1 && <strong>{new Intl.DateTimeFormat('ar-EG', { month: 'short', year: 'numeric' }).format(new Date(`${month}-01T00:00:00`))}</strong>}
                            <span>{note.text}</span>
                            <small>{new Intl.DateTimeFormat('ar-EG', { dateStyle: 'medium' }).format(new Date(note.createdAt))}</small>
                          </div>)
                          : <span className="deviation-note-pending">الانحراف قيد المراجعة والفحص</span>}
                        {canEditFleet && selectedDeviationMonths.length === 1 && <button type="button" className="deviation-note-edit" onClick={() => openDeviationNoteEditor(summary.vehicle.id)}>{matchingNotes.length ? 'تعديل' : 'إضافة ملاحظة'}</button>}
                      </div>
                    </td>}
                  </tr>;
                })}</tbody>
                <tfoot><tr><th colSpan={2}>إجمالي الانحراف للسيارات المعروضة</th><th className="ledger-warning">{currencyFormat.format(Math.round(visibleDeviationTotal))}</th>{selectedDeviationMonths.length > 0 && <th />}</tr></tfoot>
              </table>
            </div> : <p className="fleet-empty-state">{deviationSummaries.length === 0 ? 'لا توجد سيارات تجاوز انحرافها الإجمالي 250 في هذا الشهر.' : displayedDeviationSummaries.length === 0 ? 'لا توجد انحرافات ظاهرة للمستخدمين في هذه الفترة.' : 'لا توجد نتائج للسيارة المحددة.'}</p>}
          </div>}
        </section>
      </>}

      {tab === 'vehicles' && <section className={`fleet-section${selectedSummary ? ' fleet-detail-page' : ''}`}>
        {currentUser.role === 'admin' && <section className={`fleet-global-rate fleet-rate-compact fleet-rate-collapsible${isRateCardExpanded ? ' is-expanded' : ' is-collapsed'}`} aria-label="سعر التموين الموحد">
          <button type="button" className="fleet-rate-collapse-toggle" aria-expanded={isRateCardExpanded} onClick={(event) => {
            if (event.detail === 0 || event.detail > 1) setIsRateCardExpanded((expanded) => !expanded);
          }}>
            <span>سعر اللتر المعتمد</span><strong>{rateFormat.format(fuelAllowanceRate(data))} <small>جنيه / لتر</small></strong><ChevronDown aria-hidden="true" />
          </button>
          {isRateCardExpanded && <>
            <div><strong>سعر اللتر المعتمد</strong><span>يُطبق تلقائيًا على حسابات جميع السيارات</span></div>
            {canEditFleet
              ? !isRateEditorOpen
                ? <button type="button" className="fleet-rate-value" onClick={() => setIsRateEditorOpen(true)}>{rateFormat.format(fuelAllowanceRate(data))} <small>جنيه / لتر</small><b>تغيير</b></button>
                : <div className="fleet-rate-editor"><label><input aria-label="سعر التموين المعتمد بالجنيه لكل لتر" type="number" min="0.01" step="0.01" value={allowanceRateDraft} onChange={(event) => setAllowanceRateDraft(event.target.value)} /><span>جنيه / لتر</span></label><button type="button" onClick={saveAllowanceRate}>حفظ</button><button type="button" className="fleet-rate-cancel" onClick={() => setIsRateEditorOpen(false)}>إلغاء</button></div>
              : <div className="fleet-rate-value" aria-label="سعر اللتر المعتمد للقراءة فقط">{rateFormat.format(fuelAllowanceRate(data))} <small>جنيه / لتر</small></div>}
          </>}
        </section>}
        {selectedSummary ? <>
          <div className="fleet-section-heading fleet-detail-heading">
            <div className="fleet-detail-vehicle-info">
              <span className="fleet-detail-vehicle-icon"><CarFront aria-hidden="true" /></span>
              <div className="fleet-detail-vehicle-copy">
                <h2>{selectedSummary.vehicle.name}</h2>
                <span className="fleet-overline">{selectedSummary.vehicle.plate} · {vehicleTypeLabel(selectedSummary.vehicle)}</span>
              </div>
              <div className="fleet-detail-vehicle-controls">
                <label className="fleet-detail-vehicle-switcher">
                  <span>التبديل إلى سيارة</span>
                  <select
                    aria-label="التبديل إلى سيارة أخرى"
                    value={selectedSummary.vehicle.id}
                    onChange={(event) => setSelectedVehicleId(event.target.value)}
                  >
                    {data.vehicles.map((vehicle) => (
                      <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicleTypeLabel(vehicle)}</option>
                    ))}
                  </select>
                </label>
                <button className="fleet-detail-all-vehicles" type="button" onClick={() => setSelectedVehicleId(null)}>
                  <CarFront size={16} aria-hidden="true" />
                  <span>عرض كل السيارات</span>
                </button>
              </div>
            </div>
            {canEditFleet && <div className="fleet-detail-actions">
              <button className="fleet-button fleet-button-light fleet-icon-action fleet-detail-action-edit" type="button" title="تعديل السيارة" aria-label="تعديل السيارة" onClick={() => openVehicleEditor(selectedSummary.vehicle)}><Car size={15} aria-hidden="true" /></button>
              <button className="fleet-button fleet-button-light fleet-icon-action fleet-detail-action-download" type="button" title="تنزيل قالب السيارة" aria-label="تنزيل قالب السيارة" onClick={() => void downloadSelectedVehicleTemplate().catch((error: unknown) => setNotice(error instanceof Error ? `تعذر إنشاء قالب السيارة: ${error.message}` : 'تعذر إنشاء قالب السيارة.'))}><Download size={15} aria-hidden="true" /></button>
              <button className="fleet-button fleet-button-light fleet-icon-action fleet-detail-action-upload" type="button" title="استيراد بيانات السيارة" aria-label="استيراد بيانات السيارة" disabled={isImportingVehicleHistory} onClick={() => vehicleHistoryImportRef.current?.click()}><Upload size={15} aria-hidden="true" /></button>
              <input ref={vehicleHistoryImportRef} className="fleet-file-input" type="file" accept=".xlsx,.csv" aria-label="اختر قالب السيارة للاستيراد" onChange={importSelectedVehicleTemplate} />
              <button className="fleet-button fleet-button-primary fleet-icon-action fleet-detail-action-entry" type="button" title="تسجيل تعبئة" aria-label="تسجيل تعبئة" onClick={startEntry}><Plus size={17} aria-hidden="true" /></button>
            </div>}
          </div>
          <section className={`fleet-rate-summary${isRateSummaryExpanded ? ' is-expanded' : ' is-collapsed'}`} aria-label="ملخص المعدلات">
            <button type="button" className="fleet-rate-summary-toggle" aria-expanded={isRateSummaryExpanded} onClick={() => setIsRateSummaryExpanded((expanded) => !expanded)}>
              <span><Gauge aria-hidden="true" /> 📊 ملخص المعدلات <small className="fleet-rate-summary-count">عدد الحركات: {numberFormat.format(selectedEntries.length)}</small></span>
              <ChevronDown aria-hidden="true" />
            </button>
            {isRateSummaryExpanded && <div className="fleet-rate-summary-content">
            <div className="fleet-rate-summary-grid">
              <article className="fleet-rate-summary-card">
                <Gauge aria-hidden="true" />
                <span>متوسط الاستهلاك</span>
                <strong>{rateFormat.format(selectedAverageRate)} <small>كم/لتر</small></strong>
              </article>
              <article className="fleet-rate-summary-card is-best">
                <TrendingUp aria-hidden="true" />
                <span>أفضل استهلاك</span>
                <strong>{rateFormat.format(selectedBestRate)} <small>كم/لتر</small></strong>
              </article>
              <article className="fleet-rate-summary-card is-worst">
                <TrendingDown aria-hidden="true" />
                <span>أسوأ استهلاك</span>
                <strong>{rateFormat.format(selectedWorstRate)} <small>كم/لتر</small></strong>
              </article>
              <article className="fleet-rate-summary-card">
                <MapPinned aria-hidden="true" />
                <span>إجمالي المسافة</span>
                <strong>{numberFormat.format(selectedRateTotals.distance)} <small>كم</small></strong>
              </article>
              <article className="fleet-rate-summary-card">
                <Fuel aria-hidden="true" />
                <span>إجمالي الوقود (لتر)</span>
                <strong>{numberFormat.format(selectedRateTotals.actualLiters)}</strong>
              </article>
              <article className="fleet-rate-summary-card">
                <Calculator aria-hidden="true" />
                <span>مسافة الأساس</span>
                <strong>{numberFormat.format(selectedCalculatedBaseDistance)} <small>كم</small></strong>
              </article>
            </div>
            <div className="fleet-rate-summary-insights">
              <section className="fleet-rate-chart-panel" aria-label="معدل الاستهلاك خلال آخر الحركات">
                <h4>معدل الاستهلاك خلال آخر 12 حركة</h4>
                {chartRateMovements.length
                  ? <div className="fleet-rate-chart" role="img" aria-label={`معدلات الاستهلاك لآخر ${chartRateMovements.length} حركات`}>
                    {chartRateMovements.map(({ entry, rate }) => (
                      <div className="fleet-rate-chart-item" key={entry.id} title={`${formatDateArabic(entry.date)}: ${rateFormat.format(rate)} كم/لتر`}>
                        <span className="fleet-rate-chart-value">{rateFormat.format(rate)}</span>
                        <span
                          className={`fleet-rate-chart-bar ${rate > chartAverageRate ? 'is-good' : 'is-poor'}`}
                          style={{ '--rate-bar-height': `${Math.max(4, (rate / chartMaximumRate) * 100)}%` } as React.CSSProperties}
                        />
                        <span className="fleet-rate-chart-date">{entry.date.slice(5)}</span>
                      </div>
                    ))}
                  </div>
                  : <p className="fleet-rate-chart-empty">لا توجد حركات لها معدل استهلاك في هذه السيارة.</p>}
              </section>
              <section className="fleet-rate-month-comparison" aria-label="مقارنة معدل الاستهلاك الشهري">
                <h4>الشهر الماضي ({comparisonMonthLabel(lastMonthStart)}) vs الشهر الأسبق ({comparisonMonthLabel(monthBeforeLastStart)})</h4>
                <div className="fleet-rate-month-values">
                  <div><span>الشهر الماضي · كامل</span><strong>{rateFormat.format(lastMonthRateSummary.rate)} <small>كم/لتر</small></strong></div>
                  <div><span>الشهر الأسبق · كامل</span><strong>{rateFormat.format(monthBeforeLastRateSummary.rate)} <small>كم/لتر</small></strong></div>
                </div>
                <p className={lastMonthRateSummary.rate === 0 ? 'is-unavailable' : monthlyRateChange === null ? 'is-unavailable' : monthlyRateChange > 0 ? 'is-improved' : monthlyRateChange < 0 ? 'is-declined' : 'is-unchanged'}>
                  {lastMonthRateSummary.rate === 0
                    ? 'لا توجد حركات مسجلة للشهر الماضي'
                    : monthlyRateChange === null
                    ? 'لا توجد بيانات كافية للمقارنة'
                    : monthlyRateChange > 0
                      ? `تحسن ${rateFormat.format(monthlyRateChange)}٪`
                      : monthlyRateChange < 0
                        ? `تدهور ${rateFormat.format(Math.abs(monthlyRateChange))}٪`
                        : 'لا يوجد تغير عن الشهر الأسبق'}
                </p>
              </section>
            </div>
            </div>}
          </section>
          {canEditFleet && <div className="fleet-detail-column-tools">
            <div className="entries-column-picker">
              <button type="button" className="fleet-button fleet-button-light entries-column-toggle" aria-expanded={isDetailColumnsPanelOpen} onClick={() => setIsDetailColumnsPanelOpen((open) => !open)}>
                <span aria-hidden="true">☷</span> الأعمدة
              </button>
              {isDetailColumnsPanelOpen && <>
                <button type="button" className="entries-column-backdrop" aria-label="إغلاق قائمة الأعمدة" onClick={() => setIsDetailColumnsPanelOpen(false)} />
                <div className="entries-column-panel" role="dialog" aria-modal="true" aria-label={`أعمدة تفاصيل ${selectedSummary.vehicle.name}`}>
                <div className="entries-column-header">
                  <strong>الأعمدة الظاهرة</strong>
                  <button type="button" aria-label="إغلاق قائمة الأعمدة" onClick={() => setIsDetailColumnsPanelOpen(false)}>×</button>
                </div>
                <div className="entries-column-list">
                  {detailColumns.map(({ key, label }) => <label key={key}>
                    <input type="checkbox" checked={selectedDetailColumnVisibility[key] !== false} onChange={(event) => setDetailColumnVisibility((preferences) => ({
                      ...preferences,
                      [key]: event.target.checked,
                    }))} />
                    {label}
                  </label>)}
                </div>
                <div className="entries-column-bulk-actions">
                  <button type="button" className="entries-column-reset" onClick={() => setDetailColumnVisibility({ ...defaultDetailColumnVisibility })}>إظهار الكل</button>
                  <button type="button" className="entries-column-reset" onClick={() => setDetailColumnVisibility(Object.fromEntries(detailColumns.map(({ key }) => [key, false])) as Record<DetailColumnKey, boolean>)}>إخفاء الكل</button>
                </div>
                </div>
              </>}
            </div>
            <span>إعداد عام لكل السيارات ويُحفظ تلقائيًا</span>
          </div>}
          <div className="fleet-table-wrap fleet-detail-table-wrap">
            <table className="fleet-table ledger-table fleet-detail-ledger">
              <thead><tr>{visibleDetailColumns.map(({ key, label }) => <th key={key} className={`fleet-detail-column-${key.replace(':', '-')}`}>{label}</th>)}{canEditFleet && <th>إجراء</th>}</tr></thead>
              <tbody>{selectedEntryDetails.map(({ entry, distance, actualFuelValue, deviation, accumulated, totalCost }) => {
                const liters = litersForEntry(entry, selectedSummary.vehicle);
                const values = new Map<DetailColumnKey, React.ReactNode>([
                  ['documentNumber', entry.documentNumber || '—'],
                  ['date', formatDateArabic(entry.date)],
                  ['driver', entry.driver || selectedSummary.vehicle.driver || '—'],
                  ['route', entry.route || '—'],
                  ['startOdometer', numberFormat.format(entry.startOdometer)],
                  ['currentOdometer', numberFormat.format(entry.currentOdometer)],
                  ['distance', `${numberFormat.format(distance)} كم`],
                  ['liters', `${numberFormat.format(liters)} لتر`],
                  ['allowance', numberFormat.format(Math.round(entryAllowanceValue(entry, fuelAllowanceRate(data), liters)))],
                  ['actual', numberFormat.format(Math.round(actualFuelValue))],
                  ['deviation', <span className={deviationClass(deviation)}>{numberFormat.format(Math.round(deviation))}</span>],
                  ['rate', <span className={accumulated > 200 ? 'ledger-warning' : 'ledger-ok'}>{rateFormat.format(accumulated)}</span>],
                  ...expenseFields.map((field) => [`expense:${field.key}`, numberFormat.format(entry.expenses?.[field.key] || 0)] as const),
                  ['totalCost', <>{numberFormat.format(Math.round(totalCost))}{entry.note ? <small className="ledger-row-note">{entry.note}</small> : null}</>],
                ]);
                  const cardKey = `detail:${entry.id}`;
                  const isCardExpanded = expandedMovementCards.has(cardKey);
                  return <tr key={entry.id} className={`fleet-vehicle-tone${isCardExpanded ? ' is-card-expanded' : ''}`} style={vehicleToneStyle(entry.vehicleId, data.vehicles, normalizeVehicleTypeColors(data.vehicleTypeColors))}>
                    <td className="ledger-card-summary">
                      <button type="button" aria-label={`${isCardExpanded ? 'طي' : 'فتح'} تفاصيل حركة ${entry.documentNumber || formatDateArabic(entry.date)}`} aria-expanded={isCardExpanded} onClick={() => setExpandedMovementCards((current) => {
                        const next = new Set(current);
                        if (next.has(cardKey)) next.delete(cardKey);
                        else next.add(cardKey);
                        return next;
                      })}>
                        <span className="ledger-card-document">مستند {entry.documentNumber || '—'}</span>
                        <span className={`ledger-card-deviation ${deviationClass(deviation)}`}>الانحراف: {currencyFormat.format(deviation)}</span>
                        <ChevronDown aria-hidden="true" />
                      </button>
                    </td>
                    {visibleDetailColumns.map(({ key }) => <td key={key} data-label={detailColumns.find((column) => column.key === key)?.label} className={`fleet-detail-cell-${key.replace(':', '-')}`}>{values.get(key)}</td>)}
                    {canEditFleet && <td className="ledger-actions" data-label="إجراء"><button type="button" className="entry-edit-button fleet-icon-action" aria-label="تعديل حركة الوقود" title="تعديل" onClick={() => editEntry(entry)}><Pencil size={15} aria-hidden="true" /><span className="fleet-movement-action-label">تعديل</span></button><button type="button" className="entry-delete-button fleet-icon-action" aria-label="حذف حركة الوقود" title="حذف" onClick={() => openDeleteConfirmation(entry)}><Trash2 size={15} aria-hidden="true" /><span className="fleet-movement-action-label">حذف</span></button></td>}
                  </tr>;
              })}</tbody>
              {selectedEntries.length > 0 && <tfoot><tr>{visibleDetailColumns.map(({ key }) => {
                const totalByColumn = new Map<DetailColumnKey, React.ReactNode>([
                  ['distance', `${numberFormat.format(selectedSummary.distance)} كم`],
                  ['liters', `${numberFormat.format(selectedSummary.liters)} لتر`],
                  ['allowance', numberFormat.format(Math.round(selectedFuelTotals.allowance))],
                  ['actual', numberFormat.format(Math.round(selectedFuelTotals.actual))],
                  ['deviation', numberFormat.format(Math.round(selectedFuelTotals.deviation))],
                  ['rate', rateFormat.format(totalKmPerLiter)],
                  ...expenseFields.map((field) => [`expense:${field.key}`, numberFormat.format(selectedEntries.reduce((sum, entry) => sum + (entry.expenses?.[field.key] || 0), 0))] as const),
                  ['totalCost', numberFormat.format(Math.round(selectedSummary.cost))],
                ]);
                const labelColumn = visibleDetailColumns.find(({ key: visibleKey }) => ['documentNumber', 'date', 'driver', 'route', 'startOdometer', 'currentOdometer'].includes(visibleKey))?.key || visibleDetailColumns[0]?.key;
                return <th key={key}>{key === labelColumn && <strong>إجمالي السيارة</strong>}{totalByColumn.get(key) && <div>{totalByColumn.get(key)}</div>}</th>;
              })}{canEditFleet && <th />}</tr></tfoot>}
            </table>
            {selectedEntries.length > 0 && <div className="fleet-detail-mobile-total" aria-label="إجمالي السيارة">
              <strong>إجمالي السيارة</strong>
              <span>المسافة<strong>{numberFormat.format(selectedSummary.distance)} كم</strong></span>
              <span>الوقود<strong>{numberFormat.format(selectedSummary.liters)} لتر</strong></span>
              <span>التموين المقرر<strong>{numberFormat.format(selectedFuelTotals.allowance)} جنيه</strong></span>
              <span>التموين الفعلي<strong>{numberFormat.format(selectedFuelTotals.actual)} جنيه</strong></span>
              <span>الانحراف<strong>{numberFormat.format(selectedFuelTotals.deviation)} جنيه</strong></span>
              <span>المعدل الحقيقي<strong>{rateFormat.format(totalKmPerLiter)} كم/لتر</strong></span>
              <span>إجمالي التكلفة<strong>{numberFormat.format(selectedSummary.cost)} جنيه</strong></span>
            </div>}
            {selectedEntries.length === 0 && <p className="fleet-empty-state">لا توجد مصروفات مسجلة لهذه السيارة بعد.</p>}
          </div>
        </> : <>
          <div className="fleet-section-heading fleet-vehicles-heading">
            <div><h2>سيارات الشراكة</h2></div>
            {canEditFleet && <div className="fleet-vehicle-actions">
              <button className="fleet-button fleet-button-light fleet-vehicle-action-icon fleet-download-vehicle-action" type="button" title="تنزيل قالب جميع السيارات" aria-label="تنزيل قالب جميع السيارات" onClick={() => void downloadAllVehicleTemplate().catch((error: unknown) => { setNoticeIsError(true); setNotice(error instanceof Error ? `تعذر تنزيل قالب جميع السيارات: ${error.message}` : 'تعذر تنزيل قالب جميع السيارات.'); })}><Download size={15} strokeWidth={2} aria-hidden="true" /></button>
                <button className="fleet-button fleet-button-light fleet-vehicle-action-icon fleet-import-vehicle-action" type="button" title={isImportingAllVehicleData ? 'جاري استيراد السيارات' : 'استيراد جميع السيارات'} aria-label={isImportingAllVehicleData ? 'جاري استيراد السيارات' : 'استيراد جميع السيارات'} disabled={isImportingAllVehicleData} onClick={() => allVehiclesImportRef.current?.click()}>{isImportingAllVehicleData ? <RefreshCw size={15} strokeWidth={2} aria-hidden="true" /> : <Upload size={15} strokeWidth={2} aria-hidden="true" />}</button>
                <input ref={allVehiclesImportRef} className="fleet-file-input" type="file" accept=".xlsx" aria-label="اختر قالب جميع السيارات للاستيراد" onChange={importAllVehicleTemplate} />
              <button className="fleet-button fleet-button-primary fleet-vehicle-action-icon fleet-add-vehicle-action" type="button" title="إضافة سيارة" aria-label="إضافة سيارة" onClick={() => setVehicleDialog(true)}><CarFront size={15} strokeWidth={2} aria-hidden="true" /><Plus className="fleet-add-vehicle-plus" size={11} strokeWidth={3} aria-hidden="true" /></button>
            </div>}
          </div>
          {canEditFleet && <div className="fleet-vehicle-type-colors" aria-label="ألوان أنواع السيارات">
            <span className="fleet-vehicle-type-colors-title">ألوان أنواع السيارات</span>
            <div className="fleet-vehicle-type-color-options">
              {vehicleTypeOptions.map(({ value, label }) => <label className="fleet-vehicle-type-color-option" key={value}>
                <input
                  type="color"
                  aria-label={`تغيير لون ${label}`}
                  value={normalizeVehicleTypeColors(data.vehicleTypeColors)[value]}
                  onChange={(event) => setData((current) => ({
                    ...current,
                    vehicleTypeColors: {
                      ...normalizeVehicleTypeColors(current.vehicleTypeColors),
                      [value]: event.target.value,
                    },
                  }))}
                />
                <span>{label}</span>
              </label>)}
            </div>
            <small>تظهر الألوان المحددة على السيارات لجميع المستخدمين، وتُحفظ تلقائيًا.</small>
          </div>}
          <div className="fleet-vehicle-grid">{summaries.map((summary) => {
            const { vehicle, lastOdometer, entries, distance, liters, expected, deviation, cost, kmPerLiter } = summary;
            const vehicleIndex = data.vehicles.findIndex((item) => item.id === vehicle.id);
            const VehicleIcon = [CarFront, Truck, BusFront, CarTaxiFront, Car][Math.max(0, vehicleIndex) % 5];
            const isExpanded = expandedVehicleCardId === vehicle.id;
            return <article
              className={`fleet-vehicle fleet-vehicle-card fleet-vehicle-tone${isExpanded ? ' is-expanded' : ''}`}
              style={vehicleToneStyle(vehicle.id, data.vehicles, normalizeVehicleTypeColors(data.vehicleTypeColors))}
              key={vehicle.id}
              onClick={(event) => {
                if ((event.target as HTMLElement).closest('button, input, select, textarea, a, label')) return;
                setExpandedVehicleCardId((current) => current === vehicle.id ? null : vehicle.id);
              }}
            >
              <div className="fleet-vehicle-summary-row">
                <button type="button" className="fleet-vehicle-summary-toggle" aria-label={`افتح تفاصيل السيارة ${vehicle.plate} بالنقر مرتين`} onClick={(event) => event.stopPropagation()} onDoubleClick={() => setSelectedVehicleId(vehicle.id)}>
                  <span className="fleet-car-icon"><VehicleIcon size={25} strokeWidth={2.1} aria-hidden="true" /></span>
                  <span className="fleet-vehicle-summary-copy"><strong>{vehicle.plate}</strong><small>{vehicle.name}</small></span>
                </button>
                {canEditFleet && <button type="button" className="fleet-vehicle-edit-icon" aria-label={`تعديل السيارة ${vehicle.name}`} title="تعديل السيارة" onClick={() => openVehicleEditor(vehicle)}><Settings size={14} strokeWidth={2.2} aria-hidden="true" /></button>}
                <span className="fleet-vehicle-summary-spacer" />
                {Math.abs(deviation) >= 0.5 && <span className={`fleet-gap ${deviation > 250 ? 'gap-high' : deviation < 0 ? 'gap-low' : 'gap-normal'}`}>الانحراف: {currencyFormat.format(deviation)}</span>}
              </div>
              {isExpanded && <div className="fleet-vehicle-expanded" id={`vehicle-card-details-${vehicle.id}`}>
                <div className="fleet-vehicle-data"><span>السائق<strong>{vehicle.driver || 'غير محدد'}</strong></span><span>معدل اللترات<strong>{numberFormat.format(vehicleFuelRate(vehicle))} لتر / {numberFormat.format(vehicleDistanceBase(vehicle))} كم</strong></span><span>العداد الحالي<strong>{numberFormat.format(lastOdometer)} كم</strong></span><span className="fleet-fixed-reading">{canEditFleet ? <label>القراءة السابقة<input type="number" min="0" value={openingOdometerDrafts[vehicle.id] ?? String(vehicle.openingOdometer)} aria-label={`القراءة السابقة لـ ${vehicle.name}`} onChange={(event) => setOpeningOdometerDrafts((current) => ({ ...current, [vehicle.id]: event.target.value }))} /><button type="button" onClick={() => saveOpeningOdometer(vehicle, entries.length)}>حفظ القراءة</button></label> : <>القراءة السابقة<strong>{numberFormat.format(vehicle.openingOdometer)} كم</strong></>}</span><span>المسافة المسجلة<strong>{numberFormat.format(distance)} كم</strong></span><span>كم/لتر الإجمالي<strong>{rateFormat.format(kmPerLiter)}</strong></span><span>إجمالي الانحراف<strong className={`fleet-gap${deviation > 250 ? ' gap-high' : deviation < 0 ? ' gap-low' : ''}`}>{currencyFormat.format(deviation)}</strong></span><span>إجمالي التكلفة<strong>{currencyFormat.format(cost)}</strong></span></div>
                <div className="fleet-card-actions"><button type="button" className="fleet-open-details" onClick={() => setSelectedVehicleId(vehicle.id)}>عرض تفاصيل السيارة <span aria-hidden="true">←</span></button></div>
              </div>}
            </article>;
          })}</div>
        </>}
      </section>}

      {tab === 'entries' && canAccessEntries && <section className="fleet-section fleet-entry-screen">
        <div className="fleet-section-heading entries-page-heading"><div><span className="fleet-overline">حركات الوقود والمصروفات</span><div className="entries-page-title"><h2>سجل التعبئة</h2>{canEditFleet && <button type="button" className="fleet-clear-entries-heading" aria-label="مسح الحركات" title="مسح الحركات" disabled={data.entries.length === 0 || !cloudLoaded || isClearingAllEntries} onClick={() => {
          setClearEntriesConfirmationPassword('');
          setClearEntriesMode('all');
          setClearEntriesMonth(dateInputValue(new Date()).slice(0, 7));
          setIsClearEntriesMonthPickerOpen(false);
          setIsClearEntriesConfirmationOpen(true);
        }}><Trash2 size={14} aria-hidden="true" /></button>}</div></div><span className="entries-count">عرض آخر {latestEntries.length} من {data.entries.length} حركة</span></div>
        <div className="entries-actions">
          {canEditFleet && <>
            <button className="fleet-button fleet-button-primary fleet-icon-button" type="button" aria-label="تسجيل تعبئة" title="تسجيل تعبئة" onClick={() => startEntry()}><Fuel size={19} aria-hidden="true" /></button>
            <button className="fleet-button fleet-button-light fleet-icon-button" type="button" aria-label="إضافة صيانة" title="إضافة صيانة" onClick={openMaintenanceDialog}><Wrench size={18} aria-hidden="true" /></button>
          </>}
          <button className="fleet-button fleet-button-light fleet-icon-button driver-settlement-open" type="button" aria-label="تصفية عهدة سائق" title="تصفية عهدة سائق" onClick={() => {
            setDriverSettlementDriver(driverSettlementDrivers[0] || '');
            setDriverSettlementFromDraft(formatEntryDateInput(driverSettlementFrom));
            setDriverSettlementToDraft(formatEntryDateInput(driverSettlementTo));
            setIsDriverSettlementOpen(true);
          }}><ClipboardList size={18} aria-hidden="true" /></button>
          <button className="fleet-button fleet-button-light fleet-icon-button" type="button" aria-label="تصدير Excel" title="تصدير Excel" onClick={() => void exportExcel()}><svg className="excel-export-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 2.75h9l5 5v13.5H5z" fill="#fff" stroke="#217346" strokeWidth="1.5" strokeLinejoin="round" /><path d="M14 2.75v5h5" fill="#d9f0e2" stroke="#217346" strokeWidth="1.5" strokeLinejoin="round" /><path d="M2 8h11v12H2z" rx="1.5" fill="#217346" /><path d="m5 11 5 6m0-6-5 6" fill="none" stroke="#fff" strokeWidth="1.7" strokeLinecap="round" /></svg></button>
          <div className="entries-column-picker">
            <button
              className="fleet-button fleet-button-light entries-column-toggle"
              type="button"
              aria-expanded={isEntryColumnsPanelOpen}
              aria-haspopup="dialog"
              onClick={() => setIsEntryColumnsPanelOpen((open) => !open)}
            >
              <Settings size={16} aria-hidden="true" /> الأعمدة
            </button>
            {isEntryColumnsPanelOpen && <>
              <button type="button" className="entries-column-backdrop" aria-label="إغلاق قائمة الأعمدة" onClick={() => setIsEntryColumnsPanelOpen(false)} />
              <div className="entries-column-panel" role="dialog" aria-modal="true" aria-label="إظهار وإخفاء أعمدة سجل التعبئة">
              <div className="entries-column-header">
                <strong>الأعمدة الظاهرة</strong>
                <button type="button" aria-label="إغلاق قائمة الأعمدة" onClick={() => setIsEntryColumnsPanelOpen(false)}>×</button>
              </div>
              <div className="entries-column-list">
                {entryColumns.map(({ key, label }) => <label key={key}>
                  <input
                    type="checkbox"
                    checked={entryColumnVisibility[key]}
                    onChange={(event) => setEntryColumnVisibility((current) => ({ ...current, [key]: event.target.checked }))}
                  />
                  <span>{label}</span>
                </label>)}
              </div>
              <div className="entries-column-bulk-actions">
                <button type="button" className="entries-column-reset" onClick={() => setEntryColumnVisibility(defaultEntryColumnVisibility)}>إظهار الكل</button>
                <button type="button" className="entries-column-reset" onClick={() => setEntryColumnVisibility(Object.fromEntries(entryColumns.map(({ key }) => [key, false])) as Record<EntryColumnKey, boolean>)}>إخفاء الكل</button>
              </div>
              </div>
            </>}
          </div>
        </div>
        <div className="fleet-table-wrap"><table className="fleet-table fleet-entry-table" style={{ '--entries-table-min-width': `${Math.max(520, visibleEntryColumns.length * 115 + (canEditFleet ? 100 : 0))}px` } as React.CSSProperties}><thead><tr>{visibleEntryColumns.map(({ key, label }) => <th key={key}>{label}</th>)}{canEditFleet && <th>إجراء</th>}</tr></thead><tbody>
          {latestEntries.map((entry) => {
            const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
            const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
            const expected = vehicle ? litersForDistance(vehicle, distance) : 0;
            const liters = litersForEntry(entry, vehicle);
            const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters);
            const cells: Record<EntryColumnKey, React.ReactNode> = {
              date: formatDateArabic(entry.date),
              documentNumber: entry.documentNumber || '—',
              vehicle: <><strong>{vehicle?.name || 'سيارة محذوفة'}</strong><small>{vehicle?.plate}</small></>,
              startOdometer: numberFormat.format(entry.startOdometer),
              currentOdometer: numberFormat.format(entry.currentOdometer),
              distance: `${numberFormat.format(distance)} كم`,
              expectedFuel: `${numberFormat.format(expected)} لتر`,
              fuel: `${numberFormat.format(liters)} لتر`,
              deviation: <td className={deviationClass(deviation)}>{currencyFormat.format(deviation)}</td>,
              totalCost: currencyFormat.format(entryTotalCost(entry)),
            };
            const cardKey = `entry:${entry.id}`;
            const isCardExpanded = expandedMovementCards.has(cardKey);
            return <tr key={entry.id} className={`fleet-vehicle-tone${isCardExpanded ? ' is-card-expanded' : ''}`} style={vehicleToneStyle(entry.vehicleId, data.vehicles, normalizeVehicleTypeColors(data.vehicleTypeColors))}>
              <td className="ledger-card-summary">
                <button type="button" aria-label={`${isCardExpanded ? 'طي' : 'فتح'} تفاصيل حركة ${entry.documentNumber || formatDateArabic(entry.date)}`} aria-expanded={isCardExpanded} onClick={() => setExpandedMovementCards((current) => {
                  const next = new Set(current);
                  if (next.has(cardKey)) next.delete(cardKey);
                  else next.add(cardKey);
                  return next;
                })}>
                  <span className="ledger-card-document">مستند {entry.documentNumber || '—'}</span>
                  <span className={`ledger-card-deviation ${deviationClass(deviation)}`}>الانحراف: {currencyFormat.format(deviation)}</span>
                  <ChevronDown aria-hidden="true" />
                </button>
              </td>
              {visibleEntryColumns.map(({ key }) => key === 'deviation'
                ? <React.Fragment key={key}>{React.cloneElement(cells[key] as React.ReactElement, { 'data-label': entryColumns.find((column) => column.key === key)?.label })}</React.Fragment>
                : <td key={key} data-label={entryColumns.find((column) => column.key === key)?.label}>{cells[key]}</td>)}
              {canEditFleet && <td className="ledger-actions" data-label="إجراء"><button type="button" className="entry-edit-button fleet-icon-action" aria-label="تعديل حركة الوقود" title="تعديل" onClick={() => editEntry(entry)}><Pencil size={15} aria-hidden="true" /><span className="fleet-movement-action-label">تعديل</span></button><button type="button" className="entry-delete-button fleet-icon-action" aria-label="حذف حركة الوقود" title="حذف" onClick={() => openDeleteConfirmation(entry)}><Trash2 size={15} aria-hidden="true" /><span className="fleet-movement-action-label">حذف</span></button></td>}
            </tr>;
          })}
          {latestEntries.length === 0 && <tr><td colSpan={visibleEntryColumns.length + (canEditFleet ? 1 : 0)} className="entries-pagination-empty">لا توجد حركات مسجلة.</td></tr>}
        </tbody></table></div>
      </section>}

      <footer className="fleet-footer"><span>مراقبة سيارات شركة أبو غالي</span></footer>

      {!entryDialog && <nav className="fleet-mobile-nav" aria-label="التنقل الرئيسي">
        {[
          { key: 'home' as const, label: 'الرئيسية', Icon: Home },
          { key: 'overview' as const, label: 'نظرة عامة', Icon: LayoutDashboard },
          { key: 'vehicles' as const, label: 'السيارات', Icon: Car },
          { key: 'entries' as const, label: 'التعبئة', Icon: Fuel },
          { key: 'routes' as const, label: 'GPS', Icon: MapPinned },
        ].map(({ key, Icon }) => (
          <button
            key={key}
            type="button"
            className={tab === key ? 'is-active' : ''}
            aria-current={tab === key ? 'page' : undefined}
            aria-label={key === 'overview' ? 'نظرة عامة' : key === 'home' ? 'الرئيسية' : key === 'vehicles' ? 'السيارات' : key === 'routes' ? 'GPS' : 'سجل التعبئة'}
            title={key === 'overview' ? 'نظرة عامة' : key === 'home' ? 'الرئيسية' : key === 'vehicles' ? 'السيارات' : key === 'routes' ? 'GPS' : 'سجل التعبئة'}
            onClick={() => {
              if (key === 'entries' && !canAccessEntries) {
                notifyAccessDenied();
                return;
              }
              setTab(key);
              if (key !== 'vehicles') setSelectedVehicleId(null);
              setIsQuickNavOpen(false);
            }}
          >
            <Icon aria-hidden="true" />
          </button>
        ))}
      </nav>}

      {editingDeviationNote && editingDeviationVehicle && <div className="deviation-note-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setEditingDeviationNote(null)}>
        <form className="deviation-note-modal" role="dialog" aria-modal="true" aria-labelledby="deviation-note-title" onSubmit={(event) => { event.preventDefault(); saveDeviationNote(); }}>
          <div className="fleet-modal-title">
            <div><span className="fleet-overline">ملاحظة شهرية · {editingDeviationNote.month}</span><h2 id="deviation-note-title">ملاحظة الانحراف</h2></div>
            <button type="button" aria-label="إغلاق" onClick={() => setEditingDeviationNote(null)}>×</button>
          </div>
          <dl className="deviation-note-vehicle">
            <div><dt>السيارة</dt><dd>{editingDeviationVehicle.name}</dd></div>
            <div><dt>رقم اللوحة</dt><dd>{editingDeviationVehicle.plate}</dd></div>
            <div><dt>السائق</dt><dd>{editingDeviationVehicle.driver || 'غير محدد'}</dd></div>
          </dl>
          <label className="deviation-note-input">الملاحظة
            <textarea autoFocus rows={5} value={deviationNoteDraft} onChange={(event) => setDeviationNoteDraft(event.target.value)} placeholder="اكتب ملاحظة هذا الشهر..." />
          </label>
          {deviationNotes[editingDeviationNote.month]?.[editingDeviationNote.vehicleId] && <p className="deviation-note-created">
            تاريخ الإنشاء: {new Intl.DateTimeFormat('ar-EG', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(deviationNotes[editingDeviationNote.month][editingDeviationNote.vehicleId].createdAt))}
          </p>}
          <div className="deviation-note-actions">
            <button type="button" className="fleet-button" onClick={() => setEditingDeviationNote(null)}>إلغاء</button>
            <button type="submit" className="fleet-button fleet-button-primary">حفظ الملاحظة</button>
          </div>
        </form>
      </div>}

      {vehicleDialog && <div className="fleet-modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setVehicleDialog(false)}><form className="fleet-modal" onSubmit={saveVehicle}><div className="fleet-modal-title"><div><span className="fleet-overline">بيانات المركبة</span><h2>{editingVehicleId ? 'تعديل بيانات السيارة' : 'إضافة سيارة'}</h2></div><button type="button" aria-label="إغلاق" onClick={() => { setVehicleDialog(false); setEditingVehicleId(null); }}>×</button></div><label>نوع السيارة<select required value={vehicleForm.vehicleType} onChange={(event) => { if (isVehicleType(event.target.value)) setVehicleForm({ ...vehicleForm, vehicleType: event.target.value }); }}>{vehicleTypeOptions.map(({ value, label }) => <option key={value} value={value}>{label}</option>)}</select></label><label>اسم السيارة<input required value={vehicleForm.name} onChange={(event) => setVehicleForm({ ...vehicleForm, name: event.target.value })} placeholder="مثال: تويوتا هايلوكس" /></label><label>رقم اللوحة<input required value={vehicleForm.plate} onChange={(event) => setVehicleForm({ ...vehicleForm, plate: event.target.value })} /></label><label>اسم السائق<input value={vehicleForm.driver} onChange={(event) => setVehicleForm({ ...vehicleForm, driver: event.target.value })} /></label><div className="fleet-form-row"><label>بداية العداد<input required type="number" min="0" value={vehicleForm.openingOdometer} onChange={(event) => setVehicleForm({ ...vehicleForm, openingOdometer: event.target.value })} /></label><label>اللترات المحددة<input required type="number" min="0.1" step="0.1" value={vehicleForm.kmPerLiter} onChange={(event) => setVehicleForm({ ...vehicleForm, kmPerLiter: event.target.value })} /></label></div><div className="fleet-form-row"><label>مسافة الأساس بالكيلومتر<input required type="number" min="1" step="1" value={vehicleForm.distanceBaseKm} onChange={(event) => setVehicleForm({ ...vehicleForm, distanceBaseKm: event.target.value })} /></label><span className="fleet-rate-help">المعادلة: الفجوة ÷ مسافة الأساس × اللترات</span></div><button className="fleet-button fleet-button-primary fleet-submit" type="submit">{editingVehicleId ? 'حفظ التعديل' : 'حفظ السيارة'}</button></form></div>}

      {maintenanceDialog && <div className="fleet-modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setMaintenanceDialog(false)}>
        <form className="fleet-modal" onSubmit={saveMaintenance}>
          <div className="fleet-modal-title"><div><span className="fleet-overline">مصروف مستقل</span><h2>إضافة صيانة</h2></div><button type="button" aria-label="إغلاق" onClick={() => setMaintenanceDialog(false)}>×</button></div>
          <label>السيارة<input type="search" value={maintenanceVehicleSearch} onChange={(event) => updateMaintenanceVehicleSearch(event.target.value)} placeholder="ابحث بالاسم أو رقم اللوحة" /><select required value={visibleMaintenanceVehicleId} onChange={(event) => setMaintenanceForm((current) => ({ ...current, vehicleId: event.target.value }))}>
            {maintenanceDropdownVehicles.length
              ? <>{!visibleMaintenanceVehicleId && <option value="" disabled>{normalizedMaintenanceSearch ? 'اختر سيارة من النتائج' : 'اختر سيارة'}</option>}{maintenanceDropdownVehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</>
              : <option value="" disabled>{data.vehicles.length ? 'لا توجد سيارات مطابقة' : 'لا توجد سيارات'}</option>}
          </select></label>
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
        <form className="fleet-modal fleet-entry-modal" noValidate onSubmit={saveEntry} onChange={() => { entryDraftSuppressedRef.current = false; if (entryFormError) setEntryFormError(''); }}>
          <header className="entry-form-header">
            <h2>{editingEntryId ? 'تعديل حركة مسجلة' : 'تسجيل تعبئة جديدة'}</h2>
            <button type="button" className="entry-page-back" aria-label="رجوع" title="رجوع" onClick={closeEntryPage}><ArrowLeft size={20} aria-hidden="true" /></button>
          </header>
          <div className="entry-form-body">
            <section className="entry-form-section">
              <div className="entry-section-title"><span>01</span><div><h3>بيانات العهدة</h3></div></div>
              <div className="entry-form-grid entry-trip-grid">
                <label className="entry-field entry-vehicle-search-field entry-editable-field"><span>اختيار السيارة</span><input type="search" value={entryPlateSearch} onChange={(event) => updateEntryVehicleSearch(event.target.value)} placeholder="ابحث بالاسم أو رقم اللوحة" /><select required value={visibleEntryVehicleId} onChange={(event) => selectEntryVehicle(event.target.value)}>
                  {entryDropdownVehicles.length
                    ? <>{!visibleEntryVehicleId && <option value="" disabled>{normalizedVehicleSearch ? 'اختر سيارة من النتائج' : 'اختر سيارة'}</option>}{entryDropdownVehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.plate} · {vehicle.name}</option>)}</>
                    : <option value="" disabled>{data.vehicles.length ? 'لا توجد سيارات مطابقة' : 'لا توجد سيارات'}</option>}
                </select></label>
                <label className="entry-field entry-compact-field entry-editable-field"><span>رقم المستند</span><input required value={entryForm.documentNumber} onChange={(event) => setEntryForm((current) => ({ ...current, documentNumber: event.target.value }))} placeholder="أدخل رقم مستند التعبئة" /><small className="entry-field-help">رقم المستند مطلوب ويمكن البحث به في السجلات</small></label>
                <label className="entry-field entry-date-field entry-editable-field">
                  <span>التاريخ</span>
                  <div className="entry-date-input-wrap">
                    <input required type="text" inputMode="numeric" dir="ltr" maxLength={10} placeholder="يوم/شهر/سنة" autoComplete="off" value={entryForm.date} onChange={(event) => setEntryForm((current) => ({ ...current, date: event.target.value }))} onBlur={() => {
                      const isoDate = parseEntryDateInput(entryForm.date);
                      if (isoDate) setEntryForm((current) => ({ ...current, date: formatEntryDateInput(isoDate) }));
                    }} />
                    <button type="button" className="entry-date-picker-button" aria-label="اختيار التاريخ من التقويم" title="اختيار التاريخ من التقويم" onClick={() => {
                      const picker = entryDatePickerRef.current;
                      if (!picker) return;
                      if (typeof picker.showPicker === 'function') picker.showPicker();
                      else {
                        picker.focus();
                        picker.click();
                      }
                    }}><CalendarDays size={18} aria-hidden="true" /></button>
                    <input ref={entryDatePickerRef} className="entry-date-native-picker" type="date" aria-label="التاريخ من التقويم" tabIndex={-1} value={parseEntryDateInput(entryForm.date) ?? ''} onChange={(event) => {
                      if (event.target.value) setEntryForm((current) => ({ ...current, date: formatEntryDateInput(event.target.value) }));
                    }} />
                  </div>
                  <small className="entry-field-help">أدخل يوم/شهر، وسيُكمل التطبيق السنة الحالية تلقائيًا، أو أدخل السنة كاملة</small>
                </label>
                <label className="entry-field entry-compact-field entry-editable-field"><span>اسم السائق</span><input value={entryForm.driver} onChange={(event) => setEntryForm({ ...entryForm, driver: event.target.value })} /></label>
                <label className="entry-field entry-compact-field entry-editable-field"><span>خط السير</span><input value={entryForm.route} onChange={(event) => setEntryForm({ ...entryForm, route: event.target.value })} placeholder="نقطة الانطلاق والوجهة" /></label>
              </div>
            </section>

            <section className="entry-form-section entry-fuel-section">
              <button type="button" className="entry-section-title entry-section-toggle" aria-expanded={isFuelSectionExpanded} aria-controls="entry-fuel-content" onClick={() => setIsFuelSectionExpanded((expanded) => !expanded)}>
                <span>02</span><div><h3>العداد والتموين</h3>{!isFuelSectionExpanded && entryOdometerSummary && <small className="entry-section-summary">{entryOdometerSummary}</small>}</div><ChevronDown aria-hidden="true" />
              </button>
              {isFuelSectionExpanded && <div className="entry-form-grid entry-fuel-grid" id="entry-fuel-content">
                <label className="entry-field entry-field-readonly"><span>القراءة السابقة · آخر قراءة</span><input required type="number" min="0" value={entryForm.startOdometer} readOnly /></label>
                <label className="entry-field entry-field-primary entry-editable-field"><span>القراءة الحالية</span><input required type="number" min="0" value={entryForm.currentOdometer} onChange={(event) => { const currentOdometer = event.target.value; setEntryForm({ ...entryForm, currentOdometer, liters: calculateFuelLiters(entryForm.vehicleId, Number(entryForm.startOdometer), Number(currentOdometer)).toFixed(1) }); }} /></label>
                <div className="entry-distance"><span>المسافة</span><strong>{numberFormat.format(entryDistance)} <small>كم</small></strong></div>
                <label className="entry-field entry-field-readonly"><span>التموين الفعلي (لتر)</span><input type="number" min="0" step="0.1" value={entryForm.liters} readOnly /><small className="entry-field-help">يُحسب من فرق العداد ومعدل استهلاك السيارة</small></label>
                <label className="entry-field entry-field-readonly"><span>قيمة التموين وفقًا للتر (جنيه)</span><input type="number" min="0" step="1" value={calculateAllowanceValue(entryForm.vehicleId, calculateFuelLiters(entryForm.vehicleId, Number(entryForm.startOdometer), Number(entryForm.currentOdometer)), editingEntryAllowancePrice)} readOnly /><small className="entry-field-help">{rateFormat.format(editingEntryAllowancePrice)} جنيه لكل لتر لجميع السيارات</small></label>
                <div className="entry-field entry-field-primary entry-editable-field">
                  <label className="entry-field-label" htmlFor="fuel-amount-input">مصروف الوقود الفعلي (جنيه)</label>
                  <div className={`entry-amount-total${fuelActualValuePreview !== null && fuelActualValuePreview > 0 ? ' has-value' : ''}`}><span>الإجمالي</span><output aria-live="polite">{fuelActualValuePreview === null ? '—' : currencyFormat.format(fuelActualValuePreview)}</output></div>
                  <div className="entry-amount-entry"><input id="fuel-amount-input" dir="ltr" type="text" inputMode="decimal" value={fuelAmountDraft} onChange={(event) => setFuelAmountDraft(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); addFuelAmount(); } }} /><button type="button" className="entry-amount-add" onClick={addFuelAmount}>إضافة</button></div>
                  <div className="entry-amount-list" aria-label="المبالغ المضافة لمصروف الوقود">{splitAmountExpression(entryForm.fuelActualValue).length
                    ? splitAmountExpression(entryForm.fuelActualValue).map((amount, index, amounts) => {
                      const isRecentlyAdded = recentlyAddedAmount?.target === `fuel:${index}` && index === amounts.length - 1;
                      return <span className={`entry-amount-chip${isRecentlyAdded ? ' is-recently-added' : ''}`} key={`${amount}-${index}-${isRecentlyAdded ? recentlyAddedAmount.token : ''}`}><span dir="ltr">{amount}</span><button type="button" aria-label={`حذف مبلغ ${amount} من الوقود`} onClick={() => removeFuelAmount(index)}>×</button></span>;
                    })
                    : <span className="entry-amount-empty">لا توجد مبالغ مضافة</span>}</div>
                  {fuelActualValuePreview === null && <small className="entry-field-help" role="alert">راجع المبالغ المضافة</small>}
                </div>
              </div>}
            </section>

            <section className="entry-form-section">
              <button type="button" className="entry-section-title entry-section-toggle" aria-expanded={isExpenseSectionExpanded} aria-controls="entry-expense-content" onClick={() => setIsExpenseSectionExpanded((expanded) => !expanded)}>
                <span>03</span><div><h3>المصاريف الإضافية</h3>{!isExpenseSectionExpanded && entryExpenseTotal > 0 && <small className="entry-section-summary">إجمالي المصاريف: {currencyFormat.format(entryExpenseTotal)}</small>}</div><ChevronDown aria-hidden="true" />
              </button>
              {isExpenseSectionExpanded && <div className="entry-expense-grid" id="entry-expense-content">{expenseFields.map((field) => {
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
                    <span className={`entry-expense-total${amount !== null && amount > 0 ? ' has-value' : ''}`}>{amount === null ? 'مبلغ غير صالح' : numberFormat.format(amount)}</span>
                    <span className="entry-expense-chevron" aria-hidden="true">{isExpanded ? '−' : '+'}</span>
                  </button>
                  {isExpanded && <div className="entry-expense-panel" id={`expense-panel-${field.key}`}>
                    <div className="entry-amount-entry"><input id={`expense-amount-${field.key}`} dir="ltr" type="text" inputMode="decimal" aria-label={`${field.label} بالجنيه`} value={expenseAmountDrafts[field.key]} onChange={(event) => setExpenseAmountDrafts((current) => ({ ...current, [field.key]: event.target.value }))} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); addExpenseAmount(field.key); } }} /><button type="button" className="entry-amount-add" onClick={() => addExpenseAmount(field.key)}>إضافة</button></div>
                    <div className="entry-amount-list" aria-label={`المبالغ المضافة: ${field.label}`}>{splitAmountExpression(entryForm.expenses[field.key]).length
                      ? splitAmountExpression(entryForm.expenses[field.key]).map((part, index, amounts) => {
                        const isRecentlyAdded = recentlyAddedAmount?.target === `expense:${field.key}:${index}` && index === amounts.length - 1;
                        return <span className={`entry-amount-chip${isRecentlyAdded ? ' is-recently-added' : ''}`} key={`${part}-${index}-${isRecentlyAdded ? recentlyAddedAmount.token : ''}`}><span dir="ltr">{part}</span><button type="button" aria-label={`حذف ${part} من ${field.label}`} onClick={() => removeExpenseAmount(field.key, index)}>×</button></span>;
                      })
                      : <span className="entry-amount-empty">لا توجد مبالغ مضافة</span>}</div>
                    {amount === null && <small className="entry-field-help" role="alert">راجع المبالغ المضافة</small>}
                  </div>}
                </div>;
              })}</div>}
            </section>
          </div>
          {entryFormError && <p className="entry-form-error" role="alert" aria-live="assertive">{entryFormError}</p>}
          <footer className="entry-form-actions"><div><span>إجمالي الحركة</span><strong>{hasInvalidAmount ? 'راجع المبالغ' : hasUnaddedAmounts ? 'اضغط إضافة أولًا' : currencyFormat.format(entryCostPreview)}</strong></div><button type="button" className="entry-cancel-button" onClick={closeEntryPage}>العودة</button>{!editingEntryId && hasEnteredEntryData && <button type="button" className="entry-clear-button" onClick={clearEntryForm}><RotateCcw size={15} aria-hidden="true" />مسح البيانات</button>}<button className="fleet-button fleet-button-primary entry-save-button" type="submit">{editingEntryId ? 'حفظ التعديل' : 'حفظ الحركة'}</button></footer>
        </form>
      </section>}
      {pendingDeleteEntry && <div className="fleet-confirm-backdrop" role="presentation"><div className="fleet-confirm-dialog" role="alertdialog" aria-modal="true" aria-labelledby="delete-entry-title"><span className="fleet-confirm-icon">!</span><h2 id="delete-entry-title">تأكيد حذف السجل</h2><p>لحذف حركة التعبئة بتاريخ <strong>{formatDateArabic(pendingDeleteEntry.date)}</strong> اكتب كلمة <strong>موافق</strong> ثم اضغط تأكيد.</p><input className="fleet-confirm-input" aria-label="اكتب موافق للتأكيد" value={deleteConfirmationText} onChange={(event) => setDeleteConfirmationText(event.target.value)} autoFocus placeholder="اكتب: موافق" /><div><button type="button" className="fleet-confirm-cancel" onClick={cancelDeleteConfirmation}>إلغاء</button><button type="button" className="fleet-confirm-ok" disabled={deleteConfirmationText.trim() !== 'موافق'} onClick={confirmDeleteEntry}>تأكيد الحذف</button></div></div></div>}
      {isClearEntriesConfirmationOpen && canEditFleet && <div className="fleet-confirm-backdrop" role="presentation"><div className="fleet-confirm-dialog clear-entries-confirm-dialog" role="alertdialog" aria-modal="true" aria-labelledby="clear-entries-title">
        <span className="fleet-confirm-icon">!</span>
        <h2 id="clear-entries-title">تأكيد مسح الحركات</h2>
        <p>{clearEntriesMode === 'all'
          ? 'سيتم مسح كل الحركات نهائيًا. ستبقى السيارات وإعدادات سعر اللتر كما هي.'
          : 'سيتم مسح الحركات المسجلة في الشهر المحدد فقط، وستبقى الحركات الأخرى والسيارات وإعدادات سعر اللتر كما هي.'}</p>
        <div className="clear-entries-mode-options" role="radiogroup" aria-label="نطاق مسح الحركات">
          <label className={clearEntriesMode === 'all' ? 'is-selected' : ''}>
            <input type="radio" name="clear-entries-mode" value="all" checked={clearEntriesMode === 'all'} onChange={() => setClearEntriesMode('all')} disabled={isClearingAllEntries} />
            <span>مسح كل الحركات</span>
          </label>
          <label className={clearEntriesMode === 'month' ? 'is-selected' : ''}>
            <input type="radio" name="clear-entries-mode" value="month" checked={clearEntriesMode === 'month'} onChange={() => setClearEntriesMode('month')} disabled={isClearingAllEntries} />
            <span>مسح حركات شهر معين</span>
          </label>
        </div>
        {clearEntriesMode === 'month' && <div className="clear-entries-month-field">
          <span>الشهر والسنة</span>
          <MonthYearPicker
            selectedMonth={clearEntriesMonth}
            onSelectMonth={setClearEntriesMonth}
            ariaLabel="اختيار شهر حركات التعبئة المراد مسحها"
            isOpen={isClearEntriesMonthPickerOpen}
            onOpenChange={setIsClearEntriesMonthPickerOpen}
            isActive
            className="clear-entries-month-picker"
            portalStyle={monthPickerPortalStyle}
          />
          <small>{numberFormat.format(data.entries.filter((entry) => entry.date.slice(0, 7) === clearEntriesMonth).length)} حركة في الشهر المحدد</small>
        </div>}
        <label className="clear-entries-password-label" htmlFor="clear-entries-password">اكتب كلمة التأكيد</label>
        <input id="clear-entries-password" className="fleet-confirm-input" type="password" autoComplete="new-password" aria-label="كلمة التأكيد لمسح الحركات" value={clearEntriesConfirmationPassword} onChange={(event) => setClearEntriesConfirmationPassword(event.target.value)} autoFocus placeholder="كلمة التأكيد" disabled={isClearingAllEntries} />
        <div><button type="button" className="fleet-confirm-cancel" disabled={isClearingAllEntries} onClick={() => {
          setIsClearEntriesConfirmationOpen(false);
          setClearEntriesConfirmationPassword('');
          setIsClearEntriesMonthPickerOpen(false);
        }}>إلغاء</button><button type="button" className="fleet-confirm-ok" disabled={isClearingAllEntries || !isClearEntriesPasswordValid || !cloudLoaded || (clearEntriesMode === 'month' && !data.entries.some((entry) => entry.date.slice(0, 7) === clearEntriesMonth))} onClick={() => void clearAllEntries()}>{isClearingAllEntries ? 'جارٍ المسح...' : 'تأكيد المسح'}</button></div>
      </div></div>}
      {isDriverSettlementOpen && <div className="fleet-modal-backdrop driver-settlement-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setIsDriverSettlementOpen(false)}>
        <section className="fleet-modal driver-settlement-modal" role="dialog" aria-modal="true" aria-labelledby="driver-settlement-title">
          <header className="fleet-modal-title">
            <h2 id="driver-settlement-title">تصفية عهدة سواق</h2>
            <button type="button" aria-label="إغلاق تصفية عهدة السائق" onClick={() => setIsDriverSettlementOpen(false)}>×</button>
          </header>
          <div className="driver-settlement-period-toggle" role="group" aria-label="طريقة تحديد فترة التصفية">
            <button type="button" className={driverSettlementPeriodMode === 'month' ? 'is-active' : ''} aria-pressed={driverSettlementPeriodMode === 'month'} onClick={() => {
              setDriverSettlementPeriodMode('month');
              setIsDriverSettlementMonthPickerOpen(false);
            }}>شهر</button>
            <button type="button" className={driverSettlementPeriodMode === 'date' ? 'is-active' : ''} aria-pressed={driverSettlementPeriodMode === 'date'} onClick={() => {
              setDriverSettlementPeriodMode('date');
              setIsDriverSettlementMonthPickerOpen(false);
            }}>تاريخ معين</button>
          </div>
          <div className={`driver-settlement-filters${driverSettlementPeriodMode === 'month' ? ' is-month-mode' : ''}`}>
            <label>السائق
              <select value={driverSettlementDriver} onChange={(event) => setDriverSettlementDriver(event.target.value)} disabled={!driverSettlementDrivers.length}>
                {driverSettlementDrivers.length
                  ? driverSettlementDrivers.map((driver) => <option key={driver} value={driver}>{driver}</option>)
                  : <option value="">لا يوجد سائقون في الحركات</option>}
              </select>
            </label>
            {driverSettlementPeriodMode === 'month' ? <div className="driver-settlement-month-field">
              <span>الشهر</span>
              <MonthYearPicker
                selectedMonth={driverSettlementMonth}
                onSelectMonth={setDriverSettlementMonth}
                ariaLabel="اختيار شهر تصفية عهدة السائق"
                isOpen={isDriverSettlementMonthPickerOpen}
                onOpenChange={setIsDriverSettlementMonthPickerOpen}
                isActive
                className="driver-settlement-month-picker"
                portalStyle={monthPickerPortalStyle}
              />
            </div> : <>
            <label className="driver-settlement-date-field"><span>من تاريخ</span><div className="entry-date-input-wrap driver-settlement-date-wrap">
              <input type="text" inputMode="numeric" dir="ltr" maxLength={10} placeholder="يوم/شهر/سنة" autoComplete="off" value={driverSettlementFromDraft} aria-label="من تاريخ، يوم شهر سنة" onChange={(event) => {
                const value = event.target.value;
                const parsedDate = parseDriverSettlementDate(value);
                setDriverSettlementFromDraft(value);
                setDriverSettlementFrom(parsedDate || '');
              }} onBlur={() => {
                if (parsedDriverSettlementFrom) setDriverSettlementFromDraft(formatEntryDateInput(parsedDriverSettlementFrom));
              }} />
              <button type="button" className="entry-date-picker-button" aria-label="اختيار تاريخ البداية من التقويم" title="اختيار تاريخ البداية من التقويم" onClick={() => {
                const picker = driverSettlementFromPickerRef.current;
                if (!picker) return;
                if (typeof picker.showPicker === 'function') picker.showPicker();
                else {
                  picker.focus();
                  picker.click();
                }
              }}><CalendarDays size={18} aria-hidden="true" /></button>
              <input ref={driverSettlementFromPickerRef} className="entry-date-native-picker" type="date" aria-label="تقويم تاريخ البداية" tabIndex={-1} value={driverSettlementFrom} onChange={(event) => {
                const value = event.target.value;
                setDriverSettlementFrom(value);
                setDriverSettlementFromDraft(formatEntryDateInput(value));
              }} />
            </div></label>
            <label className="driver-settlement-date-field"><span>إلى تاريخ</span><div className="entry-date-input-wrap driver-settlement-date-wrap">
              <input type="text" inputMode="numeric" dir="ltr" maxLength={10} placeholder="يوم/شهر/سنة" autoComplete="off" value={driverSettlementToDraft} aria-label="إلى تاريخ، يوم شهر سنة" onChange={(event) => {
                const value = event.target.value;
                const parsedDate = parseDriverSettlementDate(value);
                setDriverSettlementToDraft(value);
                setDriverSettlementTo(parsedDate || '');
              }} onBlur={() => {
                if (parsedDriverSettlementTo) setDriverSettlementToDraft(formatEntryDateInput(parsedDriverSettlementTo));
              }} />
              <button type="button" className="entry-date-picker-button" aria-label="اختيار تاريخ النهاية من التقويم" title="اختيار تاريخ النهاية من التقويم" onClick={() => {
                const picker = driverSettlementToPickerRef.current;
                if (!picker) return;
                if (typeof picker.showPicker === 'function') picker.showPicker();
                else {
                  picker.focus();
                  picker.click();
                }
              }}><CalendarDays size={18} aria-hidden="true" /></button>
              <input ref={driverSettlementToPickerRef} className="entry-date-native-picker" type="date" aria-label="تقويم تاريخ النهاية" tabIndex={-1} value={driverSettlementTo} onChange={(event) => {
                const value = event.target.value;
                setDriverSettlementTo(value);
                setDriverSettlementToDraft(formatEntryDateInput(value));
              }} />
            </div></label>
            </>}
          </div>
          {driverSettlementDateRangeInvalid
            ? <p className="driver-settlement-empty" role="alert">{driverSettlementDateInputInvalid ? 'أدخل تاريخًا صحيحًا بصيغة يوم/شهر/سنة.' : 'تاريخ البداية يجب أن يسبق تاريخ النهاية.'}</p>
            : driverSettlementEntries.length
              ? <>
                <p className="driver-settlement-entry-count">عدد الحركات في الفترة: {numberFormat.format(driverSettlementEntries.length)}</p>
                <div className="driver-settlement-totals">
                  <article><span>إجمالي المسافة</span><strong>{numberFormat.format(driverSettlementTotals.distance)} <small>كم</small></strong></article>
                  <article><span>التموين الفعلي المدفوع</span><strong>{currencyFormat.format(driverSettlementTotals.actual)}</strong></article>
                  <article><span>التموين المقرر</span><strong>{currencyFormat.format(driverSettlementTotals.allowance)}</strong></article>
                  <article><span>الانحراف الإجمالي</span><strong>{currencyFormat.format(driverSettlementTotals.deviation)}</strong></article>
                </div>
              </>
              : <p className="driver-settlement-empty">{driverSettlementDrivers.length ? 'لا توجد حركات للسائق في الفترة المحددة.' : 'لا توجد حركات مسجلة لعرض التصفية.'}</p>}
          <footer className="driver-settlement-actions">
            <button type="button" className="fleet-button fleet-button-primary driver-settlement-close" onClick={() => setIsDriverSettlementOpen(false)}>إغلاق</button>
          </footer>
        </section>
      </div>}
    </main>
  );
};

export default CarFuelPage;