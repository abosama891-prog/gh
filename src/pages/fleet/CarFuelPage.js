import { jsx as _jsx, jsxs as _jsxs, Fragment as _Fragment } from "react/jsx-runtime";
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { ArrowRight, CarFront, Download, Fuel, LayoutDashboard, RefreshCw, UsersRound } from 'lucide-react';
import { readFleetUsers, saveFleetUsers } from '../../services/fleetUserStore';
import './car-fuel.css';
const STORAGE_KEY = 'company-fleet-fuel-v1';
const DEFAULT_ALLOWANCE_PRICE = 15.75;
const today = () => new Date().toISOString().slice(0, 10);
const vehicleToneStyle = (vehicleId, vehicles) => {
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
    return { '--vehicle-accent': accent, '--vehicle-tint': tint, '--vehicle-line': line };
};
const initialData = {
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
const readData = () => {
    try {
        const stored = localStorage.getItem(STORAGE_KEY);
        return stored ? JSON.parse(stored) : initialData;
    }
    catch {
        return initialData;
    }
};
const numberFormat = new Intl.NumberFormat('en-EG', { maximumFractionDigits: 1 });
const currencyFormat = new Intl.NumberFormat('en-EG', { style: 'currency', currency: 'EGP', maximumFractionDigits: 0 });
const rateFormat = new Intl.NumberFormat('en-EG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const dateFormat = new Intl.DateTimeFormat('ar-EG-u-nu-latn', { year: 'numeric', month: '2-digit', day: '2-digit' });
const formatDateArabic = (value) => dateFormat.format(new Date(`${value}T00:00:00`));
const fuelAllowanceRate = (data) => data.allowancePricePerLiter ?? DEFAULT_ALLOWANCE_PRICE;
const vehicleFuelRate = (vehicle) => vehicle.fuelLitersPer105Km ?? 20;
const vehicleDistanceBase = (vehicle) => vehicle.fuelDistanceBaseKm ?? 105;
const litersForDistance = (vehicle, distance) => Math.max(0, distance) / vehicleDistanceBase(vehicle) * vehicleFuelRate(vehicle);
const litersForEntry = (entry, vehicle) => vehicle ? litersForDistance(vehicle, entry.currentOdometer - entry.startOdometer) : entry.liters;
const expenseFields = [
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
const emptyExpenses = () => ({ mealAllowance: 0, tolls: 0, traffic: 0, violations: 0, maintenance: 0, extra: 0, carWash: 0, tireExpenses: 0, oilChange: 0, transport: 0, other: 0 });
const emptyExpenseInputs = () => Object.fromEntries(expenseFields.map(({ key }) => [key, '']));
const sumExpenses = (expenses) => Math.round(expenseFields.reduce((sum, field) => sum + (Number(expenses?.[field.key]) || 0), 0));
const entryFuelCost = (entry) => Math.round(entry.fuelActualValue ?? entry.liters * entry.pricePerLiter);
const allowanceCost = (liters, rate) => Math.round(Math.max(0, Number(liters) || 0) * rate);
const entryDeviation = (entry, rate, liters = entry.liters) => entryFuelCost(entry) - allowanceCost(liters, rate);
const entryTotalCost = (entry) => Math.round(entryFuelCost(entry) + sumExpenses(entry.expenses));
const parseCsvRows = (content) => {
    const text = content.replace(/^\uFEFF/, '');
    const firstLine = text.split(/\r?\n/, 1)[0] || '';
    let quoted = false;
    let commas = 0;
    let semicolons = 0;
    for (let index = 0; index < firstLine.length; index += 1) {
        if (firstLine[index] === '"') {
            if (quoted && firstLine[index + 1] === '"')
                index += 1;
            else
                quoted = !quoted;
        }
        else if (!quoted && firstLine[index] === ',')
            commas += 1;
        else if (!quoted && firstLine[index] === ';')
            semicolons += 1;
    }
    const delimiter = semicolons > commas ? ';' : ',';
    const rows = [];
    let row = [];
    let field = '';
    quoted = false;
    for (let index = 0; index < text.length; index += 1) {
        const character = text[index];
        if (quoted) {
            if (character === '"' && text[index + 1] === '"') {
                field += '"';
                index += 1;
            }
            else if (character === '"')
                quoted = false;
            else
                field += character;
        }
        else if (character === '"' && field.length === 0)
            quoted = true;
        else if (character === delimiter) {
            row.push(field);
            field = '';
        }
        else if (character === '\n' || character === '\r') {
            row.push(field);
            if (row.some((value) => value.trim() !== ''))
                rows.push(row);
            row = [];
            field = '';
            if (character === '\r' && text[index + 1] === '\n')
                index += 1;
        }
        else
            field += character;
    }
    if (quoted)
        throw new Error('يوجد حقل نصي بعلامة اقتباس غير مكتملة في الملف.');
    row.push(field);
    if (row.some((value) => value.trim() !== ''))
        rows.push(row);
    return rows;
};
const parseTemplateNumber = (value) => {
    const westernDigits = value
        .replace(/[٠-٩]/g, (digit) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)))
        .replace(/[۰-۹]/g, (digit) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit)))
        .replace(/[٬,،\s]/g, '')
        .replace(/٫/g, '.');
    return Number(westernDigits);
};
const parseAmountExpression = (value) => {
    const normalized = value
        .replace(/[٠-٩]/g, (digit) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)))
        .replace(/[۰-۹]/g, (digit) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit)))
        .replace(/٫/g, '.')
        .replace(/[٬,،\s]/g, '');
    if (!normalized)
        return 0;
    const amounts = normalized.split('+');
    if (amounts.some((amount) => !/^(?:\d+(?:\.\d*)?|\.\d+)$/.test(amount)))
        return null;
    const total = amounts.reduce((sum, amount) => sum + Number(amount), 0);
    return Number.isFinite(total) ? Math.round(total) : null;
};
const parseSingleAmount = (value) => value.includes('+') ? null : parseAmountExpression(value);
const splitAmountExpression = (value) => value.split('+').map((amount) => amount.trim()).filter(Boolean);
const excelCellToString = (value) => {
    if (value === null || value === undefined)
        return '';
    if (value instanceof Date)
        return value.toISOString().slice(0, 10);
    if (typeof value === 'object') {
        if ('result' in value)
            return excelCellToString(value.result);
        if ('richText' in value)
            return value.richText.map((part) => part.text).join('');
        if ('text' in value)
            return String(value.text ?? '');
        return '';
    }
    return String(value);
};
const worksheetToRows = (worksheet) => {
    const rows = [];
    worksheet.eachRow({ includeEmpty: true }, (row) => {
        rows.push(Array.from({ length: worksheet.columnCount }, (_, index) => excelCellToString(row.getCell(index + 1).value)));
    });
    return rows;
};
const fleetSections = [
    { key: 'overview', label: 'نظرة عامة', description: 'ملخص الأداء والحركات', Icon: LayoutDashboard },
    { key: 'vehicles', label: 'السيارات', description: 'بيانات ومصروفات المركبات', Icon: CarFront },
    { key: 'entries', label: 'سجل التعبئة', description: 'حركات الوقود المسجلة', Icon: Fuel },
];
const fleetHomeActions = [
    { key: 'users', label: 'إدارة المستخدمين والصلاحيات', description: 'إدارة الحسابات والأدوار', Icon: UsersRound },
    { key: 'update', label: 'تحديث النظام', description: 'التحقق من أحدث نسخة', Icon: RefreshCw },
    { key: 'install', label: 'تثبيت التطبيق', description: 'إضافة التطبيق إلى جهازك', Icon: Download },
];
const fleetRoleLabels = { admin: 'مدير النظام', manager: 'مدير فرع', viewer: 'مشاهد' };
const dateInputValue = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
const subtractCalendarMonths = (date, months) => {
    const shifted = new Date(date);
    const day = shifted.getDate();
    shifted.setDate(1);
    shifted.setMonth(shifted.getMonth() - months);
    shifted.setDate(Math.min(day, new Date(shifted.getFullYear(), shifted.getMonth() + 1, 0).getDate()));
    return shifted;
};
const CarFuelPage = ({ onLogout, currentUser = readFleetUsers()[0] }) => {
    const [data, setData] = useState(readData);
    const [tab, setTab] = useState('home');
    const [users, setUsers] = useState(readFleetUsers);
    const [newUser, setNewUser] = useState({ username: '', password: '', role: 'viewer' });
    const [selectedVehicleId, setSelectedVehicleId] = useState(null);
    const [search, setSearch] = useState('');
    const [ledgerFilters, setLedgerFilters] = useState({ vehicleId: 'all', from: '', to: '', search: '' });
    const [ledgerPeriod, setLedgerPeriod] = useState('all');
    const [vehicleDialog, setVehicleDialog] = useState(false);
    const [editingVehicleId, setEditingVehicleId] = useState(null);
    const vehicleImportRef = useRef(null);
    const [isImportingVehicles, setIsImportingVehicles] = useState(false);
    const vehicleHistoryImportRef = useRef(null);
    const [isImportingVehicleHistory, setIsImportingVehicleHistory] = useState(false);
    const [entryDialog, setEntryDialog] = useState(false);
    const [editingEntryId, setEditingEntryId] = useState(null);
    const [allowanceRateDraft, setAllowanceRateDraft] = useState(String(DEFAULT_ALLOWANCE_PRICE));
    const [installPrompt, setInstallPrompt] = useState(null);
    const [notice, setNotice] = useState('');
    const [pendingDeleteEntry, setPendingDeleteEntry] = useState(null);
    const [deleteConfirmationText, setDeleteConfirmationText] = useState('');
    const [isRateEditorOpen, setIsRateEditorOpen] = useState(false);
    const [vehicleForm, setVehicleForm] = useState({ name: '', plate: '', driver: '', openingOdometer: '', kmPerLiter: '20', distanceBaseKm: '105' });
    const [entryForm, setEntryForm] = useState({ vehicleId: '', driver: '', startOdometer: '', currentOdometer: '', liters: '', pricePerLiter: '2.18', fuelActualValue: '', allowanceValue: '', date: today(), route: '', note: '', expenses: emptyExpenseInputs() });
    const [fuelAmountDraft, setFuelAmountDraft] = useState('');
    const [expenseAmountDrafts, setExpenseAmountDrafts] = useState(emptyExpenseInputs);
    const canManageUsers = currentUser.role === 'admin';
    const canEditFleet = currentUser.role !== 'viewer';
    const persistUsers = (nextUsers) => {
        saveFleetUsers(nextUsers);
        setUsers(nextUsers);
    };
    const addUser = (event) => {
        event.preventDefault();
        const username = newUser.username.trim();
        if (!canManageUsers || !username || !newUser.password.trim()) {
            setNotice('أدخل اسم المستخدم وكلمة المرور');
            return;
        }
        if (users.some((user) => user.username.trim().toLowerCase() === username.toLowerCase())) {
            setNotice('اسم المستخدم موجود بالفعل');
            return;
        }
        persistUsers([...users, { id: `u-${Date.now()}`, username, password: newUser.password, role: newUser.role, active: true, createdAt: new Date().toISOString() }]);
        setNewUser({ username: '', password: '', role: 'viewer' });
        setNotice('تمت إضافة المستخدم');
    };
    const updateUser = (userId, changes) => {
        if (!canManageUsers)
            return;
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
    const deleteUser = (userId) => {
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
    }, [data]);
    useEffect(() => {
        setAllowanceRateDraft(String(fuelAllowanceRate(data)));
    }, [data.allowancePricePerLiter]);
    useEffect(() => {
        const captureInstall = (event) => {
            event.preventDefault();
            setInstallPrompt(event);
        };
        window.addEventListener('beforeinstallprompt', captureInstall);
        if ('serviceWorker' in navigator)
            navigator.serviceWorker.register('/sw.js').catch(() => undefined);
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
    const calculateFuelLiters = (vehicleId, startOdometer, currentOdometer) => {
        const vehicle = data.vehicles.find((item) => item.id === vehicleId);
        if (!vehicle)
            return 0;
        return litersForDistance(vehicle, currentOdometer - startOdometer);
    };
    const calculateAllowanceValue = (vehicleId, liters) => {
        return allowanceCost(liters, fuelAllowanceRate(data));
    };
    const allowanceValueForEntry = (entry) => calculateAllowanceValue(entry.vehicleId, entry.liters);
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
    const selectedEntryDetails = selectedEntries.slice().reverse().reduce((rows, entry) => {
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
    const selectLedgerPeriod = (period) => {
        setLedgerPeriod(period);
        if (period === 'custom')
            return;
        if (period === 'all') {
            setLedgerFilters((current) => ({ ...current, from: '', to: '' }));
            return;
        }
        const now = new Date();
        const endDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        const startDate = new Date(endDate);
        if (period === '30d' || period === '60d' || period === '90d') {
            startDate.setDate(startDate.getDate() - Number(period.slice(0, -1)) + 1);
        }
        else {
            const months = period === '6m' ? 6 : 12;
            startDate.setTime(subtractCalendarMonths(endDate, months).getTime());
        }
        setLedgerFilters((current) => ({ ...current, from: dateInputValue(startDate), to: dateInputValue(endDate) }));
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
    const startEntry = (vehicleId = data.vehicles[0]?.id) => {
        const summary = summaries.find((item) => item.vehicle.id === vehicleId);
        const start = summary?.lastOdometer ?? 0;
        setEditingEntryId(null);
        setFuelAmountDraft('');
        setExpenseAmountDrafts(emptyExpenseInputs());
        setEntryForm({ vehicleId: vehicleId ?? '', driver: summaries.find((item) => item.vehicle.id === vehicleId)?.vehicle.driver || '', startOdometer: String(start), currentOdometer: '', liters: '0.0', pricePerLiter: '0', fuelActualValue: '', allowanceValue: '', date: today(), route: '', note: '', expenses: emptyExpenseInputs() });
        setEntryDialog(true);
    };
    const editEntry = (entry) => {
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
            expenses: Object.fromEntries(expenseFields.map(({ key }) => [key, entry.expenseExpressions?.[key] ?? (entry.expenses?.[key] ? String(entry.expenses[key]) : '')])),
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
    const removeFuelAmount = (indexToRemove) => {
        setEntryForm((current) => ({ ...current, fuelActualValue: splitAmountExpression(current.fuelActualValue).filter((_, index) => index !== indexToRemove).join('+') }));
    };
    const addExpenseAmount = (key) => {
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
    const removeExpenseAmount = (key, indexToRemove) => {
        setEntryForm((current) => ({
            ...current,
            expenses: { ...current.expenses, [key]: splitAmountExpression(current.expenses[key]).filter((_, index) => index !== indexToRemove).join('+') },
        }));
    };
    const saveVehicle = (event) => {
        event.preventDefault();
        const vehicle = {
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
    const openVehicleEditor = (vehicle) => {
        setEditingVehicleId(vehicle.id);
        setVehicleForm({ name: vehicle.name, plate: vehicle.plate, driver: vehicle.driver, openingOdometer: String(vehicle.openingOdometer), kmPerLiter: String(vehicleFuelRate(vehicle)), distanceBaseKm: String(vehicleDistanceBase(vehicle)) });
        setVehicleDialog(true);
    };
    const downloadVehicleTemplate = () => {
        const rows = [
            ['رقم اللوحة', 'اسم السيارة', 'اسم السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)'],
            ...data.vehicles.map((vehicle) => [vehicle.plate, vehicle.name, vehicle.driver, vehicle.openingOdometer, vehicle.kmPerLiter]),
        ];
        const csv = `\uFEFF${rows.map((row) => row.map((value) => `"${String(value ?? '').replaceAll('"', '""')}"`).join(',')).join('\r\n')}`;
        const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
        const link = document.createElement('a');
        link.href = url;
        link.download = 'قالب-سيارات-الشركة.csv';
        link.click();
        window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    };
    const importVehicleTemplate = async (event) => {
        const input = event.currentTarget;
        const file = input.files?.[0];
        if (!file)
            return;
        setIsImportingVehicles(true);
        try {
            const rows = parseCsvRows(await file.text());
            if (rows.length < 2)
                throw new Error('القالب لا يحتوي على صفوف سيارات.');
            const headers = rows[0].map((value) => value.trim());
            const requiredHeaders = ['رقم اللوحة', 'اسم السيارة', 'اسم السائق', 'بداية العداد', 'معدل الاستهلاك (كم/لتر)'];
            const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));
            if (missingHeaders.length)
                throw new Error(`القالب غير مطابق. الأعمدة الناقصة: ${missingHeaders.join('، ')}`);
            const column = (header) => headers.indexOf(header);
            const imported = [];
            const seenPlates = new Set();
            const rowErrors = [];
            rows.slice(1).forEach((row, index) => {
                const plate = (row[column('رقم اللوحة')] || '').trim();
                const name = (row[column('اسم السيارة')] || '').trim();
                const driver = (row[column('اسم السائق')] || '').trim();
                const openingOdometer = parseTemplateNumber(row[column('بداية العداد')] || '');
                const kmPerLiter = parseTemplateNumber(row[column('معدل الاستهلاك (كم/لتر)')] || '');
                if (!plate && !name && !row.some((value) => value.trim()))
                    return;
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
            if (rowErrors.length)
                throw new Error(`راجع بيانات الصفوف التالية (رقم اللوحة والاسم والعداد ومعدل الاستهلاك): ${rowErrors.join('، ')}`);
            if (!imported.length)
                throw new Error('لا توجد سيارات صالحة للاستيراد.');
            const existingPlates = new Set(data.vehicles.map((vehicle) => vehicle.plate.trim().toLocaleLowerCase('ar-EG')));
            const addedCount = imported.filter((vehicle) => !existingPlates.has(vehicle.plate.trim().toLocaleLowerCase('ar-EG'))).length;
            const updatedCount = imported.length - addedCount;
            setData((current) => {
                const vehicles = [...current.vehicles];
                imported.forEach((vehicle, index) => {
                    const matchIndex = vehicles.findIndex((item) => item.plate.trim().toLocaleLowerCase('ar-EG') === vehicle.plate.trim().toLocaleLowerCase('ar-EG'));
                    if (matchIndex >= 0)
                        vehicles[matchIndex] = { ...vehicles[matchIndex], ...vehicle };
                    else
                        vehicles.push({ id: `v${Date.now()}-${index}`, ...vehicle });
                });
                return { ...current, vehicles };
            });
            setNotice(`تم استيراد ${imported.length} سيارة: إضافة ${addedCount} وتحديث ${updatedCount}`);
        }
        catch (error) {
            setNotice(error instanceof Error ? error.message : 'تعذر استيراد الملف.');
        }
        finally {
            setIsImportingVehicles(false);
            input.value = '';
        }
    };
    const downloadSelectedVehicleTemplate = async () => {
        if (!selectedSummary)
            return;
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
        const infoRows = [
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
                }
                else {
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
            const values = [
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
                if (columnNumber === 4)
                    cell.numFmt = 'yyyy-mm-dd';
                if ([11, 12, 13, 14, headers.length].includes(columnNumber))
                    cell.numFmt = '#,##0 "جنيه"';
                if ([7, 8, 9].includes(columnNumber))
                    cell.numFmt = '#,##0';
                if (columnNumber === 10)
                    cell.numFmt = '0.0';
            });
        }
        sheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4, column: headers.length } };
        sheet.getColumn(4).numFmt = 'yyyy-mm-dd';
        const buffer = await workbook.xlsx.writeBuffer();
        const url = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        const link = document.createElement('a');
        link.href = url;
        link.download = `قالب-${selectedSummary.vehicle.plate.replace(/[\\/:*?"<>|\s]+/g, '-')}.xlsx`;
        link.click();
        window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    };
    const importSelectedVehicleTemplate = async (event) => {
        const input = event.currentTarget;
        const file = input.files?.[0];
        if (!file || !selectedSummary)
            return;
        const vehicle = selectedSummary.vehicle;
        setIsImportingVehicleHistory(true);
        try {
            let rows;
            if (/\.xlsx$/i.test(file.name)) {
                const ExcelJS = (await import('exceljs')).default;
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(await file.arrayBuffer());
                const worksheet = workbook.getWorksheet('سجل السيارة') || workbook.worksheets.find((sheet) => sheet.name !== 'بيانات السيارة');
                if (!worksheet)
                    throw new Error('لم أجد ورقة سجل السيارة داخل ملف Excel.');
                rows = worksheetToRows(worksheet);
            }
            else {
                rows = parseCsvRows(await file.text());
            }
            if (rows.length < 2)
                throw new Error('القالب لا يحتوي على حركات.');
            const headerRowIndex = rows.findIndex((row) => row.some((value) => value.trim() === 'معرف الحركة'));
            if (headerRowIndex < 0)
                throw new Error('لم أجد صف عناوين قالب الحركات.');
            const headers = rows[headerRowIndex].map((value) => value.trim());
            const required = [
                'معرف الحركة', 'رقم اللوحة', 'التاريخ', 'اسم السائق', 'خط السير',
                'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)', 'التفصيل',
                ...expenseFields.map((field) => `${field.label} (جنيه)`),
            ];
            const missing = required.filter((header) => !headers.includes(header));
            if (missing.length)
                throw new Error(`القالب غير مطابق. الأعمدة الناقصة: ${missing.join('، ')}`);
            const column = (header) => headers.indexOf(header);
            const oldEntries = data.entries.filter((entry) => entry.vehicleId === vehicle.id);
            const oldIds = new Set(oldEntries.map((entry) => entry.id));
            const seenIds = new Set();
            const imported = [];
            const errors = [];
            rows.slice(headerRowIndex + 1).forEach((row, index) => {
                const value = (header) => (row[column(header)] || '').trim();
                const hasMovement = ['التاريخ', 'القراءة السابقة', 'القراءة الحالية', 'مصروف الوقود الفعلي (جنيه)'].some((header) => value(header) !== '');
                if (!hasMovement)
                    return;
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
                if (suppliedId)
                    seenIds.add(suppliedId);
                const id = suppliedId && oldIds.has(suppliedId) ? suppliedId : `e${Date.now()}-${index}`;
                const liters = calculateFuelLiters(vehicle.id, start, current);
                const expenses = Object.fromEntries(expenseFields.map((field) => [field.key, Math.max(0, Math.round(parseTemplateNumber(value(`${field.label} (جنيه)`)) || 0))]));
                imported.push({
                    id, vehicleId: vehicle.id, date, startOdometer: start, currentOdometer: current,
                    liters, pricePerLiter: liters > 0 ? actualFuelValue / liters : 0, fuelActualValue: actualFuelValue,
                    driver: value('اسم السائق'), route: value('خط السير'), allowanceValue: calculateAllowanceValue(vehicle.id, liters),
                    expenses, note: value('التفصيل'),
                });
            });
            if (errors.length)
                throw new Error(`راجع الصفوف التالية: ${errors.join('، ')}. تأكد من رقم اللوحة والعدادات والمصروف.`);
            if (!imported.length)
                throw new Error('لم يتم العثور على حركات صالحة في القالب.');
            const updatedCount = imported.filter((entry) => oldIds.has(entry.id)).length;
            const addedCount = imported.length - updatedCount;
            setData((current) => {
                const importedById = new Map(imported.map((entry) => [entry.id, entry]));
                const existingIds = new Set(current.entries.map((entry) => entry.id));
                const entries = current.entries.map((entry) => importedById.get(entry.id) || entry);
                imported.forEach((entry) => { if (!existingIds.has(entry.id))
                    entries.push(entry); });
                return { ...current, entries };
            });
            setNotice(`تم استيراد سجل ${vehicle.plate}: تحديث ${updatedCount} وإضافة ${addedCount}`);
        }
        catch (error) {
            setNotice(error instanceof Error ? error.message : 'تعذر استيراد سجل السيارة.');
        }
        finally {
            setIsImportingVehicleHistory(false);
            input.value = '';
        }
    };
    const saveEntry = (event) => {
        event.preventDefault();
        if (fuelAmountDraft.trim() || Object.values(expenseAmountDrafts).some((amount) => amount.trim())) {
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
        const parsedExpenses = Object.fromEntries(expenseFields.map(({ key }) => [key, parseAmountExpression(entryForm.expenses[key])]));
        if (fuelActualValue === null || Object.values(parsedExpenses).some((amount) => amount === null)) {
            setNotice('راجع المبالغ المضافة وتأكد من صحتها');
            return;
        }
        const calculatedLiters = calculateFuelLiters(entryForm.vehicleId, start, current);
        const entry = {
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
            expenses: Object.fromEntries(expenseFields.map(({ key }) => [key, parsedExpenses[key] ?? 0])),
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
        if (!pendingDeleteEntry || deleteConfirmationText.trim() !== 'موافق')
            return;
        setData((current) => ({ ...current, entries: current.entries.filter((entry) => entry.id !== pendingDeleteEntry.id) }));
        setPendingDeleteEntry(null);
        setDeleteConfirmationText('');
        setNotice('تم حذف سجل التعبئة');
    };
    const openDeleteConfirmation = (entry) => {
        setPendingDeleteEntry(entry);
        setDeleteConfirmationText('');
    };
    const cancelDeleteConfirmation = () => {
        setPendingDeleteEntry(null);
        setDeleteConfirmationText('');
    };
    const exportCsv = () => {
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
        const csv = `\uFEFF${rows.map((row) => row.map((value) => `"${String(value).replaceAll('"', '""')}"`).join(',')).join('\r\n')}`;
        const link = document.createElement('a');
        link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
        link.download = 'سجل-مصروفات-السيارات.csv';
        link.click();
        window.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    };
    const installApp = async () => {
        if (!installPrompt) {
            setNotice('من قائمة المتصفح اختر «إضافة إلى الشاشة الرئيسية» لتثبيت التطبيق');
            return;
        }
        await installPrompt.prompt();
        await installPrompt.userChoice;
        setInstallPrompt(null);
    };
    const updateSystem = async () => {
        try {
            if ('serviceWorker' in navigator) {
                const registration = await navigator.serviceWorker.getRegistration();
                await registration?.update();
            }
            window.location.reload();
        }
        catch {
            setNotice('تعذر التحقق من تحديث النظام. تحقق من اتصال الإنترنت ثم أعد المحاولة');
        }
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
    const expenseAmountPreviews = Object.fromEntries(expenseFields.map(({ key }) => [key, parseAmountExpression(entryForm.expenses[key])]));
    const hasInvalidAmount = fuelActualValuePreview === null || Object.values(expenseAmountPreviews).some((amount) => amount === null);
    const hasUnaddedAmounts = Boolean(fuelAmountDraft.trim() || Object.values(expenseAmountDrafts).some((amount) => amount.trim()));
    const entryCostPreview = Math.round((fuelActualValuePreview ?? 0) + Object.values(expenseAmountPreviews).reduce((sum, amount) => sum + (amount ?? 0), 0));
    const navigationTiles = [
        ...fleetSections,
        ...(canManageUsers ? [fleetHomeActions[0]] : []),
        ...(tab === 'home' ? [fleetHomeActions[1], fleetHomeActions[2]] : []),
    ];
    const openSection = fleetSections.find((section) => section.key === tab)
        || (tab === 'users' ? fleetHomeActions[0] : undefined);
    return (_jsxs("main", { className: `fleet-app${selectedSummary ? ' fleet-app-detail' : ''}${tab === 'home' ? ' fleet-app-home' : ''}`, dir: "rtl", children: [!selectedSummary && _jsxs(_Fragment, { children: [_jsxs("header", { className: "fleet-header", children: [_jsxs("a", { className: "fleet-brand", href: "/cars", "aria-label": "\u0645\u0631\u0627\u0642\u0628\u0629 \u0633\u064A\u0627\u0631\u0627\u062A \u0634\u0631\u0643\u0629 \u0623\u0628\u0648 \u063A\u0627\u0644\u064A", children: [_jsx("span", { className: "fleet-brand-mark", children: _jsx("img", { src: "/car-mark.svg", alt: "" }) }), _jsx("span", { className: "fleet-brand-copy", children: _jsxs("strong", { children: ["\u0645\u0631\u0627\u0642\u0628\u0629 ", _jsx("em", { children: "\u0633\u064A\u0627\u0631\u0627\u062A \u0623\u0628\u0648 \u063A\u0627\u0644\u064A" })] }) })] }), _jsxs("div", { className: "fleet-header-actions", children: [_jsxs("span", { className: "fleet-live", children: [_jsx("i", {}), " \u0645\u062A\u0627\u0628\u0639\u0629 \u0645\u0628\u0627\u0634\u0631\u0629"] }), _jsx("button", { className: "fleet-install", type: "button", onClick: installApp, children: "\u062A\u062B\u0628\u064A\u062A \u0627\u0644\u062A\u0637\u0628\u064A\u0642" }), onLogout && _jsx("button", { className: "fleet-logout", type: "button", onClick: onLogout, children: "\u062A\u0633\u062C\u064A\u0644 \u0627\u0644\u062E\u0631\u0648\u062C" })] })] }), tab === 'home' ? _jsx("nav", { className: "fleet-tabs fleet-tabs-home", "aria-label": "\u0623\u0642\u0633\u0627\u0645 \u0645\u062A\u0627\u0628\u0639\u0629 \u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A", children: navigationTiles.map(({ key, label, description, Icon }) => (_jsxs("button", { type: "button", className: tab === key ? 'is-active' : '', onClick: () => {
                                if (key === 'install') {
                                    void installApp();
                                    return;
                                }
                                if (key === 'update') {
                                    void updateSystem();
                                    return;
                                }
                                setTab(key);
                                if (key !== 'vehicles')
                                    setSelectedVehicleId(null);
                            }, children: [_jsx("span", { className: "fleet-tab-icon", children: _jsx(Icon, { "aria-hidden": "true", strokeWidth: 1.8 }) }), _jsxs("span", { className: "fleet-tab-copy", children: [_jsx("strong", { children: label }), _jsx("small", { children: description })] })] }, key))) }) : openSection ? _jsx("nav", { className: "fleet-section-bar", "aria-label": "\u0627\u0644\u0642\u0633\u0645 \u0627\u0644\u0645\u0641\u062A\u0648\u062D", children: _jsxs("div", { className: "fleet-section-bar-center", dir: "rtl", children: [_jsx("button", { type: "button", className: "fleet-section-title-button", "aria-label": `العودة للرئيسية من ${openSection.label}`, onClick: () => { setTab('home'); setSelectedVehicleId(null); }, children: openSection.label }), _jsx("button", { type: "button", className: "fleet-section-back-button", "aria-label": "\u0631\u062C\u0648\u0639 \u0625\u0644\u0649 \u0627\u0644\u0635\u0641\u062D\u0629 \u0627\u0644\u0631\u0626\u064A\u0633\u064A\u0629", onClick: () => { setTab('home'); setSelectedVehicleId(null); }, children: _jsx(ArrowRight, { "aria-hidden": "true" }) })] }) }) : null] }), notice && _jsxs("button", { className: "fleet-notice", type: "button", onClick: () => setNotice(''), children: [notice, _jsx("span", { children: "\u00D7" })] }), tab === 'users' && canManageUsers && _jsxs("section", { className: "fleet-section fleet-user-management", children: [_jsxs("div", { className: "fleet-section-heading", children: [_jsxs("div", { children: [_jsx("span", { className: "fleet-overline", children: "\u0627\u0644\u062D\u0633\u0627\u0628\u0627\u062A \u0639\u0644\u0649 \u0647\u0630\u0627 \u0627\u0644\u062C\u0647\u0627\u0632" }), _jsx("h2", { children: "\u0625\u062F\u0627\u0631\u0629 \u0627\u0644\u0645\u0633\u062A\u062E\u062F\u0645\u064A\u0646 \u0648\u0627\u0644\u0635\u0644\u0627\u062D\u064A\u0627\u062A" })] }), _jsx("button", { type: "button", className: "fleet-link-button", onClick: () => setTab('home'), children: "\u0627\u0644\u0639\u0648\u062F\u0629 \u0644\u0644\u0631\u0626\u064A\u0633\u064A\u0629" })] }), _jsxs("form", { className: "fleet-user-form", onSubmit: addUser, children: [_jsxs("label", { children: ["\u0627\u0633\u0645 \u0627\u0644\u0645\u0633\u062A\u062E\u062F\u0645", _jsx("input", { required: true, value: newUser.username, onChange: (event) => setNewUser((current) => ({ ...current, username: event.target.value })), autoComplete: "off" })] }), _jsxs("label", { children: ["\u0643\u0644\u0645\u0629 \u0627\u0644\u0645\u0631\u0648\u0631", _jsx("input", { required: true, type: "password", value: newUser.password, onChange: (event) => setNewUser((current) => ({ ...current, password: event.target.value })), autoComplete: "new-password" })] }), _jsxs("label", { children: ["\u0627\u0644\u0635\u0644\u0627\u062D\u064A\u0629", _jsxs("select", { value: newUser.role, onChange: (event) => setNewUser((current) => ({ ...current, role: event.target.value })), children: [_jsx("option", { value: "admin", children: "\u0645\u062F\u064A\u0631 \u0627\u0644\u0646\u0638\u0627\u0645" }), _jsx("option", { value: "manager", children: "\u0645\u062F\u064A\u0631 \u0641\u0631\u0639" }), _jsx("option", { value: "viewer", children: "\u0645\u0634\u0627\u0647\u062F" })] })] }), _jsx("button", { className: "fleet-button fleet-button-primary", type: "submit", children: "\u0625\u0636\u0627\u0641\u0629 \u0645\u0633\u062A\u062E\u062F\u0645" })] }), _jsx("div", { className: "fleet-user-table-wrap", children: _jsxs("table", { className: "fleet-table fleet-user-table", children: [_jsx("thead", { children: _jsxs("tr", { children: [_jsx("th", { children: "\u0627\u0633\u0645 \u0627\u0644\u0645\u0633\u062A\u062E\u062F\u0645" }), _jsx("th", { children: "\u0627\u0644\u0635\u0644\u0627\u062D\u064A\u0629" }), _jsx("th", { children: "\u0627\u0644\u062D\u0627\u0644\u0629" }), _jsx("th", { children: "\u0625\u062C\u0631\u0627\u0621" })] }) }), _jsx("tbody", { children: users.map((user) => _jsxs("tr", { children: [_jsxs("td", { children: [_jsx("strong", { children: user.username }), user.id === currentUser.id && _jsx("small", { children: "\u0627\u0644\u062D\u0633\u0627\u0628 \u0627\u0644\u062D\u0627\u0644\u064A" })] }), _jsx("td", { children: _jsxs("select", { "aria-label": `صلاحية ${user.username}`, value: user.role, disabled: user.id === currentUser.id, onChange: (event) => updateUser(user.id, { role: event.target.value }), children: [_jsx("option", { value: "admin", children: "\u0645\u062F\u064A\u0631 \u0627\u0644\u0646\u0638\u0627\u0645" }), _jsx("option", { value: "manager", children: "\u0645\u062F\u064A\u0631 \u0641\u0631\u0639" }), _jsx("option", { value: "viewer", children: "\u0645\u0634\u0627\u0647\u062F" })] }) }), _jsx("td", { children: _jsx("button", { type: "button", className: `fleet-user-status${user.active ? ' is-active' : ''}`, disabled: user.id === currentUser.id, onClick: () => updateUser(user.id, { active: !user.active }), children: user.active ? 'نشط' : 'معطل' }) }), _jsx("td", { children: _jsx("button", { type: "button", className: "entry-delete-button", disabled: user.id === currentUser.id, onClick: () => deleteUser(user.id), children: "\u062D\u0630\u0641" }) })] }, user.id)) })] }) }), _jsx("p", { className: "fleet-user-storage-note", children: "\u062A\u064F\u062D\u0641\u0638 \u0627\u0644\u062D\u0633\u0627\u0628\u0627\u062A \u0639\u0644\u0649 \u0647\u0630\u0627 \u0627\u0644\u062C\u0647\u0627\u0632 \u0641\u0642\u0637\u060C \u0648\u0643\u0644\u0645\u0627\u062A \u0627\u0644\u0645\u0631\u0648\u0631 \u063A\u064A\u0631 \u0645\u0634\u0641\u0651\u0631\u0629." })] }), tab === 'overview' && _jsxs(_Fragment, { children: [_jsxs("section", { className: "ledger-filters", "aria-label": "\u0641\u0644\u062A\u0631\u0629 \u0643\u0634\u0641 \u0627\u0644\u0648\u0642\u0648\u062F", children: [_jsx("div", { className: "ledger-filters-heading", children: _jsx("h2", { children: "\u0641\u0644\u062A\u0631\u0629 \u0633\u062C\u0644 \u0627\u0644\u0648\u0642\u0648\u062F" }) }), _jsx("div", { className: "ledger-period-controls", role: "group", "aria-label": "\u0627\u0644\u0641\u062A\u0631\u0629 \u0627\u0644\u0632\u0645\u0646\u064A\u0629 \u0644\u0633\u062C\u0644 \u0627\u0644\u0648\u0642\u0648\u062F", children: [['all', 'كل الفترة'], ['30d', '30 يوم'], ['60d', '60 يوم'], ['90d', '90 يوم (3 أشهر)'], ['6m', '6 أشهر'], ['1y', 'سنة كاملة'], ['custom', 'تاريخ محدد']].map(([period, label]) => _jsx("button", { type: "button", className: ledgerPeriod === period ? 'is-active' : '', "aria-pressed": ledgerPeriod === period, onClick: () => selectLedgerPeriod(period), children: label }, period)) }), _jsxs("div", { className: `ledger-filter-grid${ledgerPeriod === 'custom' ? ' ledger-filter-grid-custom' : ''}`, children: [_jsxs("label", { children: [_jsx("span", { children: "\u0627\u0644\u0633\u064A\u0627\u0631\u0629" }), _jsxs("select", { value: ledgerFilters.vehicleId, onChange: (event) => setLedgerFilters((current) => ({ ...current, vehicleId: event.target.value })), children: [_jsx("option", { value: "all", children: "\u0643\u0644 \u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A" }), data.vehicles.map((vehicle) => _jsxs("option", { value: vehicle.id, children: [vehicle.plate, " \u00B7 ", vehicle.name] }, vehicle.id))] })] }), ledgerPeriod === 'custom' && _jsxs(_Fragment, { children: [_jsxs("label", { children: [_jsx("span", { children: "\u0645\u0646 \u062A\u0627\u0631\u064A\u062E" }), _jsx("input", { type: "date", value: ledgerFilters.from, onChange: (event) => { setLedgerPeriod('custom'); setLedgerFilters((current) => ({ ...current, from: event.target.value })); } })] }), _jsxs("label", { children: [_jsx("span", { children: "\u0625\u0644\u0649 \u062A\u0627\u0631\u064A\u062E" }), _jsx("input", { type: "date", value: ledgerFilters.to, onChange: (event) => { setLedgerPeriod('custom'); setLedgerFilters((current) => ({ ...current, to: event.target.value })); } })] })] }), _jsxs("label", { className: "ledger-filter-search", children: [_jsx("span", { children: "\u0628\u062D\u062B \u0641\u064A \u0627\u0644\u0633\u062C\u0644" }), _jsx("input", { type: "search", value: ledgerFilters.search, onChange: (event) => setLedgerFilters((current) => ({ ...current, search: event.target.value })), placeholder: "\u0627\u0644\u0633\u064A\u0627\u0631\u0629 \u0623\u0648 \u0627\u0644\u0644\u0648\u062D\u0629 \u0623\u0648 \u0627\u0644\u0633\u0627\u0626\u0642" })] }), _jsxs("button", { className: "ledger-filter-reset", type: "button", onClick: () => { setLedgerPeriod('all'); setLedgerFilters({ vehicleId: 'all', from: '', to: '', search: '' }); }, children: [_jsx("span", { "aria-hidden": "true", children: "\u21B6" }), " \u0625\u0639\u0627\u062F\u0629 \u062A\u0639\u064A\u064A\u0646"] })] })] }), _jsxs("section", { className: "ledger-summary-block", "aria-label": "\u0645\u0644\u062E\u0635 \u0623\u062F\u0627\u0621 \u0627\u0644\u0623\u0633\u0637\u0648\u0644", children: [_jsxs("div", { className: "ledger-summary-heading", children: [_jsxs("div", { children: [_jsx("span", { className: "fleet-overline", children: "\u0645\u0644\u062E\u0635 \u0627\u0644\u0623\u062F\u0627\u0621" }), _jsx("h2", { children: "\u0635\u0648\u0631\u0629 \u0633\u0631\u064A\u0639\u0629 \u0639\u0646 \u0627\u0644\u0623\u0633\u0637\u0648\u0644" })] }), _jsxs("span", { children: [ledgerEntries.length, " \u062D\u0631\u0643\u0629 \u0645\u0633\u062C\u0644\u0629"] })] }), _jsxs("div", { className: "ledger-summary", children: [_jsxs("div", { children: [_jsx("span", { children: "\u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A \u0641\u064A \u0627\u0644\u0643\u0634\u0641" }), _jsx("strong", { children: numberFormat.format(new Set(ledgerEntries.map((entry) => entry.vehicleId)).size) }), _jsx("small", { children: "\u0645\u0631\u0643\u0628\u0629 \u0636\u0645\u0646 \u0627\u0644\u0641\u0644\u062A\u0631" })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0627\u0644\u0645\u0633\u0627\u0641\u0629 \u0627\u0644\u0643\u0644\u064A\u0629" }), _jsxs("strong", { children: [numberFormat.format(ledgerTotals.distance), " ", _jsx("small", { children: "\u0643\u0645" })] }), _jsx("small", { children: "\u0636\u0645\u0646 \u0627\u0644\u0641\u062A\u0631\u0629 \u0627\u0644\u0645\u062D\u062F\u062F\u0629" })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0645\u0635\u0631\u0648\u0641" }), _jsxs("strong", { children: [numberFormat.format(ledgerTotals.liters), " ", _jsx("small", { children: "\u0644\u062A\u0631" })] }), _jsx("small", { children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u062A\u0639\u0628\u0626\u0629" })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u0645\u0635\u0631\u0648\u0641" }), _jsx("strong", { children: currencyFormat.format(ledgerTotals.cost) }), _jsx("small", { children: "\u0648\u0642\u0648\u062F \u0648\u0645\u0635\u0627\u0631\u064A\u0641 \u0625\u0636\u0627\u0641\u064A\u0629" })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0627\u0646\u062D\u0631\u0627\u0641 \u0627\u0644\u062A\u0645\u0648\u064A\u0646" }), _jsx("strong", { className: Math.abs(ledgerTotals.deviation) > 200 ? 'ledger-warning' : 'ledger-ok', children: currencyFormat.format(ledgerTotals.deviation) }), _jsx("small", { children: "\u0645\u0642\u0627\u0631\u0646\u0629 \u0628\u0627\u0644\u0633\u0639\u0631 \u0627\u0644\u0645\u0639\u062A\u0645\u062F" })] })] })] }), _jsxs("section", { className: "ledger-sheet", children: [_jsxs("div", { className: "ledger-titlebar", children: [_jsxs("div", { children: [_jsx("span", { children: "\u0633\u062C\u0644 \u0627\u0644\u062A\u0634\u063A\u064A\u0644 \u0648\u0627\u0644\u0648\u0642\u0648\u062F" }), _jsx("strong", { children: "\u0643\u0634\u0641 \u0645\u062A\u0627\u0628\u0639\u0629 \u0645\u0635\u0631\u0648\u0641\u0627\u062A \u0633\u064A\u0627\u0631\u0627\u062A \u0627\u0644\u0634\u0631\u0643\u0629" })] }), _jsxs("span", { className: "ledger-period", children: [ledgerEntries.length, " \u062D\u0631\u0643\u0629 \u0645\u0633\u062C\u0644\u0629"] })] }), _jsx("div", { className: "fleet-table-wrap", children: _jsxs("table", { className: "fleet-table ledger-table", children: [_jsx("thead", { children: _jsxs("tr", { children: [_jsx("th", { children: "\u0645" }), _jsx("th", { children: "\u0631\u0642\u0645 \u0627\u0644\u0633\u064A\u0627\u0631\u0629" }), _jsx("th", { children: "\u0627\u0633\u0645 \u0627\u0644\u0645\u0631\u0643\u0628\u0629" }), _jsx("th", { children: "\u0627\u0644\u0633\u0627\u0626\u0642" }), _jsx("th", { children: "\u062A\u0627\u0631\u064A\u062E \u0627\u0644\u062A\u0639\u0628\u0626\u0629" }), _jsx("th", { children: "\u0628\u062F\u0627\u064A\u0629 \u0627\u0644\u0639\u062F\u0627\u062F" }), _jsx("th", { children: "\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u0639\u062F\u0627\u062F" }), _jsx("th", { children: "\u0627\u0644\u0641\u062C\u0648\u0629 (\u0643\u0645)" }), _jsx("th", { children: "\u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0645\u062A\u0648\u0642\u0639" }), _jsx("th", { children: "\u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0645\u0635\u0631\u0648\u0641" }), _jsx("th", { children: "\u0627\u0646\u062D\u0631\u0627\u0641 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 (\u062C\u0646\u064A\u0647)" }), _jsx("th", { children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u0645\u0635\u0631\u0648\u0641" })] }) }), _jsx("tbody", { children: ledgerEntries.map((entry, index) => {
                                                const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId);
                                                const distance = Math.max(0, entry.currentOdometer - entry.startOdometer);
                                                const expected = vehicle ? litersForDistance(vehicle, distance) : 0;
                                                const liters = litersForEntry(entry, vehicle);
                                                const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters);
                                                return _jsxs("tr", { className: "fleet-vehicle-tone", style: vehicleToneStyle(entry.vehicleId, data.vehicles), children: [_jsx("td", { className: "ledger-index", children: numberFormat.format(index + 1) }), _jsx("td", { className: "ledger-plate", children: vehicle ? _jsx("button", { type: "button", className: "ledger-vehicle-link", onClick: () => { setSelectedVehicleId(vehicle.id); setTab('vehicles'); }, children: vehicle.plate }) : '—' }), _jsx("td", { children: vehicle?.name || 'سيارة محذوفة' }), _jsx("td", { children: entry.driver || vehicle?.driver || '—' }), _jsx("td", { children: formatDateArabic(entry.date) }), _jsx("td", { children: numberFormat.format(entry.startOdometer) }), _jsx("td", { children: numberFormat.format(entry.currentOdometer) }), _jsxs("td", { children: [numberFormat.format(distance), " \u0643\u0645"] }), _jsxs("td", { children: [numberFormat.format(expected), " \u0644\u062A\u0631"] }), _jsxs("td", { children: [numberFormat.format(liters), " \u0644\u062A\u0631"] }), _jsx("td", { className: Math.abs(deviation) > 200 ? 'ledger-warning' : 'ledger-ok', children: currencyFormat.format(deviation) }), _jsx("td", { children: currencyFormat.format(entryTotalCost(entry)) })] }, entry.id);
                                            }) }), _jsx("tfoot", { children: _jsxs("tr", { children: [_jsx("th", { colSpan: 7, children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u0643\u0634\u0641" }), _jsxs("th", { children: [numberFormat.format(ledgerTotals.distance), " \u0643\u0645"] }), _jsxs("th", { children: [numberFormat.format(ledgerTotals.expected), " \u0644\u062A\u0631"] }), _jsxs("th", { children: [numberFormat.format(ledgerTotals.liters), " \u0644\u062A\u0631"] }), _jsx("th", { className: Math.abs(ledgerTotals.deviation) > 200 ? 'ledger-warning' : 'ledger-ok', children: currencyFormat.format(ledgerTotals.deviation) }), _jsx("th", { children: currencyFormat.format(ledgerTotals.cost) })] }) })] }) }), _jsx("div", { className: "ledger-note", children: "\u0627\u0644\u0641\u062C\u0648\u0629 = \u0627\u0644\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u062D\u0627\u0644\u064A\u0629 \u2212 \u0627\u0644\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u0633\u0627\u0628\u0642\u0629. \u0627\u0646\u062D\u0631\u0627\u0641 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 = \u0645\u0635\u0631\u0648\u0641 \u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0641\u0639\u0644\u064A \u2212 \u0642\u064A\u0645\u0629 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0627\u0644\u0645\u0639\u062A\u0645\u062F\u0629." })] })] }), tab === 'vehicles' && _jsxs("section", { className: `fleet-section${selectedSummary ? ' fleet-detail-page' : ''}`, children: [_jsxs("section", { className: "fleet-global-rate fleet-rate-compact", "aria-label": "\u0633\u0639\u0631 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0627\u0644\u0645\u0648\u062D\u062F", children: [_jsxs("div", { children: [_jsx("strong", { children: "\u0633\u0639\u0631 \u0627\u0644\u0644\u062A\u0631 \u0627\u0644\u0645\u0639\u062A\u0645\u062F" }), _jsx("span", { children: "\u064A\u064F\u0637\u0628\u0642 \u062A\u0644\u0642\u0627\u0626\u064A\u064B\u0627 \u0639\u0644\u0649 \u062D\u0633\u0627\u0628\u0627\u062A \u062C\u0645\u064A\u0639 \u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A" })] }), !isRateEditorOpen ? _jsxs("button", { type: "button", className: "fleet-rate-value", onClick: () => setIsRateEditorOpen(true), children: [rateFormat.format(fuelAllowanceRate(data)), " ", _jsx("small", { children: "\u062C\u0646\u064A\u0647 / \u0644\u062A\u0631" }), _jsx("b", { children: "\u062A\u063A\u064A\u064A\u0631" })] }) : _jsxs("div", { className: "fleet-rate-editor", children: [_jsxs("label", { children: [_jsx("input", { "aria-label": "\u0633\u0639\u0631 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0627\u0644\u0645\u0639\u062A\u0645\u062F \u0628\u0627\u0644\u062C\u0646\u064A\u0647 \u0644\u0643\u0644 \u0644\u062A\u0631", type: "number", min: "0.01", step: "0.01", value: allowanceRateDraft, onChange: (event) => setAllowanceRateDraft(event.target.value) }), _jsx("span", { children: "\u062C\u0646\u064A\u0647 / \u0644\u062A\u0631" })] }), _jsx("button", { type: "button", onClick: saveAllowanceRate, children: "\u062D\u0641\u0638" }), _jsx("button", { type: "button", className: "fleet-rate-cancel", onClick: () => setIsRateEditorOpen(false), children: "\u0625\u0644\u063A\u0627\u0621" })] })] }), selectedSummary ? _jsxs(_Fragment, { children: [_jsxs("div", { className: "fleet-section-heading fleet-detail-heading", children: [_jsxs("div", { children: [_jsx("button", { type: "button", className: "fleet-link-button", onClick: () => setSelectedVehicleId(null), children: "\u2190 \u0627\u0644\u0639\u0648\u062F\u0629 \u0625\u0644\u0649 \u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A" }), _jsx("h2", { children: selectedSummary.vehicle.name }), _jsxs("span", { className: "fleet-overline", children: [selectedSummary.vehicle.plate, " \u00B7 ", selectedSummary.vehicle.driver || 'السائق غير محدد'] })] }), _jsx("div", { className: "fleet-detail-actions", children: _jsx("button", { className: "fleet-button fleet-button-primary", type: "button", onClick: () => startEntry(selectedSummary.vehicle.id), children: "+ \u062A\u0633\u062C\u064A\u0644 \u062A\u0639\u0628\u0626\u0629" }) })] }), _jsxs("div", { className: "fleet-detail-stats", children: [_jsxs("div", { children: [_jsx("span", { children: "\u0628\u062F\u0627\u064A\u0629 \u0627\u0644\u0639\u062F\u0627\u062F" }), _jsxs("strong", { children: [numberFormat.format(selectedSummary.vehicle.openingOdometer), " \u0643\u0645"] })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0627\u0644\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u062D\u0627\u0644\u064A\u0629" }), _jsxs("strong", { children: [numberFormat.format(selectedSummary.lastOdometer), " \u0643\u0645"] })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0627\u0644\u0645\u0633\u0627\u0641\u0629 \u0627\u0644\u0645\u0642\u0637\u0648\u0639\u0629" }), _jsxs("strong", { children: [numberFormat.format(selectedSummary.distance), " \u0643\u0645"] })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0645\u062A\u0648\u0642\u0639" }), _jsxs("strong", { children: [numberFormat.format(selectedSummary.expected), " \u0644\u062A\u0631"] })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0639\u062F\u062F \u0627\u0644\u0644\u062A\u0631\u0627\u062A \u0627\u0644\u0645\u0635\u0631\u0648\u0641\u0629" }), _jsxs("strong", { children: [numberFormat.format(selectedSummary.liters), " \u0644\u062A\u0631"] })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0645\u0639\u062F\u0644 \u0627\u0644\u0644\u062A\u0631\u0627\u062A" }), _jsxs("strong", { children: [numberFormat.format(vehicleFuelRate(selectedSummary.vehicle)), " \u0644\u062A\u0631 / ", numberFormat.format(vehicleDistanceBase(selectedSummary.vehicle)), " \u0643\u0645"] })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0645\u0635\u0627\u0631\u064A\u0641 \u0627\u0644\u0633\u064A\u0627\u0631\u0629" }), _jsx("strong", { children: currencyFormat.format(selectedSummary.cost) })] }), _jsxs("div", { children: [_jsx("span", { children: "\u0633\u0639\u0631 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0627\u0644\u0645\u0648\u062D\u062F" }), _jsxs("strong", { children: [rateFormat.format(fuelAllowanceRate(data)), " \u062C\u0646\u064A\u0647 / \u0644\u062A\u0631"] })] })] }), _jsxs("div", { className: "fleet-table-wrap fleet-detail-table-wrap", children: [_jsxs("table", { className: "fleet-table ledger-table fleet-detail-ledger", children: [_jsx("thead", { children: _jsxs("tr", { children: [_jsx("th", { children: "\u0627\u0644\u062A\u0627\u0631\u064A\u062E" }), _jsx("th", { children: "\u0627\u0633\u0645 \u0627\u0644\u0633\u0627\u0626\u0642" }), _jsx("th", { children: "\u062E\u0637 \u0627\u0644\u0633\u064A\u0631" }), _jsx("th", { children: "\u0627\u0644\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u0633\u0627\u0628\u0642\u0629" }), _jsx("th", { children: "\u0627\u0644\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u062D\u0627\u0644\u064A\u0629" }), _jsx("th", { children: "\u0627\u0644\u0641\u062C\u0648\u0629 (\u0643\u0645)" }), _jsx("th", { children: "\u0639\u062F\u062F \u0627\u0644\u0644\u062A\u0631\u0627\u062A" }), _jsx("th", { children: "\u0642\u064A\u0645\u0629 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0627\u0644\u0645\u0642\u0631\u0631\u0629" }), _jsx("th", { children: "\u0642\u064A\u0645\u0629 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0627\u0644\u0641\u0639\u0644\u064A\u0629" }), _jsx("th", { children: "\u0627\u0644\u0627\u0646\u062D\u0631\u0627\u0641" }), _jsx("th", { children: "\u0627\u0644\u062A\u0631\u0627\u0643\u0645" }), expenseFields.map((field) => _jsx("th", { children: field.label }, field.key)), _jsx("th", { children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0645\u0635\u0627\u0631\u064A\u0641 \u0627\u0644\u0633\u064A\u0627\u0631\u0629" }), _jsx("th", { children: "\u0625\u062C\u0631\u0627\u0621" })] }) }), _jsx("tbody", { children: selectedEntryDetails.map(({ entry, distance, actualFuelValue, deviation, accumulated, totalCost }) => {
                                                    const liters = litersForEntry(entry, selectedSummary.vehicle);
                                                    return _jsxs("tr", { className: "fleet-vehicle-tone", style: vehicleToneStyle(entry.vehicleId, data.vehicles), children: [_jsx("td", { children: formatDateArabic(entry.date) }), _jsx("td", { children: entry.driver || selectedSummary.vehicle.driver || '—' }), _jsx("td", { children: entry.route || '—' }), _jsx("td", { children: numberFormat.format(entry.startOdometer) }), _jsx("td", { children: numberFormat.format(entry.currentOdometer) }), _jsxs("td", { children: [numberFormat.format(distance), " \u0643\u0645"] }), _jsxs("td", { children: [numberFormat.format(liters), " \u0644\u062A\u0631"] }), _jsx("td", { children: currencyFormat.format(allowanceCost(liters, fuelAllowanceRate(data))) }), _jsx("td", { children: currencyFormat.format(actualFuelValue) }), _jsx("td", { className: Math.abs(deviation) > 200 ? 'ledger-warning' : 'ledger-ok', children: currencyFormat.format(deviation) }), _jsx("td", { className: accumulated > 200 ? 'ledger-warning' : 'ledger-ok', children: currencyFormat.format(accumulated) }), expenseFields.map((field) => _jsx("td", { children: currencyFormat.format(entry.expenses?.[field.key] || 0) }, field.key)), _jsxs("td", { children: [currencyFormat.format(totalCost), entry.note ? _jsx("small", { className: "ledger-row-note", children: entry.note }) : null] }), _jsxs("td", { className: "ledger-actions", children: [_jsx("button", { type: "button", className: "entry-edit-button", onClick: () => editEntry(entry), children: "\u062A\u0639\u062F\u064A\u0644" }), _jsx("button", { type: "button", className: "entry-delete-button", onClick: () => openDeleteConfirmation(entry), children: "\u062D\u0630\u0641" })] })] }, entry.id);
                                                }) }), selectedEntries.length > 0 && _jsx("tfoot", { children: _jsxs("tr", { children: [_jsx("th", { colSpan: 5, children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u0633\u064A\u0627\u0631\u0629" }), _jsxs("th", { children: [numberFormat.format(selectedSummary.distance), " \u0643\u0645"] }), _jsxs("th", { children: [numberFormat.format(selectedSummary.liters), " \u0644\u062A\u0631"] }), _jsx("th", { colSpan: 3, children: currencyFormat.format(selectedEntryDetails.reduce((sum, row) => sum + row.deviation, 0)) }), _jsx("th", { children: "\u2014" }), expenseFields.map((field) => _jsx("th", { children: currencyFormat.format(selectedEntries.reduce((sum, entry) => sum + (entry.expenses?.[field.key] || 0), 0)) }, field.key)), _jsx("th", { children: currencyFormat.format(selectedSummary.cost) }), _jsx("th", {})] }) })] }), selectedEntries.length === 0 && _jsx("p", { className: "fleet-empty-state", children: "\u0644\u0627 \u062A\u0648\u062C\u062F \u0645\u0635\u0631\u0648\u0641\u0627\u062A \u0645\u0633\u062C\u0644\u0629 \u0644\u0647\u0630\u0647 \u0627\u0644\u0633\u064A\u0627\u0631\u0629 \u0628\u0639\u062F." })] })] }) : _jsxs(_Fragment, { children: [_jsxs("div", { className: "fleet-section-heading fleet-vehicles-heading", children: [_jsxs("div", { children: [_jsx("span", { className: "fleet-overline", children: "\u0628\u064A\u0627\u0646\u0627\u062A \u0627\u0644\u0623\u0633\u0637\u0648\u0644" }), _jsx("h2", { children: "\u0645\u0644\u062E\u0635 \u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A" })] }), _jsx("div", { className: "fleet-vehicle-actions", children: _jsx("button", { className: "fleet-button fleet-button-primary", type: "button", onClick: () => setVehicleDialog(true), children: "+ \u0625\u0636\u0627\u0641\u0629 \u0633\u064A\u0627\u0631\u0629" }) })] }), _jsx("div", { className: "fleet-vehicle-grid", children: summaries.map((summary) => {
                                    const { vehicle, lastOdometer, entries, distance, liters, expected, deviation, cost } = summary;
                                    return _jsxs("article", { className: "fleet-vehicle fleet-vehicle-card fleet-vehicle-tone", style: vehicleToneStyle(vehicle.id, data.vehicles), onClick: () => setSelectedVehicleId(vehicle.id), children: [_jsxs("div", { className: "fleet-vehicle-top", children: [_jsx("span", { className: "fleet-car-icon", children: _jsx("img", { src: "/car-mark.svg", alt: "" }) }), _jsx("span", { className: `fleet-gap ${Math.abs(deviation) > 200 ? 'gap-high' : 'gap-low'}`, children: currencyFormat.format(deviation) })] }), _jsx("h3", { children: vehicle.name }), _jsx("p", { className: "fleet-plate", children: vehicle.plate }), _jsxs("div", { className: "fleet-vehicle-data", children: [_jsxs("span", { children: ["\u0627\u0644\u0633\u0627\u0626\u0642", _jsx("strong", { children: vehicle.driver || 'غير محدد' })] }), _jsxs("span", { children: ["\u0645\u0639\u062F\u0644 \u0627\u0644\u0644\u062A\u0631\u0627\u062A", _jsxs("strong", { children: [numberFormat.format(vehicleFuelRate(vehicle)), " \u0644\u062A\u0631 / ", numberFormat.format(vehicleDistanceBase(vehicle)), " \u0643\u0645"] })] }), _jsxs("span", { children: ["\u0627\u0644\u0639\u062F\u0627\u062F \u0627\u0644\u062D\u0627\u0644\u064A", _jsxs("strong", { children: [numberFormat.format(lastOdometer), " \u0643\u0645"] })] }), _jsxs("span", { children: ["\u0627\u0644\u0645\u0633\u0627\u0641\u0629 \u0627\u0644\u0645\u0633\u062C\u0644\u0629", _jsxs("strong", { children: [numberFormat.format(distance), " \u0643\u0645"] })] }), _jsxs("span", { children: ["\u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0645\u062A\u0648\u0642\u0639 / \u0627\u0644\u0645\u0635\u0631\u0648\u0641", _jsxs("strong", { children: [numberFormat.format(expected), " / ", numberFormat.format(liters), " \u0644\u062A\u0631"] })] }), _jsxs("span", { children: ["\u0639\u062F\u062F \u0627\u0644\u062A\u0639\u0628\u0626\u0627\u062A", _jsx("strong", { children: entries.length })] }), _jsxs("span", { children: ["\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u062A\u0643\u0644\u0641\u0629", _jsx("strong", { children: currencyFormat.format(cost) })] })] }), _jsxs("div", { className: "fleet-card-actions", children: [_jsxs("button", { type: "button", className: "fleet-open-details", onClick: () => setSelectedVehicleId(vehicle.id), children: ["\u0639\u0631\u0636 \u062A\u0641\u0627\u0635\u064A\u0644 \u0627\u0644\u0633\u064A\u0627\u0631\u0629 ", _jsx("span", { "aria-hidden": "true", children: "\u2190" })] }), _jsx("button", { type: "button", className: "fleet-vehicle-edit", onClick: () => openVehicleEditor(vehicle), children: "\u062A\u0639\u062F\u064A\u0644 \u0627\u0644\u0633\u064A\u0627\u0631\u0629" })] })] }, vehicle.id);
                                }) })] })] }), tab === 'entries' && _jsxs("section", { className: "fleet-section", children: [_jsxs("div", { className: "fleet-section-heading", children: [_jsxs("div", { children: [_jsx("span", { className: "fleet-overline", children: "\u0627\u0644\u0645\u0633\u0627\u0641\u0629 \u0648\u0627\u0644\u0648\u0642\u0648\u062F \u0648\u0627\u0644\u062A\u0643\u0644\u0641\u0629" }), _jsx("h2", { children: "\u0633\u062C\u0644 \u0627\u0644\u062A\u0639\u0628\u0626\u0629" })] }), _jsx("button", { className: "fleet-button fleet-button-primary", type: "button", onClick: () => startEntry(), children: "+ \u062A\u0633\u062C\u064A\u0644 \u062A\u0639\u0628\u0626\u0629" })] }), _jsxs("div", { className: "fleet-search", children: [_jsx("span", { children: "\u2315" }), _jsx("input", { value: search, onChange: (event) => setSearch(event.target.value), placeholder: "\u0627\u0628\u062D\u062B \u0628\u0627\u0644\u0633\u064A\u0627\u0631\u0629 \u0623\u0648 \u0627\u0644\u0644\u0648\u062D\u0629 \u0623\u0648 \u0627\u0644\u0633\u0627\u0626\u0642" })] }), _jsx("div", { className: "fleet-table-wrap", children: _jsxs("table", { className: "fleet-table fleet-entry-table", children: [_jsx("thead", { children: _jsxs("tr", { children: [_jsx("th", { children: "\u0627\u0644\u062A\u0627\u0631\u064A\u062E" }), _jsx("th", { children: "\u0627\u0644\u0633\u064A\u0627\u0631\u0629" }), _jsx("th", { children: "\u0628\u062F\u0627\u064A\u0629 \u0627\u0644\u0639\u062F\u0627\u062F" }), _jsx("th", { children: "\u0627\u0644\u0639\u062F\u0627\u062F \u0627\u0644\u062D\u0627\u0644\u064A" }), _jsx("th", { children: "\u0627\u0644\u0641\u062C\u0648\u0629 (\u0643\u0645)" }), _jsx("th", { children: "\u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0645\u062A\u0648\u0642\u0639" }), _jsx("th", { children: "\u0627\u0644\u0645\u0635\u0631\u0648\u0641" }), _jsx("th", { children: "\u0627\u0646\u062D\u0631\u0627\u0641 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 (\u062C\u0646\u064A\u0647)" }), _jsx("th", { children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u0645\u0635\u0631\u0648\u0641" }), _jsx("th", { children: "\u0625\u062C\u0631\u0627\u0621" })] }) }), _jsx("tbody", { children: filteredEntries.map((entry) => { const vehicle = data.vehicles.find((item) => item.id === entry.vehicleId); const distance = Math.max(0, entry.currentOdometer - entry.startOdometer); const expected = vehicle ? litersForDistance(vehicle, distance) : 0; const liters = litersForEntry(entry, vehicle); const deviation = entryDeviation(entry, fuelAllowanceRate(data), liters); return _jsxs("tr", { className: "fleet-vehicle-tone", style: vehicleToneStyle(entry.vehicleId, data.vehicles), children: [_jsx("td", { children: formatDateArabic(entry.date) }), _jsxs("td", { children: [_jsx("strong", { children: vehicle?.name || 'سيارة محذوفة' }), _jsx("small", { children: vehicle?.plate })] }), _jsx("td", { children: numberFormat.format(entry.startOdometer) }), _jsx("td", { children: numberFormat.format(entry.currentOdometer) }), _jsxs("td", { children: [numberFormat.format(distance), " \u0643\u0645"] }), _jsxs("td", { children: [numberFormat.format(expected), " \u0644\u062A\u0631"] }), _jsxs("td", { children: [numberFormat.format(liters), " \u0644\u062A\u0631"] }), _jsx("td", { className: Math.abs(deviation) > 200 ? 'ledger-warning' : 'ledger-ok', children: currencyFormat.format(deviation) }), _jsx("td", { children: currencyFormat.format(entryTotalCost(entry)) }), _jsxs("td", { className: "ledger-actions", children: [_jsx("button", { type: "button", className: "entry-edit-button", onClick: () => editEntry(entry), children: "\u062A\u0639\u062F\u064A\u0644" }), _jsx("button", { type: "button", className: "entry-delete-button", onClick: () => openDeleteConfirmation(entry), children: "\u062D\u0630\u0641" })] })] }, entry.id); }) })] }) })] }), _jsxs("footer", { className: "fleet-footer", children: [_jsx("span", { children: "\u0645\u0631\u0627\u0642\u0628\u0629 \u0633\u064A\u0627\u0631\u0627\u062A \u0634\u0631\u0643\u0629 \u0623\u0628\u0648 \u063A\u0627\u0644\u064A" }), _jsx("span", { children: "\u0627\u0644\u0628\u064A\u0627\u0646\u0627\u062A \u0645\u062D\u0641\u0648\u0638\u0629 \u0639\u0644\u0649 \u0647\u0630\u0627 \u0627\u0644\u062C\u0647\u0627\u0632" })] }), vehicleDialog && _jsx("div", { className: "fleet-modal-backdrop", role: "presentation", onMouseDown: (event) => event.target === event.currentTarget && setVehicleDialog(false), children: _jsxs("form", { className: "fleet-modal", onSubmit: saveVehicle, children: [_jsxs("div", { className: "fleet-modal-title", children: [_jsxs("div", { children: [_jsx("span", { className: "fleet-overline", children: "\u0628\u064A\u0627\u0646\u0627\u062A \u0627\u0644\u0645\u0631\u0643\u0628\u0629" }), _jsx("h2", { children: editingVehicleId ? 'تعديل بيانات السيارة' : 'إضافة سيارة' })] }), _jsx("button", { type: "button", "aria-label": "\u0625\u063A\u0644\u0627\u0642", onClick: () => { setVehicleDialog(false); setEditingVehicleId(null); }, children: "\u00D7" })] }), _jsxs("label", { children: ["\u0627\u0633\u0645 \u0627\u0644\u0633\u064A\u0627\u0631\u0629", _jsx("input", { required: true, value: vehicleForm.name, onChange: (event) => setVehicleForm({ ...vehicleForm, name: event.target.value }), placeholder: "\u0645\u062B\u0627\u0644: \u062A\u0648\u064A\u0648\u062A\u0627 \u0647\u0627\u064A\u0644\u0648\u0643\u0633" })] }), _jsxs("label", { children: ["\u0631\u0642\u0645 \u0627\u0644\u0644\u0648\u062D\u0629", _jsx("input", { required: true, value: vehicleForm.plate, onChange: (event) => setVehicleForm({ ...vehicleForm, plate: event.target.value }) })] }), _jsxs("label", { children: ["\u0627\u0633\u0645 \u0627\u0644\u0633\u0627\u0626\u0642", _jsx("input", { value: vehicleForm.driver, onChange: (event) => setVehicleForm({ ...vehicleForm, driver: event.target.value }) })] }), _jsxs("div", { className: "fleet-form-row", children: [_jsxs("label", { children: ["\u0628\u062F\u0627\u064A\u0629 \u0627\u0644\u0639\u062F\u0627\u062F", _jsx("input", { required: true, type: "number", min: "0", value: vehicleForm.openingOdometer, onChange: (event) => setVehicleForm({ ...vehicleForm, openingOdometer: event.target.value }) })] }), _jsxs("label", { children: ["\u0627\u0644\u0644\u062A\u0631\u0627\u062A \u0627\u0644\u0645\u062D\u062F\u062F\u0629", _jsx("input", { required: true, type: "number", min: "0.1", step: "0.1", value: vehicleForm.kmPerLiter, onChange: (event) => setVehicleForm({ ...vehicleForm, kmPerLiter: event.target.value }) })] })] }), _jsxs("div", { className: "fleet-form-row", children: [_jsxs("label", { children: ["\u0645\u0633\u0627\u0641\u0629 \u0627\u0644\u0623\u0633\u0627\u0633 \u0628\u0627\u0644\u0643\u064A\u0644\u0648\u0645\u062A\u0631", _jsx("input", { required: true, type: "number", min: "1", step: "1", value: vehicleForm.distanceBaseKm, onChange: (event) => setVehicleForm({ ...vehicleForm, distanceBaseKm: event.target.value }) })] }), _jsx("span", { className: "fleet-rate-help", children: "\u0627\u0644\u0645\u0639\u0627\u062F\u0644\u0629: \u0627\u0644\u0641\u062C\u0648\u0629 \u00F7 \u0645\u0633\u0627\u0641\u0629 \u0627\u0644\u0623\u0633\u0627\u0633 \u00D7 \u0627\u0644\u0644\u062A\u0631\u0627\u062A" })] }), _jsx("button", { className: "fleet-button fleet-button-primary fleet-submit", type: "submit", children: editingVehicleId ? 'حفظ التعديل' : 'حفظ السيارة' })] }) }), entryDialog && _jsx("div", { className: "fleet-modal-backdrop fleet-entry-backdrop", role: "presentation", onMouseDown: (event) => event.target === event.currentTarget && setEntryDialog(false), children: _jsxs("form", { className: "fleet-modal fleet-entry-modal", onSubmit: saveEntry, children: [_jsxs("header", { className: "entry-form-header", children: [_jsxs("div", { children: [_jsx("span", { children: "\u0625\u062F\u0627\u0631\u0629 \u0627\u0644\u0623\u0633\u0637\u0648\u0644 / \u0633\u062C\u0644 \u0627\u0644\u0645\u0635\u0627\u0631\u064A\u0641" }), _jsx("h2", { children: editingEntryId ? 'تعديل حركة مسجلة' : 'إدخال حركة جديدة' })] }), _jsx("button", { type: "button", "aria-label": "\u0625\u063A\u0644\u0627\u0642", onClick: () => setEntryDialog(false), children: "\u00D7" })] }), _jsxs("div", { className: "entry-form-body", children: [_jsxs("section", { className: "entry-form-section", children: [_jsxs("div", { className: "entry-section-title", children: [_jsx("span", { children: "01" }), _jsxs("div", { children: [_jsx("h3", { children: "\u0628\u064A\u0627\u0646\u0627\u062A \u0627\u0644\u0631\u062D\u0644\u0629" }), _jsx("p", { children: "\u0627\u0644\u0633\u064A\u0627\u0631\u0629 \u0648\u0627\u0644\u062A\u0627\u0631\u064A\u062E \u0648\u0627\u0644\u0633\u0627\u0626\u0642 \u0648\u0627\u0644\u0645\u0633\u0627\u0631" })] })] }), _jsxs("div", { className: "entry-form-grid entry-trip-grid", children: [_jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u0627\u0644\u0633\u064A\u0627\u0631\u0629 / \u0631\u0642\u0645 \u0627\u0644\u0644\u0648\u062D\u0629" }), _jsx("select", { required: true, value: entryForm.vehicleId, onChange: (event) => { const next = event.target.value; const vehicleSummary = summaries.find((item) => item.vehicle.id === next); const latest = vehicleSummary?.lastOdometer ?? 0; setEntryForm({ ...entryForm, vehicleId: next, driver: vehicleSummary?.vehicle.driver || '', startOdometer: String(latest), currentOdometer: '', liters: '0.0' }); }, children: data.vehicles.map((vehicle) => _jsxs("option", { value: vehicle.id, children: [vehicle.plate, " \u00B7 ", vehicle.name] }, vehicle.id)) })] }), _jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u0627\u0644\u062A\u0627\u0631\u064A\u062E" }), _jsx("input", { required: true, type: "date", value: entryForm.date, onChange: (event) => setEntryForm({ ...entryForm, date: event.target.value }) })] }), _jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u0627\u0633\u0645 \u0627\u0644\u0633\u0627\u0626\u0642" }), _jsx("input", { value: entryForm.driver, onChange: (event) => setEntryForm({ ...entryForm, driver: event.target.value }) })] }), _jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u062E\u0637 \u0627\u0644\u0633\u064A\u0631" }), _jsx("input", { value: entryForm.route, onChange: (event) => setEntryForm({ ...entryForm, route: event.target.value }), placeholder: "\u0646\u0642\u0637\u0629 \u0627\u0644\u0627\u0646\u0637\u0644\u0627\u0642 \u0648\u0627\u0644\u0648\u062C\u0647\u0629" })] })] })] }), _jsxs("section", { className: "entry-form-section", children: [_jsxs("div", { className: "entry-section-title", children: [_jsx("span", { children: "02" }), _jsxs("div", { children: [_jsx("h3", { children: "\u0627\u0644\u0639\u062F\u0627\u062F \u0648\u0627\u0644\u062A\u0645\u0648\u064A\u0646" }), _jsx("p", { children: "\u064A\u064F\u062D\u0633\u0628 \u0641\u0631\u0642 \u0627\u0644\u0639\u062F\u0627\u062F \u062A\u0644\u0642\u0627\u0626\u064A\u064B\u0627" })] })] }), _jsxs("div", { className: "entry-form-grid entry-fuel-grid", children: [_jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u0627\u0644\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u0633\u0627\u0628\u0642\u0629 \u00B7 \u0622\u062E\u0631 \u0642\u0631\u0627\u0621\u0629" }), _jsx("input", { required: true, type: "number", min: "0", value: entryForm.startOdometer, readOnly: true })] }), _jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u0627\u0644\u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u062D\u0627\u0644\u064A\u0629" }), _jsx("input", { required: true, type: "number", min: "0", value: entryForm.currentOdometer, onChange: (event) => { const currentOdometer = event.target.value; setEntryForm({ ...entryForm, currentOdometer, liters: calculateFuelLiters(entryForm.vehicleId, Number(entryForm.startOdometer), Number(currentOdometer)).toFixed(1) }); } })] }), _jsxs("div", { className: "entry-distance", children: [_jsx("span", { children: "\u0627\u0644\u0645\u0633\u0627\u0641\u0629" }), _jsxs("strong", { children: [numberFormat.format(entryDistance), " ", _jsx("small", { children: "\u0643\u0645" })] })] }), _jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0627\u0644\u0641\u0639\u0644\u064A (\u0644\u062A\u0631)" }), _jsx("input", { type: "number", min: "0", step: "0.1", value: entryForm.liters, readOnly: true }), _jsx("small", { className: "entry-field-help", children: "\u064A\u064F\u062D\u0633\u0628 \u0645\u0646 \u0641\u0631\u0642 \u0627\u0644\u0639\u062F\u0627\u062F \u0648\u0645\u0639\u062F\u0644 \u0627\u0633\u062A\u0647\u0644\u0627\u0643 \u0627\u0644\u0633\u064A\u0627\u0631\u0629" })] }), _jsxs("label", { className: "entry-field", children: [_jsx("span", { children: "\u0642\u064A\u0645\u0629 \u0627\u0644\u062A\u0645\u0648\u064A\u0646 \u0648\u0641\u0642\u064B\u0627 \u0644\u0644\u062A\u0631 (\u062C\u0646\u064A\u0647)" }), _jsx("input", { type: "number", min: "0", step: "1", value: calculateAllowanceValue(entryForm.vehicleId, calculateFuelLiters(entryForm.vehicleId, Number(entryForm.startOdometer), Number(entryForm.currentOdometer))), readOnly: true }), _jsxs("small", { className: "entry-field-help", children: [rateFormat.format(fuelAllowanceRate(data)), " \u062C\u0646\u064A\u0647 \u0644\u0643\u0644 \u0644\u062A\u0631 \u0644\u062C\u0645\u064A\u0639 \u0627\u0644\u0633\u064A\u0627\u0631\u0627\u062A"] })] }), _jsxs("div", { className: "entry-field", children: [_jsx("label", { className: "entry-field-label", htmlFor: "fuel-amount-input", children: "\u0645\u0635\u0631\u0648\u0641 \u0627\u0644\u0648\u0642\u0648\u062F \u0627\u0644\u0641\u0639\u0644\u064A (\u062C\u0646\u064A\u0647)" }), _jsxs("div", { className: "entry-amount-total", children: [_jsx("span", { children: "\u0627\u0644\u0625\u062C\u0645\u0627\u0644\u064A" }), _jsx("output", { "aria-live": "polite", children: fuelActualValuePreview === null ? '—' : currencyFormat.format(fuelActualValuePreview) })] }), _jsxs("div", { className: "entry-amount-entry", children: [_jsx("input", { id: "fuel-amount-input", dir: "ltr", type: "text", inputMode: "decimal", value: fuelAmountDraft, onChange: (event) => setFuelAmountDraft(event.target.value), onKeyDown: (event) => { if (event.key === 'Enter') {
                                                                        event.preventDefault();
                                                                        addFuelAmount();
                                                                    } } }), _jsx("button", { type: "button", className: "entry-amount-add", onClick: addFuelAmount, children: "\u0625\u0636\u0627\u0641\u0629" })] }), _jsx("div", { className: "entry-amount-list", "aria-label": "\u0627\u0644\u0645\u0628\u0627\u0644\u063A \u0627\u0644\u0645\u0636\u0627\u0641\u0629 \u0644\u0645\u0635\u0631\u0648\u0641 \u0627\u0644\u0648\u0642\u0648\u062F", children: splitAmountExpression(entryForm.fuelActualValue).map((amount, index) => _jsxs("span", { className: "entry-amount-chip", children: [_jsx("span", { dir: "ltr", children: amount }), _jsx("button", { type: "button", "aria-label": `حذف مبلغ ${amount} من الوقود`, onClick: () => removeFuelAmount(index), children: "\u00D7" })] }, `${amount}-${index}`)) }), fuelActualValuePreview === null && _jsx("small", { className: "entry-field-help", role: "alert", children: "\u0631\u0627\u062C\u0639 \u0627\u0644\u0645\u0628\u0627\u0644\u063A \u0627\u0644\u0645\u0636\u0627\u0641\u0629" })] })] })] }), _jsxs("section", { className: "entry-form-section", children: [_jsxs("div", { className: "entry-section-title", children: [_jsx("span", { children: "03" }), _jsxs("div", { children: [_jsx("h3", { children: "\u0627\u0644\u0645\u0635\u0627\u0631\u064A\u0641 \u0627\u0644\u0625\u0636\u0627\u0641\u064A\u0629" }), _jsx("p", { children: "\u0623\u062F\u062E\u0644 \u0643\u0644 \u0645\u0628\u0644\u063A \u0648\u0627\u0636\u063A\u0637 \u0625\u0636\u0627\u0641\u0629 \u0644\u064A\u064F\u062D\u0633\u0628 \u0641\u064A \u0627\u0644\u0625\u062C\u0645\u0627\u0644\u064A" })] })] }), _jsx("div", { className: "entry-expense-grid", children: expenseFields.map((field) => {
                                                const amount = expenseAmountPreviews[field.key];
                                                return _jsxs("div", { className: "entry-field", children: [_jsx("label", { className: "entry-field-label", htmlFor: `expense-amount-${field.key}`, children: field.label }), _jsxs("div", { className: "entry-amount-total", children: [_jsx("span", { children: "\u0627\u0644\u0625\u062C\u0645\u0627\u0644\u064A" }), _jsx("output", { "aria-live": "polite", children: amount === null ? '—' : currencyFormat.format(amount) })] }), _jsxs("div", { className: "entry-amount-entry", children: [_jsx("input", { id: `expense-amount-${field.key}`, dir: "ltr", type: "text", inputMode: "decimal", "aria-label": `${field.label} بالجنيه`, value: expenseAmountDrafts[field.key], onChange: (event) => setExpenseAmountDrafts((current) => ({ ...current, [field.key]: event.target.value })), onKeyDown: (event) => { if (event.key === 'Enter') {
                                                                        event.preventDefault();
                                                                        addExpenseAmount(field.key);
                                                                    } } }), _jsx("button", { type: "button", className: "entry-amount-add", onClick: () => addExpenseAmount(field.key), children: "\u0625\u0636\u0627\u0641\u0629" })] }), _jsx("div", { className: "entry-amount-list", "aria-label": `المبالغ المضافة: ${field.label}`, children: splitAmountExpression(entryForm.expenses[field.key]).map((part, index) => _jsxs("span", { className: "entry-amount-chip", children: [_jsx("span", { dir: "ltr", children: part }), _jsx("button", { type: "button", "aria-label": `حذف ${part} من ${field.label}`, onClick: () => removeExpenseAmount(field.key, index), children: "\u00D7" })] }, `${part}-${index}`)) }), amount === null && _jsx("small", { className: "entry-field-help", role: "alert", children: "\u0631\u0627\u062C\u0639 \u0627\u0644\u0645\u0628\u0627\u0644\u063A \u0627\u0644\u0645\u0636\u0627\u0641\u0629" })] }, field.key);
                                            }) })] })] }), _jsxs("footer", { className: "entry-form-actions", children: [_jsxs("div", { children: [_jsx("span", { children: "\u0625\u062C\u0645\u0627\u0644\u064A \u0627\u0644\u062D\u0631\u0643\u0629" }), _jsx("strong", { children: hasInvalidAmount ? 'راجع المبالغ' : hasUnaddedAmounts ? 'اضغط إضافة أولًا' : currencyFormat.format(entryCostPreview) })] }), _jsx("button", { type: "button", className: "entry-cancel-button", onClick: () => { setEntryDialog(false); setEditingEntryId(null); }, children: "\u0625\u0644\u063A\u0627\u0621" }), _jsx("button", { className: "fleet-button fleet-button-primary entry-save-button", type: "submit", children: editingEntryId ? 'حفظ التعديل' : 'حفظ الحركة' })] })] }) }), pendingDeleteEntry && _jsx("div", { className: "fleet-confirm-backdrop", role: "presentation", children: _jsxs("div", { className: "fleet-confirm-dialog", role: "alertdialog", "aria-modal": "true", "aria-labelledby": "delete-entry-title", children: [_jsx("span", { className: "fleet-confirm-icon", children: "!" }), _jsx("h2", { id: "delete-entry-title", children: "\u062A\u0623\u0643\u064A\u062F \u062D\u0630\u0641 \u0627\u0644\u0633\u062C\u0644" }), _jsxs("p", { children: ["\u0644\u062D\u0630\u0641 \u062D\u0631\u0643\u0629 \u0627\u0644\u062A\u0639\u0628\u0626\u0629 \u0628\u062A\u0627\u0631\u064A\u062E ", _jsx("strong", { children: formatDateArabic(pendingDeleteEntry.date) }), " \u0627\u0643\u062A\u0628 \u0643\u0644\u0645\u0629 ", _jsx("strong", { children: "\u0645\u0648\u0627\u0641\u0642" }), " \u062B\u0645 \u0627\u0636\u063A\u0637 \u062A\u0623\u0643\u064A\u062F."] }), _jsx("input", { className: "fleet-confirm-input", "aria-label": "\u0627\u0643\u062A\u0628 \u0645\u0648\u0627\u0641\u0642 \u0644\u0644\u062A\u0623\u0643\u064A\u062F", value: deleteConfirmationText, onChange: (event) => setDeleteConfirmationText(event.target.value), autoFocus: true, placeholder: "\u0627\u0643\u062A\u0628: \u0645\u0648\u0627\u0641\u0642" }), _jsxs("div", { children: [_jsx("button", { type: "button", className: "fleet-confirm-cancel", onClick: cancelDeleteConfirmation, children: "\u0625\u0644\u063A\u0627\u0621" }), _jsx("button", { type: "button", className: "fleet-confirm-ok", disabled: deleteConfirmationText.trim() !== 'موافق', onClick: confirmDeleteEntry, children: "\u062A\u0623\u0643\u064A\u062F \u0627\u0644\u062D\u0630\u0641" })] })] }) })] }));
};
export default CarFuelPage;
