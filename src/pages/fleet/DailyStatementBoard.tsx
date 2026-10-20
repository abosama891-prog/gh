import React, { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { Download, Pencil, Plus, Trash2 } from 'lucide-react';

export type DailyStatementRow = {
  id: string;
  description: string;
  plate: string;
  departure: string;
  route: string;
  representative: string;
  driver: string;
  note: string;
};

const DAILY_STATEMENT_STORAGE_KEY = 'company-fleet-daily-statement-v1';
export const DAILY_STATEMENT_DESCRIPTIONS = ['جامبو', 'ديماكس', 'ميني باص'];
const createEmptyForm = () => ({ description: 'جامبو', plate: '', departure: '', route: '', representative: '', driver: '', note: '' });

const readRows = (): DailyStatementRow[] => {
  try {
    const raw = localStorage.getItem(DAILY_STATEMENT_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((row) => row && typeof row.id === 'string') : [];
  } catch {
    return [];
  }
};

type Props = { canEdit: boolean; dateLabel: string; notify: (message: string, isError: boolean) => void };

const DailyStatementBoard: React.FC<Props> = ({ canEdit, dateLabel, notify }) => {
  const [rows, setRows] = useState<DailyStatementRow[]>(readRows);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [form, setForm] = useState(createEmptyForm);
  const [isExporting, setIsExporting] = useState(false);

  useEffect(() => {
    try {
      localStorage.setItem(DAILY_STATEMENT_STORAGE_KEY, JSON.stringify(rows));
    } catch {
    }
  }, [rows]);

  const updateField = (field: keyof ReturnType<typeof createEmptyForm>, value: string) => {
    setForm((current) => ({ ...current, [field]: value }));
  };

  const openAddModal = () => {
    setEditingId(null);
    setForm(createEmptyForm());
    setIsModalOpen(true);
  };

  const openEditModal = (row: DailyStatementRow) => {
    setEditingId(row.id);
    setForm({ description: row.description, plate: row.plate, departure: row.departure, route: row.route, representative: row.representative, driver: row.driver, note: row.note });
    setIsModalOpen(true);
  };

  const handleDelete = (id: string) => {
    setRows((current) => current.filter((row) => row.id !== id));
    notify('تم حذف السيارة من البيان', false);
  };

  const handleSubmit = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!form.plate.trim()) {
      notify('برجاء إدخال رقم السيارة', true);
      return;
    }
    if (editingId) {
      setRows((current) => current.map((row) => (row.id === editingId ? { ...row, ...form, plate: form.plate.trim() } : row)));
      notify('تم حفظ بيانات السيارة في البيان', false);
    } else {
      setRows((current) => [...current, { ...form, plate: form.plate.trim(), id: `ds-${Date.now()}-${Math.random().toString(36).slice(2, 7)}` }]);
      notify('تمت إضافة السيارة إلى البيان', false);
    }
    setIsModalOpen(false);
  };
  const exportToExcel = async () => {
    setIsExporting(true);
    try {
      const ExcelJS = (await import('exceljs')).default;
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'fleet-app';
      workbook.subject = 'daily-statement';
      const worksheet = workbook.addWorksheet('sheet1');
      worksheet.columns = [{ width: 8 }, { width: 16 }, { width: 19 }, { width: 26 }, { width: 34 }, { width: 24 }, { width: 24 }, { width: 30 }];
      worksheet.mergeCells('A1:H1');
      worksheet.getCell('A1').value = 'daily statement';
      worksheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      worksheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9A682C' } };
      worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(1).height = 32;
      const headerRow = worksheet.addRow(['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8']);
      headerRow.height = 26;
      headerRow.eachCell((cell) => {
        cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5D4025' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      rows.forEach((row, index) => {
        const excelRow = worksheet.addRow([index + 1, row.description, row.plate, row.departure, row.route, row.representative, row.driver, row.note]);
        excelRow.eachCell((cell) => {
          cell.font = { name: 'Arial', size: 11 };
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          if (index % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAF4E9' } };
        });
      });
      const buffer = await workbook.xlsx.writeBuffer();
      const url = URL.createObjectURL(new Blob([buffer as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const link = document.createElement('a');
      link.href = url;
      link.download = 'daily-statement.xlsx';
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      notify('done', false);
    } catch (error) {
      console.error('export failed', error);
      notify('failed', true);
    } finally {
      setIsExporting(false);
    }
  };
  return (
    <div className="daily-stmt-board">
      <div className="daily-stmt-toolbar">
        {canEdit && <button type="button" className="fleet-button fleet-button-primary" onClick={openAddModal}><Plus size={16} aria-hidden="true" /> إضافة سيارة للبيان</button>}
        <button type="button" className="fleet-button fleet-button-light" onClick={exportToExcel} disabled={isExporting || rows.length === 0}><Download size={16} aria-hidden="true" /> {isExporting ? 'جارٍ التصدير…' : 'تصدير إلى Excel'}</button>
        <span className="daily-stmt-count">عدد السيارات: {rows.length}</span>
      </div>
      {rows.length === 0 ? (
        <div className="fleet-empty-state">لا توجد سيارات في البيان بعد. اضغط "إضافة سيارة للبيان" لإضافة أول سيارة.</div>
      ) : (
        <>
          <div className="daily-stmt-table-wrap">
            <table className="daily-stmt-table">
              <thead>
                <tr><th>م</th><th>الوصف</th><th>رقم السيارة</th><th>نقطة الإنطلاق</th><th>خط السير</th><th>المندوب</th><th>السائق</th><th>إجراءات</th></tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={row.id}>
                    <td>{index + 1}</td>
                    <td><strong>{row.description || '—'}</strong></td>
                    <td><strong dir="ltr">{row.plate || '—'}</strong></td>
                    <td>{row.departure || '—'}</td>
                    <td>{row.route || '—'}</td>
                    <td>{row.representative || '—'}</td>
                    <td>{row.driver || '—'}</td>
                    <td>
                      <div className="daily-stmt-actions">
                        {canEdit && <button type="button" className="daily-stmt-action" onClick={() => openEditModal(row)} aria-label="تعديل"><Pencil size={15} aria-hidden="true" /></button>}
                        {canEdit && <button type="button" className="daily-stmt-action daily-stmt-action-danger" onClick={() => handleDelete(row.id)} aria-label="حذف"><Trash2 size={15} aria-hidden="true" /></button>}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="daily-stmt-cards">
            {rows.map((row, index) => (
              <article key={row.id} className="daily-stmt-card">
                <header className="daily-stmt-card-head"><span className="daily-stmt-card-index">{index + 1}</span><strong>{row.description || '—'}</strong><strong dir="ltr">{row.plate || '—'}</strong></header>
                <dl className="daily-stmt-card-body">
                  <div><dt>نقطة الإنطلاق</dt><dd>{row.departure || '—'}</dd></div>
                  <div><dt>خط السير</dt><dd>{row.route || '—'}</dd></div>
                  <div><dt>المندوب</dt><dd>{row.representative || '—'}</dd></div>
                  <div><dt>السائق</dt><dd>{row.driver || '—'}</dd></div>
                  {row.note ? <div className="daily-stmt-card-note"><dt>ملاحظات</dt><dd>{row.note}</dd></div> : null}
                </dl>
                {canEdit && (
                  <footer className="daily-stmt-card-foot">
                    <button type="button" className="daily-stmt-action" onClick={() => openEditModal(row)} aria-label="تعديل"><Pencil size={15} aria-hidden="true" /> تعديل</button>
                    <button type="button" className="daily-stmt-action daily-stmt-action-danger" onClick={() => handleDelete(row.id)} aria-label="حذف"><Trash2 size={15} aria-hidden="true" /> حذف</button>
                  </footer>
                )}
              </article>
            ))}
          </div>
        </>
      )}
      {isModalOpen && createPortal(
        <div className="fleet-modal-backdrop" onClick={() => setIsModalOpen(false)}>
          <form className="fleet-modal daily-stmt-modal" onClick={(event) => event.stopPropagation()} onSubmit={handleSubmit}>
            <div className="daily-stmt-modal-head"><h3>{editingId ? 'تعديل سيارة في البيان' : 'إضافة سيارة للبيان'}</h3></div>
            <div className="daily-stmt-form-grid">
              <label className="entry-field"><span>الوصف</span>
                <select value={form.description} onChange={(event) => updateField('description', event.target.value)}>
                  {DAILY_STATEMENT_DESCRIPTIONS.map((option) => <option key={option} value={option}>{option}</option>)}
                </select>
              </label>
              <label className="entry-field"><span>رقم السيارة</span><input type="text" dir="ltr" value={form.plate} onChange={(event) => updateField('plate', event.target.value)} placeholder="مثال: 1234" /></label>
              <label className="entry-field"><span>نقطة الإنطلاق</span><input type="text" value={form.departure} onChange={(event) => updateField('departure', event.target.value)} placeholder="نقطة الإنطلاق" /></label>
              <label className="entry-field"><span>خط السير</span><input type="text" value={form.route} onChange={(event) => updateField('route', event.target.value)} placeholder="خط السير" /></label>
              <label className="entry-field"><span>المندوب</span><input type="text" value={form.representative} onChange={(event) => updateField('representative', event.target.value)} placeholder="اسم المندوب" /></label>
              <label className="entry-field"><span>السائق</span><input type="text" value={form.driver} onChange={(event) => updateField('driver', event.target.value)} placeholder="اسم السائق" /></label>
              <label className="entry-field entry-field-wide"><span>ملاحظات</span><input type="text" value={form.note} onChange={(event) => updateField('note', event.target.value)} placeholder="ملاحظات (اختياري)" /></label>
            </div>
            <div className="daily-stmt-modal-actions">
              <button type="button" className="fleet-button fleet-button-light" onClick={() => setIsModalOpen(false)}>إلغاء</button>
              <button type="submit" className="fleet-button fleet-button-primary">{editingId ? 'حفظ التعديل' : 'إضافة للبيان'}</button>
            </div>
          </form>
        </div>, document.body)}
    </div>
  );
};

export default DailyStatementBoard;