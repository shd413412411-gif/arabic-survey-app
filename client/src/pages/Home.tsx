import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";

declare global {
  interface Window {
    XLSX?: any;
  }
}

type AnswerMap = Record<string, number>;
type CurrentForm = { answers: AnswerMap; recordId?: string | null };
type SurveyRecord = {
  id: string;
  studentNumber: number | string;
  answers: AnswerMap;
  date: string;
};

const STORAGE = {
  records: "surveyAppData",
  nextNumber: "surveyAppNextNumber",
  currentForm: "surveyAppCurrentForm",
};
const QUESTION_COUNT = 28;
const REVERSE_ITEMS = new Set([1, 8, 9, 12, 20]);
const QUESTION_TEXTS: string[] = Array.from({ length: QUESTION_COUNT }, () => "");
const SCALE = [1, 2, 3, 4, 5];
const EMPTY_FORM: CurrentForm = { answers: {}, recordId: null };

function makeId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function safeRead<T>(key: string, fallback: T): T {
  try {
    const stored = localStorage.getItem(key);
    return stored ? (JSON.parse(stored) as T) : fallback;
  } catch {
    return fallback;
  }
}

function reverseScore(question: number, value: number) {
  return value === 0 ? 0 : REVERSE_ITEMS.has(question) ? 6 - value : value;
}

function normalizeNumber(value: unknown) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 && number <= 5 ? Math.round(number) : 0;
}

function initialRecords(): SurveyRecord[] {
  const raw = safeRead<unknown>(STORAGE.records, []);
  if (!Array.isArray(raw)) return [];
  return raw.map((record: any, index: number) => {
    const answers: AnswerMap = {};
    for (let q = 1; q <= QUESTION_COUNT; q += 1) {
      answers[String(q)] = normalizeNumber(record?.answers?.[q] ?? record?.answers?.[String(q)] ?? record?.[`Q${q}`]);
    }
    return {
      id: String(record?.id ?? makeId()),
      studentNumber: record?.studentNumber ?? index + 1,
      answers,
      date: String(record?.date ?? record?.["التاريخ"] ?? new Date().toISOString()),
    };
  });
}

function scoreForm(raw: AnswerMap): AnswerMap {
  const result: AnswerMap = {};
  for (let q = 1; q <= QUESTION_COUNT; q += 1) {
    const value = normalizeNumber(raw[String(q)] ?? 0);
    result[String(q)] = reverseScore(q, value);
  }
  return result;
}

function getQuestionText(question: number) {
  return QUESTION_TEXTS[question - 1]?.trim() || `لم يُرفق نص البند ${question}؛ أضف النص في مصفوفة QUESTION_TEXTS داخل Home.tsx.`;
}

function parseStudentNumber(value: unknown, fallback: number) {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric >= 0 ? numeric : fallback;
}

function normalizeImportedRecord(source: any, index: number): SurveyRecord {
  const nested = source?.answers && typeof source.answers === "object" ? source.answers : {};
  const answers: AnswerMap = {};
  for (let q = 1; q <= QUESTION_COUNT; q += 1) {
    const value = nested[q] ?? nested[String(q)] ?? nested[`Q${q}`] ?? source?.[`Q${q}`] ?? source?.[`Q${q}*`] ?? source?.[String(q)];
    answers[String(q)] = normalizeNumber(value);
  }
  const dateValue = source?.date ?? source?.["التاريخ"] ?? source?.Date ?? new Date().toISOString();
  const date = dateValue instanceof Date ? dateValue.toISOString() : String(dateValue);
  return {
    id: makeId(),
    studentNumber: source?.studentNumber ?? source?.["رقم الطالب"] ?? source?.student_number ?? index + 1,
    answers,
    date,
  };
}

function isEmptyRow(row: unknown[]) {
  return row.every((value) => value === undefined || value === null || String(value).trim() === "");
}

function importWorkbook(file: File): Promise<SurveyRecord[]> {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) {
      reject(new Error("مكتبة Excel غير متاحة. اتصل بالإنترنت مرة واحدة لتهيئتها، ثم أعد المحاولة."));
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("تعذر قراءة الملف."));
    reader.onload = (event) => {
      try {
        const workbook = window.XLSX.read(event.target?.result, { type: "array", cellDates: true });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        if (!firstSheet) throw new Error("لا توجد ورقة بيانات في الملف.");
        const rows: unknown[][] = window.XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "", raw: true, blankrows: false });
        if (rows.length < 2) {
          resolve([]);
          return;
        }
        const headers = (rows[0] as unknown[]).map((value) => String(value ?? "").trim());
        const studentColumn = headers.findIndex((header) => header === "رقم الطالب" || header.toLowerCase() === "studentnumber");
        const dateColumn = headers.findIndex((header) => header === "التاريخ" || header.toLowerCase() === "date");
        const questionColumns = new Map<number, number>();
        headers.forEach((header, column) => {
          const match = header.match(/^Q\s*(\d+)\s*\*?$/i);
          if (match) questionColumns.set(Number(match[1]), column);
        });
        const records: SurveyRecord[] = [];
        rows.slice(1).forEach((row, rowIndex) => {
          const values = row as unknown[];
          if (isEmptyRow(values)) return;
          const source: Record<string, unknown> = {};
          source.studentNumber = studentColumn >= 0 ? values[studentColumn] : rowIndex + 1;
          source.date = dateColumn >= 0 && values[dateColumn] ? values[dateColumn] : new Date().toISOString();
          questionColumns.forEach((column, question) => {
            if (question >= 1 && question <= QUESTION_COUNT) source[`Q${question}`] = values[column];
          });
          records.push(normalizeImportedRecord(source, records.length + 1));
        });
        resolve(records);
      } catch (error) {
        reject(error instanceof Error ? error : new Error("تعذر قراءة ملف Excel."));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

export default function Home() {
  const [records, setRecords] = useState<SurveyRecord[]>(initialRecords);
  const [nextStudentNumber, setNextStudentNumber] = useState<number>(() => {
    const stored = Number(localStorage.getItem(STORAGE.nextNumber));
    if (Number.isFinite(stored) && stored > 0) return Math.floor(stored);
    const current = safeRead<SurveyRecord[]>(STORAGE.records, []);
    const maximum = Array.isArray(current) ? Math.max(0, ...current.map((record: any) => Number(record?.studentNumber) || 0)) : 0;
    return maximum + 1;
  });
  const [currentForm, setCurrentForm] = useState<CurrentForm>(() => {
    const saved = safeRead<CurrentForm>(STORAGE.currentForm, EMPTY_FORM);
    return saved && typeof saved.answers === "object" ? { ...EMPTY_FORM, ...saved } : EMPTY_FORM;
  });
  const [activeTab, setActiveTab] = useState<"list" | "row" | "saved">("list");
  const [currentQuestion, setCurrentQuestion] = useState(1);
  const [toast, setToast] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const rowTableRef = useRef<HTMLDivElement>(null);
  const toastTimer = useRef<number | undefined>(undefined);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE.records, JSON.stringify(records));
      localStorage.setItem(STORAGE.nextNumber, String(nextStudentNumber));
      localStorage.setItem(STORAGE.currentForm, JSON.stringify(currentForm));
    } catch {
      showToast("تعذر الحفظ المحلي؛ مساحة التخزين قد تكون ممتلئة.");
    }
  }, [records, nextStudentNumber, currentForm]);

  useEffect(() => {
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("/service-worker.js").catch(() => undefined);
    }
    return () => window.clearTimeout(toastTimer.current);
  }, []);

  useEffect(() => {
    if (activeTab === "row" && rowTableRef.current) {
      const cellWidth = 58;
      window.setTimeout(() => {
        rowTableRef.current?.scrollTo({ left: Math.max(0, (currentQuestion - 1) * cellWidth - 110), behavior: "smooth" });
      }, 30);
    }
  }, [activeTab, currentQuestion]);

  function showToast(message: string) {
    setToast(message);
    window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast(""), 2600);
  }

  const currentValues = useMemo(() => scoreForm(currentForm.answers), [currentForm.answers]);
  const answeredCount = useMemo(() => Object.keys(currentForm.answers).filter((question) => normalizeNumber(currentForm.answers[question]) > 0).length, [currentForm.answers]);

  function selectAnswer(question: number, value: number) {
    const currentValue = currentForm.answers[String(question)];
    const isCancel = currentValue === value;
    setCurrentForm((previous) => {
      const answers = { ...previous.answers };
      if (answers[String(question)] === value) delete answers[String(question)];
      else answers[String(question)] = value;
      return { ...previous, answers };
    });
    return !isCancel;
  }

  function commitCurrentForm(startNew: boolean) {
    const existingId = currentForm.recordId || makeId();
    const record: SurveyRecord = {
      id: existingId,
      studentNumber: nextStudentNumber,
      answers: scoreForm(currentForm.answers),
      date: new Date().toISOString(),
    };
    setRecords((previous) => {
      const exists = previous.some((item) => item.id === existingId);
      return exists ? previous.map((item) => item.id === existingId ? record : item) : [...previous, record];
    });
    if (startNew) {
      setNextStudentNumber((number) => number + 1);
      setCurrentForm(EMPTY_FORM);
      setCurrentQuestion(1);
      showToast("تم الحفظ وفتح نموذج طالب جديد ✓");
    } else {
      setCurrentForm((previous) => ({ ...previous, recordId: existingId }));
      showToast("تم حفظ بيانات الطالب ✓");
    }
  }

  function clearCurrentForm() {
    if (answeredCount > 0 && !window.confirm("سيتم مسح إجابات النموذج الحالي. هل تريد المتابعة؟")) return;
    setCurrentForm(EMPTY_FORM);
    setCurrentQuestion(1);
    showToast("تم مسح النموذج");
  }

  function deleteRecord(id: string) {
    setRecords((previous) => previous.filter((record) => record.id !== id));
    showToast("تم حذف سجل الطالب");
  }

  function deleteAllRecords() {
    if (!records.length) return showToast("لا توجد بيانات لحذفها");
    if (!window.confirm(`سيتم حذف جميع السجلات المحفوظة (${records.length}). هل تريد المتابعة؟`)) return;
    setRecords([]);
    showToast("تم حذف جميع البيانات المحفوظة");
  }

  function exportExcel() {
    if (!window.XLSX) {
      showToast("مكتبة Excel غير متاحة. اتصل بالإنترنت مرة واحدة ثم أعد المحاولة.");
      return;
    }
    try {
      const headers = ["رقم الطالب", ...Array.from({ length: QUESTION_COUNT }, (_, index) => {
        const question = index + 1;
        return `Q${question}${REVERSE_ITEMS.has(question) ? "*" : ""}`;
      }), "التاريخ"];
      const rows = records.map((record) => {
        const row: Record<string, unknown> = { "رقم الطالب": record.studentNumber };
        for (let question = 1; question <= QUESTION_COUNT; question += 1) {
          const header = `Q${question}${REVERSE_ITEMS.has(question) ? "*" : ""}`;
          row[header] = normalizeNumber(record.answers[String(question)] ?? 0);
        }
        row["التاريخ"] = record.date;
        return row;
      });
      const worksheet = window.XLSX.utils.json_to_sheet(rows, { header: headers });
      const workbook = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(workbook, worksheet, "البيانات");
      window.XLSX.writeFile(workbook, "survey-data.xlsx");
      showToast("تم تصدير ملف Excel ✓");
    } catch {
      showToast("تعذر إنشاء ملف Excel.");
    }
  }

  async function handleImport(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      let imported: SurveyRecord[];
      if (file.name.toLowerCase().endsWith(".json")) {
        const parsed: unknown = JSON.parse(await file.text());
        if (!Array.isArray(parsed)) throw new Error("يجب أن يحتوي الملف على مصفوفة سجلات");
        imported = parsed.map((record, index) => normalizeImportedRecord(record, index + 1));
      } else if (/\.(xlsx|xls)$/i.test(file.name)) {
        imported = await importWorkbook(file);
      } else {
        throw new Error("صيغة الملف غير مدعومة. اختر JSON أو Excel.");
      }
      const replace = window.confirm("موافق: استبدال البيانات الحالية بالكامل.\nإلغاء: إضافة السجلات إلى البيانات الموجودة.");
      setRecords((previous) => replace ? imported : [...previous, ...imported]);
      const maximumImported = imported.reduce((maximum, record) => Math.max(maximum, Number(record.studentNumber) || 0), 0);
      setNextStudentNumber((current) => maximumImported >= current ? maximumImported + 1 : current);
      showToast("تم الاستيراد بنجاح ✓");
    } catch (error) {
      showToast(error instanceof Error ? error.message : "تعذر استيراد الملف.");
    } finally {
      event.target.value = "";
    }
  }

  async function refreshApp() {
    showToast("جارٍ تحديث التطبيق…");
    try {
      if ("serviceWorker" in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.all(registrations.map((registration) => registration.unregister()));
      }
      if ("caches" in window) {
        const keys = await caches.keys();
        await Promise.all(keys.map((key) => caches.delete(key)));
      }
    } finally {
      window.location.reload();
    }
  }

  function renderAnswerButtons(question: number, value: number, rowMode = false) {
    const order = REVERSE_ITEMS.has(question) ? [...SCALE].reverse() : SCALE;
    return (
      <div className={`answer-options ${rowMode ? "answer-options-large" : ""} ${REVERSE_ITEMS.has(question) ? "answer-options-reverse" : ""}`} role="group" aria-label={`إجابة البند ${question}`}>
        {order.map((option) => (
          <button
            type="button"
            key={option}
            className={`answer-button ${value === option ? "is-selected" : ""}`}
            aria-pressed={value === option}
            aria-label={`القيمة ${option}`}
            onClick={() => {
              const newlyAnswered = selectAnswer(question, option);
              if (rowMode && newlyAnswered && question < QUESTION_COUNT) {
                window.setTimeout(() => setCurrentQuestion((current) => Math.min(QUESTION_COUNT, current + 1)), 100);
              }
            }}
          >{option}</button>
        ))}
      </div>
    );
  }

  function renderValuesBox() {
    return (
      <section className="values-box" aria-label="القيم الحالية">
        <div className="values-heading"><span>القيم الحالية</span><span>غير المُجاب = 0 · البنود العكسية معكوسة</span></div>
        <code dir="ltr">{Array.from({ length: QUESTION_COUNT }, (_, index) => currentValues[String(index + 1)] ?? 0).join(", ")}</code>
      </section>
    );
  }

  function renderActions() {
    return (
      <div className="form-actions">
        <div className="form-action-primary">
          <button type="button" className="button button-save-new" onClick={() => commitCurrentForm(true)}>✓ حفظ وإدخال طالب جديد</button>
          <button type="button" className="button button-save" onClick={() => commitCurrentForm(false)}>💾 حفظ فقط</button>
        </div>
        <button type="button" className="button button-clear" onClick={clearCurrentForm}>مسح النموذج</button>
      </div>
    );
  }

  function renderListTab() {
    return (
      <>
        <div className="section-heading">
          <div><p className="eyebrow">نموذج قيد الإدخال</p><h2>إجابات الطالب الحالي</h2><p className="section-subtitle">اختر قيمة لكل بند. مرّر المؤشر على رقم البند لمعرفة نصّه.</p></div>
          <div className="progress-pill"><span className="progress-dot" />تمت الإجابة على <b>{answeredCount}</b> من 28</div>
        </div>
        <div className="question-grid">
          {Array.from({ length: QUESTION_COUNT }, (_, index) => {
            const question = index + 1;
            const reverse = REVERSE_ITEMS.has(question);
            return (
              <article className={`question-card ${reverse ? "is-reverse" : ""}`} key={question}>
                <button type="button" className="question-number" title={getQuestionText(question)} aria-label={`البند ${question}: ${getQuestionText(question)}`}>{question}</button>
                {renderAnswerButtons(question, currentForm.answers[String(question)] ?? 0)}
                <span className={`question-kind ${reverse ? "reverse-kind" : ""}`}>{reverse ? "عكسي" : "بند"}</span>
              </article>
            );
          })}
        </div>
        {renderValuesBox()}
        {renderActions()}
      </>
    );
  }

  function renderRowTab() {
    const response = currentValues[String(currentQuestion)] ?? 0;
    const isReverse = REVERSE_ITEMS.has(currentQuestion);
    const displayed = currentForm.answers[String(currentQuestion)] ?? 0;
    return (
      <>
        <div className="section-heading row-heading">
          <div><p className="eyebrow">إدخال متتابع</p><h2>البند {currentQuestion} من 28</h2><p className="section-subtitle" title={getQuestionText(currentQuestion)}>{getQuestionText(currentQuestion)}</p></div>
          <div className="progress-pill"><span className="progress-dot" />تمت الإجابة على <b>{answeredCount}</b> من 28</div>
        </div>
        <div className="row-table-shell" ref={rowTableRef}>
          <table className="row-entry-table" dir="ltr">
            <thead><tr><th className="row-label-cell">السؤال</th>{Array.from({ length: QUESTION_COUNT }, (_, index) => {
              const q = index + 1;
              return <th className={`${REVERSE_ITEMS.has(q) ? "reverse-header" : ""} ${q === currentQuestion ? "current-column" : ""}`} title={getQuestionText(q)} key={q}>Q{q}{REVERSE_ITEMS.has(q) ? "*" : ""}</th>;
            })}</tr></thead>
            <tbody><tr><th className="row-label-cell">الإجابة</th>{Array.from({ length: QUESTION_COUNT }, (_, index) => {
              const q = index + 1;
              return <td className={q === currentQuestion ? "current-column" : ""} key={q}>{currentForm.answers[String(q)] ? reverseScore(q, currentForm.answers[String(q)]) : 0}</td>;
            })}</tr></tbody>
          </table>
        </div>
        <div className="row-question-label"><span className={`row-question-index ${isReverse ? "is-reverse" : ""}`}>Q{currentQuestion}{isReverse ? "*" : ""}</span><span>{isReverse ? "بند عكسي" : "بند عادي"}</span><span className="row-current-answer">القيمة المحسوبة: <b>{response}</b></span></div>
        {renderAnswerButtons(currentQuestion, displayed, true)}
        <div className="row-navigation">
          <button type="button" className="button button-secondary" onClick={() => setCurrentQuestion((question) => Math.max(1, question - 1))} disabled={currentQuestion === 1}>→ السابق</button>
          <span>السؤال {currentQuestion} / 28</span>
          {currentQuestion === QUESTION_COUNT
            ? <button type="button" className="button button-save-new" onClick={() => commitCurrentForm(true)}>حفظ الطالب ✓</button>
            : <button type="button" className="button button-next" onClick={() => setCurrentQuestion((question) => Math.min(QUESTION_COUNT, question + 1))}>التالي ←</button>}
        </div>
        {renderValuesBox()}
        {renderActions()}
      </>
    );
  }

  function renderSavedTab() {
    const headers = ["رقم الطالب", ...Array.from({ length: QUESTION_COUNT }, (_, index) => `Q${index + 1}${REVERSE_ITEMS.has(index + 1) ? "*" : ""}`), "التاريخ", "حذف"];
    return (
      <>
        <div className="section-heading saved-heading">
          <div><p className="eyebrow">محفوظ على هذا الجهاز</p><h2>البيانات المحفوظة</h2><p className="section-subtitle">لا تُرسل السجلات إلى خادم؛ تبقى في مساحة التخزين المحلية لهذا المتصفح.</p></div>
          <span className="record-count">{records.length} سجل</span>
        </div>
        <div className="saved-table-shell">
          <table className="saved-table" dir="ltr">
            <thead><tr>{headers.map((header, index) => <th className={header.endsWith("*") ? "reverse-header" : ""} key={`${header}-${index}`}>{header}</th>)}</tr></thead>
            <tbody>
              {records.length === 0 ? <tr><td className="empty-table" colSpan={headers.length}><span className="empty-icon">▤</span><b>لا توجد سجلات بعد</b><span>احفظ نموذج الطالب ليظهر هنا.</span></td></tr> : records.map((record) => (
                <tr key={record.id}>
                  <td className="student-cell">{record.studentNumber}</td>
                  {Array.from({ length: QUESTION_COUNT }, (_, index) => <td key={index}>{normalizeNumber(record.answers[String(index + 1)] ?? 0)}</td>)}
                  <td className="date-cell">{record.date ? new Date(record.date).toLocaleDateString("ar") : "—"}</td>
                  <td><button type="button" className="delete-row" aria-label={`حذف سجل الطالب ${record.studentNumber}`} onClick={() => deleteRecord(record.id)}>حذف</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="data-actions">
          <button type="button" className="button button-export" onClick={exportExcel}>💾 حفظ الكل كملف Excel (.xlsx)</button>
          <div className="data-actions-secondary">
            <button type="button" className="button button-import" onClick={() => fileInputRef.current?.click()}>📥 استيراد من Excel / JSON</button>
            <button type="button" className="button button-delete-all" onClick={deleteAllRecords}>🗑️ حذف كل البيانات</button>
          </div>
          <input ref={fileInputRef} type="file" accept=".json,.xlsx,.xls,application/json,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" className="visually-hidden" onChange={handleImport} />
        </div>
      </>
    );
  }

  const tabs = [
    { id: "list" as const, icon: "▦", label: "القائمة" },
    { id: "row" as const, icon: "☷", label: "الصف" },
    { id: "saved" as const, icon: "▤", label: "البيانات المحفوظة" },
  ];

  return (
    <main className="survey-page" dir="rtl">
      <div className="survey-frame">
        <header className="app-header">
          <div className="brand-lockup">
            <div className="brand-mark" aria-hidden="true"><span>س</span><i /></div>
            <div><p className="brand-kicker">مساحة عمل محلية</p><h1>استبيان نفسي</h1></div>
          </div>
          <div className="header-actions">
            <div className="student-badge"><span className="student-label">رقم الطالب الحالي</span><strong>{nextStudentNumber}</strong><span className="student-separator" /><span className="student-caption">قيد الإدخال</span></div>
            <button type="button" className="refresh-button" onClick={refreshApp} title="مسح الكاش وتحديث التطبيق" aria-label="تحديث التطبيق">↻<span>تحديث</span></button>
          </div>
        </header>
        <div className="privacy-note"><span className="privacy-lock">▣</span><span>بياناتك محفوظة محليًا على هذا الجهاز</span><span className="privacy-dot" />{navigator.onLine ? "متصل" : "دون اتصال"}</div>
        <section className="workspace-card">
          <nav className="tab-bar" aria-label="أقسام الاستبيان" role="tablist">
            {tabs.map((tab) => <button key={tab.id} type="button" role="tab" aria-selected={activeTab === tab.id} className={`tab-button ${activeTab === tab.id ? "active" : ""}`} onClick={() => setActiveTab(tab.id)}><span className="tab-icon" aria-hidden="true">{tab.icon}</span><span>{tab.label}</span>{tab.id === "saved" && records.length > 0 && <span className="tab-count">{records.length}</span>}</button>)}
          </nav>
          <div className="tab-content" role="tabpanel">
            {activeTab === "list" && renderListTab()}
            {activeTab === "row" && renderRowTab()}
            {activeTab === "saved" && renderSavedTab()}
          </div>
        </section>
        <footer className="app-footer"><span>استبيان نفسي · إصدار محلي</span><span>مناسب للاستخدام دون اتصال بعد التحميل الأول</span></footer>
      </div>
      <div className={`toast-message ${toast ? "visible" : ""}`} role="status" aria-live="polite">{toast}</div>
    </main>
  );
}
