/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from "react";
import { 
  Upload, 
  Check, 
  FileSpreadsheet, 
  UserCheck, 
  UserX, 
  MessageSquare, 
  Clock, 
  Search, 
  Copy, 
  RotateCcw, 
  Sparkles, 
  ExternalLink, 
  FileText, 
  Phone, 
  HelpCircle, 
  CheckSquare, 
  ChevronLeft,
  Calendar,
  AlertCircle,
  Database,
  ArrowLeftRight,
  ChevronRight,
  Info,
  Sliders,
  Bell,
  Heart,
  FileDown,
  UserPlus,
  Plus,
  Settings,
  Trash2,
  CheckCircle2,
  XCircle,
  FileEdit,
  RefreshCw,
  TrendingUp,
  Layers
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";

// ==========================================
// 1. OBJECT INTERFACES & TYPES DEFINITION
// ==========================================

export interface Participant {
  id: string; // מספר משתתף
  recruitmentDate: Date | null; // תאריך גיוס
  recruitmentDateRaw: string | number; // תאריך גיוס המקורי בשביל תצוגה
  tz: string; // ת.ז.
  firstName: string; // שם פרטי
  lastName: string; // שם משפחה
  address: string; // כתובת
  phone: string; // מספר טלפון
  consent: string; // כתב הסכמה
  stool: string; // צואה
  saliva: string; // רוק
  serum: string; // סרום
  colonoscopy: string; // דגימת קולונוסקופיה
  generalQuestionnaire: string; // שאלון כללי
  dietaryQuestionnaire: string; // שאלון תזונתי
  gynecology: string; // גניקולוגיה
  researchStatus: string; // סטטוס במחקר (עמודה P)
  returningParticipant: string; // משתתף חוזר
  dietReportTyped: string; // דו"ח תזונה הוקלד
  dietReportReady: string; // דו"ח תזונה מוכן
  dietReportSent: string; // דו"ח תזונה נשלח
  bacteriaReport: string; // דו"ח חיידקים
  notes: string; // הערות
  
  // Computed fields
  monthsElapsed: number | null; // כמה חודשים עברו
  isWithdrawn: boolean; // האם פרש (עמודה P עם הערך "פרש")
  formattedPhone: string; // טלפון בפורמט בינלאומי לווטסאפ
  missingPreviousStool: boolean; // חסרה דגימה משנה שעברה (לפי ת.ז)
}

// ==========================================
// 2. HELPER UTILITIES & EXCEL PARSERS
// ==========================================

export function formatWhatsAppPhone(phoneStr: any): string {
  if (!phoneStr) return "";
  let clean = String(phoneStr).replace(/[^\d+]/g, ""); // Keep digits and +
  
  if (clean.startsWith("+")) {
    clean = clean.substring(1);
  }
  
  if (clean.startsWith("972") && clean.length >= 11) {
    return clean;
  }
  
  if (clean.startsWith("05") && clean.length === 10) {
    return "972" + clean.substring(1);
  }
  
  if (clean.startsWith("5") && clean.length === 9) {
    return "972" + clean;
  }
  
  return clean;
}

export function parseHebrewDate(val: any): Date | null {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  
  if (typeof val === "number") {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  }
  
  const str = String(val).trim();
  if (!str) return null;
  
  const dateParts = str.split(/[\/\-\.]/);
  if (dateParts.length === 3) {
    let day = parseInt(dateParts[0], 10);
    let month = parseInt(dateParts[1], 10) - 1;
    let year = parseInt(dateParts[2], 10);
    
    if (day > 100) {
      // YYYY-MM-DD
      year = day;
      month = parseInt(dateParts[1], 10) - 1;
      day = parseInt(dateParts[2], 10);
    } else if (year < 100) {
      year += year < 70 ? 2000 : 1900;
    }
    
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      const date = new Date(year, month, day);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }
  
  const parsed = Date.parse(str);
  if (!isNaN(parsed)) {
    return new Date(parsed);
  }
  
  return null;
}

export function calculateMonthsElapsed(startDate: Date | null, referenceDate: Date = new Date()): number | null {
  if (!startDate) return null;
  const start = new Date(startDate);
  const ref = new Date(referenceDate);
  if (isNaN(start.getTime()) || isNaN(ref.getTime())) return null;
  
  let months = (ref.getFullYear() - start.getFullYear()) * 12;
  months -= start.getMonth();
  months += ref.getMonth();
  
  const dayDiff = ref.getDate() - start.getDate();
  const fractionalMonth = dayDiff / 30.4375;
  const elapsed = parseFloat((months + fractionalMonth).toFixed(1));
  
  return elapsed < 0 ? 0 : elapsed;
}

export function buildHeaderMap(headersRow: any[]): Record<string, number> {
  const map: Record<string, number> = {};
  if (!headersRow || !Array.isArray(headersRow)) return map;
  
  const norm = (s: any) => String(s).replace(/[\s\.\-\'\"]+/g, "").trim().toLowerCase();
  
  const keywordsMap: Record<string, string[]> = {
    "מספר משתתף": ["מספרמשתתף", "משתתף", "קודמשתתף", "id", "participant", "מסמשתתף"],
    "תאריך גיוס": ["תאריךגיוס", "גיוס", "תאריך", "date", "recruitment", "תאריךמעקב"],
    "ת.ז.": ["תז", "ת.ז.", "תעודתזהות", "identity", "tz", "מספרזהות", "מסזהות", "מסתז", "מספרתעודתזהות", "מספרתז", "תא"],
    "שם פרטי": ["שםפרטי", "פרטי", "שם", "name", "firstname"],
    "שם משפחה": ["שםמשפחה", "משפחה", "lastname", "familyname"],
    "כתובת": ["כתובת", "ישוב", "עיר", "address"],
    "מספר טלפון": ["טלפון", "נייד", "מספרטלפון", "סלולרי", "cell", "phone"],
    "כתב הסכמה": ["כתבהסכמה", "הסכמה", "consent"],
    "צואה": ["צואה", "stool", "feces"],
    "רוק": ["רוק", "saliva"],
    "סרום": ["סרום", "serum"],
    "דגימת קולונוסקופיה": ["קולונוסקופיה", "דגימתקולונוסקופיה", "colonoscopy"],
    "שאלון כללי": ["שאלוןכללי", "כללי", "generalquestionnaire"],
    "שאלון תזונתי": ["שאלוןתזונתי", "תזונתי", "שאלוןתזונה", "dietary"],
    "סטטוס במחקר": ["סטטוסבמחקר", "סטטוס", "עמודהp", "status"],
    "משתתף חוזר": ["משתתףחוזר", "חוזר"],
    "הערות": ["הערות", "הערה", "comment", "notes"],
  };

  headersRow.forEach((val, idx) => {
    if (val === null || val === undefined) return;
    const cleanCell = norm(val);
    if (!cleanCell) return;
    
    for (const [key, aliases] of Object.entries(keywordsMap)) {
      if (aliases.some(alias => cleanCell === norm(alias))) {
        map[key] = idx;
      }
    }
  });

  headersRow.forEach((val, idx) => {
    if (val === null || val === undefined) return;
    const cleanCell = norm(val);
    if (!cleanCell) return;

    for (const [key, aliases] of Object.entries(keywordsMap)) {
      if (map[key] !== undefined) continue;

      if (aliases.some(alias => {
        const normAlias = norm(alias);
        if (normAlias.length <= 3) {
          return cleanCell === normAlias;
        } else {
          return cleanCell.includes(normAlias) || normAlias.includes(cleanCell);
        }
      })) {
        map[key] = idx;
      }
    }
  });

  return map;
}

export function mapRowToParticipant(row: any[], headerMap: Record<string, number>, referenceDate: Date): Participant {
  const getVal = (colName: string, fallbackIdx: number): string => {
    const idx = headerMap[colName];
    const actualIdx = idx !== undefined ? idx : fallbackIdx;
    const cell = row[actualIdx];
    if (cell === undefined || cell === null) return "";
    return String(cell).trim();
  };
  
  const getRawDate = (): any => {
    const idx = headerMap["תאריך גיוס"];
    const actualIdx = idx !== undefined ? idx : 1;
    return row[actualIdx];
  };

  const id = getVal("מספר משתתף", 0);
  const rawDate = getRawDate();
  const recruitmentDate = parseHebrewDate(rawDate);
  const monthsElapsed = recruitmentDate ? calculateMonthsElapsed(recruitmentDate, referenceDate) : null;
  
  let formattedRawDate = "";
  if (recruitmentDate) {
    const d = recruitmentDate.getDate().toString().padStart(2, "0");
    const m = (recruitmentDate.getMonth() + 1).toString().padStart(2, "0");
    const y = recruitmentDate.getFullYear();
    formattedRawDate = `${d}/${m}/${y}`;
  } else if (rawDate) {
    formattedRawDate = String(rawDate);
  }

  const phone = getVal("מספר טלפון", 6);
  const researchStatus = getVal("סטטוס במחקר", 15);
  const stool = getVal("צואה", 8);
  const isWithdrawn = researchStatus.includes("פרש") || researchStatus.includes("הופסק");

  return {
    id,
    recruitmentDate,
    recruitmentDateRaw: formattedRawDate,
    tz: getVal("ת.ז.", 2),
    firstName: getVal("שם פרטי", 3),
    lastName: getVal("שם משפחה", 4),
    address: getVal("כתובת", 5),
    phone,
    consent: getVal("כתב הסכמה", 7),
    stool,
    saliva: getVal("רוק", 9),
    serum: getVal("סרום", 10),
    colonoscopy: getVal("דגימת קולונוסקופיה", 11),
    generalQuestionnaire: getVal("שאלון כללי", 12),
    dietaryQuestionnaire: getVal("שאלון תזונתי", 13),
    gynecology: "",
    researchStatus,
    returningParticipant: getVal("משתתף חוזר", 16),
    dietReportTyped: "",
    dietReportReady: "",
    dietReportSent: "",
    bacteriaReport: "",
    notes: getVal("הערות", 17),
    
    monthsElapsed,
    isWithdrawn,
    formattedPhone: formatWhatsAppPhone(phone),
    missingPreviousStool: false,
  };
}

export function getSampleParticipants(refDate: Date): Participant[] {
  const rawSample = [
    {
      id: "1001",
      recruitmentDateRaw: "02/07/2025", 
      tz: "123456782",
      firstName: "עדי",
      lastName: "אשכנזי",
      address: "אהרונסון 4, תל אביב",
      phone: "0549932201",
      consent: "V",
      stool: "", 
      saliva: "V",
      serum: "V",
      colonoscopy: "",
      generalQuestionnaire: "V",
      dietaryQuestionnaire: "V",
      researchStatus: "פעיל",
      returningParticipant: "",
      notes: "דרושה פנייה למעקב 11 חודשים",
    },
    {
      id: "1002",
      recruitmentDateRaw: "20/06/2025", 
      tz: "987654321",
      firstName: "אורן",
      lastName: "לוי",
      address: "הלבנון 14, רמת גן",
      phone: "0523456789",
      consent: "V",
      stool: "", 
      saliva: "V",
      serum: "",
      colonoscopy: "",
      generalQuestionnaire: "V",
      dietaryQuestionnaire: "",
      researchStatus: "פעיל",
      returningParticipant: "",
      notes: "משתף פעולה, נא לשלוח ערכת איסוף צואה",
    },
    {
      id: "1003",
      recruitmentDateRaw: "15/06/2025", 
      tz: "334455667",
      firstName: "שירה",
      lastName: "אברהם",
      address: "העצמאות 44, גבעתיים",
      phone: "0501112222",
      consent: "V",
      stool: "", 
      saliva: "V",
      serum: "V",
      colonoscopy: "",
      generalQuestionnaire: "V",
      dietaryQuestionnaire: "V",
      researchStatus: "פעיל",
      returningParticipant: "כן",
      notes: "מתעניינת בדו\"ח חיידקי מעי חוזר",
    },
    {
      id: "1004",
      recruitmentDateRaw: "20/05/2025", 
      tz: "223344556",
      firstName: "ירון",
      lastName: "ברק",
      address: "ויצמן 80, גבעתיים",
      phone: "0539151515",
      consent: "V",
      stool: "", 
      saliva: "V",
      serum: "V",
      colonoscopy: "V",
      generalQuestionnaire: "V",
      dietaryQuestionnaire: "V",
      researchStatus: "פעיל",
      returningParticipant: "",
      notes: "כבר עבר את חלון המעקב של 11 חודשים",
    },
    {
      id: "1005",
      recruitmentDateRaw: "05/08/2025", 
      tz: "445566778",
      firstName: "רון",
      lastName: "פרנקל",
      address: "האלה 9, הרצליה",
      phone: "0559988112",
      consent: "V",
      stool: "", 
      saliva: "V",
      serum: "V",
      colonoscopy: "",
      generalQuestionnaire: "V",
      dietaryQuestionnaire: "V",
      researchStatus: "פעיל",
      returningParticipant: "",
      notes: "טרם הגיע לזמן המעקב המבוקש",
    },
    {
      id: "1006",
      recruitmentDateRaw: "25/06/2025", 
      tz: "445577881",
      firstName: "מיכל",
      lastName: "טל",
      address: "רוטשילד 11, תל אביב",
      phone: "0548844331",
      consent: "V",
      stool: "V", 
      saliva: "V",
      serum: "V",
      colonoscopy: "",
      generalQuestionnaire: "V",
      dietaryQuestionnaire: "V",
      researchStatus: "פעיל",
      returningParticipant: "",
      notes: "כבר מסרה מיוזמתה דגימה",
    },
    {
      id: "1007",
      recruitmentDateRaw: "28/06/2025", 
      tz: "998877112",
      firstName: "גיא",
      lastName: "שגב",
      address: "הבנים 10, רחובות",
      phone: "0509998887",
      consent: "V",
      stool: "", 
      saliva: "V",
      serum: "",
      colonoscopy: "",
      generalQuestionnaire: "V",
      dietaryQuestionnaire: "",
      researchStatus: "פרש - עבר דירה", 
      returningParticipant: "",
      notes: "סטטוס 'פרש' - אין לפנות שוב!",
    }
  ];

  return rawSample.map(p => {
    const recruitmentDate = parseHebrewDate(p.recruitmentDateRaw);
    const monthsElapsed = recruitmentDate ? calculateMonthsElapsed(recruitmentDate, refDate) : null;
    return {
      id: p.id,
      recruitmentDate,
      recruitmentDateRaw: p.recruitmentDateRaw,
      tz: p.tz,
      firstName: p.firstName,
      lastName: p.lastName,
      address: p.address,
      phone: p.phone,
      consent: p.consent,
      stool: p.stool,
      saliva: p.saliva,
      serum: p.serum,
      colonoscopy: p.colonoscopy,
      generalQuestionnaire: p.generalQuestionnaire,
      dietaryQuestionnaire: p.dietaryQuestionnaire,
      gynecology: "",
      researchStatus: p.researchStatus,
      returningParticipant: p.returningParticipant,
      dietReportTyped: "",
      dietReportReady: "",
      dietReportSent: "",
      bacteriaReport: "",
      notes: p.notes,
      
      monthsElapsed,
  isWithdrawn: p.researchStatus.includes("פרש") || p.researchStatus.includes("הופסק"),
      formattedPhone: formatWhatsAppPhone(p.phone),
      missingPreviousStool: false,
    };
  });
}

const DEFAULT_TEMPLATE = `מרכז המיקרוביום, שיבא

עברה כשנה מאז השתתפותך במחקר המיקרוביום כחלק מבדיקות הסקר בשיבא.
נשמח לקבל דוגמה נוספת ועדכון שאלונים (במידה ואת/ה מעוניין/ת).

- אם גם השנה אתה משתתף בבדיקות הסקר, ניתן לתאם הגעה מול הסקר, ולעדכן גם אותנו.
- אם השנה אינך משתתף בסקר, עדיין ניתן להשתתף בבדיקת המיקרוביום ללא עלות, אך יש צורך לדאוג להבאת הדוגמא לשיבא. צור עימנו קשר.

מרכז המיקרוביום שיבא,
טלפון (או ווטסאפ): 03-5304985
דוא״ל: Microbiome.Center@sheba.health.gov.il`;

const CLINICAL_TEMPLATES = [
  {
    id: "std_11m",
    name: "🩺 פנייה סטנדרטית מעקב חודשי מעמיק",
    text: DEFAULT_TEMPLATE
  },
  {
    id: "std_10m",
    name: "🩺 פנייה למעקב 10 חודשים (צואה)",
    text: DEFAULT_TEMPLATE
  },
  {
    id: "gut_only",
    name: "💩 בקשת דגימת צואה בלבד (איסוף ביתי מהיר)",
    text: `שלום {שם פרטי} היקר/ה, 🌸\nכאן צוות מחקר המיקרוביום והתזונה בשיבא תל-השומר.\n\nנשמח מאוד להזמין אותך למסור דגימת מעקב חוזרת של {חודשים שעברו} חודשים (דגימת צואה בלבד).\n\nזוהי בדיקה פשוטה וחשובה שניתן לבצע בקלות ובנוחות בבית! אנו נתאם עבורך שליח מיוחד שיביא את ערכת האיסוף וייקח אותה בחזרה ישירות מפתח ביתך ללא כל עלות. 📦\n\nהאם נוכל להזמין עבורך שליח או לתאם איסוף חוזר? נשמח אם תשיב/י לנו כאן בהודעה.\nתודה רבה מקרב לב! 🔬✨`
  },
  {
    id: "friendly_reminder",
    name: "⏳ תזכורת ידידותית קצרה וממוקדת",
    text: `היי {שם פרטי}, מה שלומך? 😊\nרק רצינו להזכיר שבקרוב יחלפו {חודשים שעברו} חודשים מאז ביקור המחקרי שלך בשיבא.\n\nנשמח מאוד לעדכן איתך פרטים לגבי ערכת המעקב הביתית הפשוטה עבור מחקר המיקרוביום.\nנשמח אם תוכל/י לכתוב לנו מתי נוח שנתקשר, או פשוט לענות לנו כאן בווטסאפ.\n\nהמשך שבוע מצויין,\nצוות מחקר המעי בשיבא 🌸`
  }
];

// ==========================================
// 3. MAIN PRESTIGE REDESIGNED APP
// ==========================================

export default function App() {
  const [rawRows, setRawRows] = useState<any[]>([]);         
  const [headers, setHeaders] = useState<any[]>([]);         
  const [participants, setParticipants] = useState<Participant[]>([]);
  const [headerMap, setHeaderMap] = useState<Record<string, number>>({});
  const [fileName, setFileName] = useState<string>("");
  const [isDragOver, setIsDragOver] = useState<boolean>(false);
  const [showAdvancedSettings, setShowAdvancedSettings] = useState<boolean>(false); 
  const [uploadMethod, setUploadMethod] = useState<"file" | "paste">("file"); 
  const [pasteText, setPasteText] = useState<string>("");
  
  const [referenceDate, setReferenceDate] = useState<string>(
    new Date("2026-06-02").toISOString().substring(0, 10)
  );

  const [exactMatchMode, setExactMatchMode] = useState<"floor" | "round">("floor"); 
  const [templateText, setTemplateText] = useState<string>(() => {
    try {
      return localStorage.getItem("sheba_whatsapp_template_v3") || CLINICAL_TEMPLATES.find(t => t.id === "std_10m")?.text || DEFAULT_TEMPLATE;
    } catch (e) {
      console.error("Error accessing localStorage", e);
      return CLINICAL_TEMPLATES.find(t => t.id === "std_10m")?.text || DEFAULT_TEMPLATE;
    }
  });
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>("std_10m");
  const [linkMode, setLinkMode] = useState<"universal" | "web">("universal"); 
  const [isEditingTemplate, setIsEditingTemplate] = useState<boolean>(false);
  const [targetFollowUpMonth, setTargetFollowUpMonth] = useState<number>(10);
  
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filterWithdrawn, setFilterWithdrawn] = useState<boolean>(true);
  const [selectedParticipantId, setSelectedParticipantId] = useState<string | null>(null);
  const [copySuccessId, setCopySuccessId] = useState<string | null>(null);
  const [autoAdvance, setAutoAdvance] = useState<boolean>(true);

  const [sentRecords, setSentRecords] = useState<Record<string, boolean>>(() => {
    try {
      const saved = localStorage.getItem("sheba_sent_records_v1");
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Sync sentRecords to localStorage
  useEffect(() => {
    try {
      localStorage.setItem("sheba_sent_records_v1", JSON.stringify(sentRecords));
    } catch (e) {
      console.error("Error setting localStorage", e);
    }
  }, [sentRecords]);

  // Tabs & Interactive workflows state
  const [activeTab, setActiveTab] = useState<"matching" | "all" | "sent">("matching"); 
  const [quickFilterSample, setQuickFilterSample] = useState<"all" | "missing_stool" | "missing_prev" | "returning" | "has_notes" | "withdrawn">("all");
  const [showManualForm, setShowManualForm] = useState<boolean>(false);
  const [saveSuccess, setSaveSuccess] = useState<boolean>(false);
  const [editingNotes, setEditingNotes] = useState<string>("");
  const tableRowRef = useRef<HTMLTableRowElement>(null);

  useEffect(() => {
    if (selectedParticipantId && tableRowRef.current) {
      tableRowRef.current.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
  }, [selectedParticipantId]);


  // Manual form inputs state
  const [manualId, setManualId] = useState<string>("");
  const [manualFirstName, setManualFirstName] = useState<string>("");
  const [manualLastName, setManualLastName] = useState<string>("");
  const [manualTz, setManualTz] = useState<string>("");
  const [manualPhone, setManualPhone] = useState<string>("");
  const [manualRecruitmentDate, setManualRecruitmentDate] = useState<string>("2025-07-02");
  const [manualStool, setManualStool] = useState<string>(""); 
  const [manualSaliva, setManualSaliva] = useState<string>("V");
  const [manualSerum, setManualSerum] = useState<string>("V");
  const [manualNotes, setManualNotes] = useState<string>("דגימות מעקב נרשמו ידנית בקליניקה");

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Store configuration templates on modify
  useEffect(() => {
    try {
      localStorage.setItem("sheba_whatsapp_template_v3", templateText);
    } catch (e) {
      console.error("Error setting localStorage", e);
    }
  }, [templateText]);

  // Recalculate participant parameters on referenceDate / targetFollowUpMonth / exactMatchMode / rawRows change
  useEffect(() => {
    if (rawRows.length > 0) {
      const baseRefDate = new Date(referenceDate);
      const mappedList: Participant[] = [];

      for (let i = 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        if (!row || row.length === 0 || row.every(cell => cell === null || cell === undefined || cell === "")) {
          continue;
        }
        const p = mapRowToParticipant(row, headerMap, baseRefDate);
        mappedList.push(p);
      }

      // Detect missing previous stool samples by TZ
      const participantsWithMissingStool = new Set<string>();
      mappedList.forEach(p => {
        if (!p.tz) return;
        const hasStool = p.stool && (p.stool.toUpperCase() === "V" || p.stool.includes("כן") || p.stool.includes("✓") || p.stool.includes("הוגשה"));
        // If they have a record where stool is missing AND they're NOT the current year (monthsElapsed > 1.5 roughly)
        if (!hasStool && p.monthsElapsed !== null && p.monthsElapsed > 6) {
          participantsWithMissingStool.add(p.tz);
        }
      });

      const enrichedList = mappedList.map(p => ({
        ...p,
        missingPreviousStool: p.tz ? participantsWithMissingStool.has(p.tz) : false
      }));

      setParticipants(enrichedList);

      // Keep selection or auto-pick first eligible candidate
      if (mappedList.length > 0 && !mappedList.some(item => item.id === selectedParticipantId)) {
        const candidates = mappedList.filter(p => {
          if (p.monthsElapsed === null) return false;
          const matchMonth = exactMatchMode === "floor" 
            ? Math.floor(p.monthsElapsed) === targetFollowUpMonth 
            : Math.round(p.monthsElapsed) === targetFollowUpMonth;
          return matchMonth && (!filterWithdrawn || !p.isWithdrawn);
        });
        if (candidates.length > 0) {
          setSelectedParticipantId(candidates[0].id);
        } else {
          setSelectedParticipantId(mappedList[0].id);
        }
      }
    }
  }, [referenceDate, headerMap, rawRows, exactMatchMode, filterWithdrawn, targetFollowUpMonth]);

  const handlePasteImport = (pastedText: string) => {
    if (!pastedText || !pastedText.trim()) {
      alert("אנא הדבק נתונים בתיבת הטקסט.");
      return;
    }
    try {
      const lines = pastedText.split(/\r?\n/).filter(line => line.trim() !== "");
      const rawGrid = lines
        .map(line => line.split('\t').map(cell => cell.trim()))
        .filter(row => row.some(cell => cell !== "")); 

      if (rawGrid.length < 2) {
        alert("הנתונים שהודבקו אינם מכילים מספיק שורות או עמודות. אנא וודא שהעתקת טבלה תקינה מאקסל (כולל שורת כותרות).");
        return;
      }

      setFileName("נתונים שהודבקו מאקסל.tsv");
      
      const headerRowParsed = rawGrid[0] || [];
      setHeaders(headerRowParsed);
      setRawRows(rawGrid);

      const autoMapResult = buildHeaderMap(headerRowParsed);
      setHeaderMap(autoMapResult);
      setPasteText("");
    } catch (err) {
      console.error(err);
      alert("שגיאה בפענוח הנתונים שהודבקו. ודא שהעתקת טבלה תקינה מתוך אקסל (כולל כותרות) ונסה שוב.");
    }
  };

  const handleFileImport = (file: File) => {
    if (!file) return;
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const rawGrid = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });

        if (rawGrid.length < 2) {
          alert("קובץ שגוי או אינו מכיל נתוני משתתפים.");
          return;
        }

        const headerRowParsed = rawGrid[0] || [];
        setHeaders(headerRowParsed);
        setRawRows(rawGrid);

        const autoMapResult = buildHeaderMap(headerRowParsed);
        setHeaderMap(autoMapResult);

      } catch (err) {
        console.error(err);
        alert("שגיאה בפענוח קובץ האקסל. וודא שהוא תקין, מכיל עמודת תאריכי רישום, ונסה שוב.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileUploadChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      handleFileImport(files[0]);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      handleFileImport(files[0]);
    }
  };

  const handleLoadSample = () => {
    const baseRefDate = new Date(referenceDate);
    const samples = getSampleParticipants(baseRefDate);
    
    const mockHeaders = [
      "מספר משתתף", "תאריך גיוס", "ת.ז.", "שם פרטי", "שם משפחה", 
      "כתובת", "מספר טלפון", "כתב הסכמה", "צואה", "רוק", "סרום", 
      "דגימת קולונוסקופיה", "שאלון כללי", "שאלון תזונתי", "סטטוס במחקר", 
      "משתתף חוזר", "הערות"
    ];
    
    setHeaders(mockHeaders);
    const mockGridRows: any[][] = [mockHeaders];
    
    samples.forEach((p) => {
      const row = mockHeaders.map((hdr) => {
        switch(hdr) {
          case "מספר משתתף": return p.id;
          case "תאריך גיוס": return p.recruitmentDateRaw;
          case "ת.ז.": return p.tz;
          case "שם פרטי": return p.firstName;
          case "שם משפחה": return p.lastName;
          case "כתובת": return p.address;
          case "מספר טלפון": return p.phone;
          case "כתב הסכמה": return p.consent;
          case "צואה": return p.stool;
          case "רוק": return p.saliva;
          case "סרום": return p.serum;
          case "דגימת קולונוסקופיה": return p.colonoscopy;
          case "שאלון כללי": return p.generalQuestionnaire;
          case "שאלון תזונתי": return p.dietaryQuestionnaire;
          case "סטטוס במחקר": return p.researchStatus;
          case "משתתף חוזר": return p.returningParticipant;
          case "הערות": return p.notes;
          default: return "";
        }
      });
      mockGridRows.push(row);
    });

    setRawRows(mockGridRows);

    const defaultMapping: Record<string, number> = {};
    mockHeaders.forEach((hdr, idx) => {
      defaultMapping[hdr] = idx;
    });
    setHeaderMap(defaultMapping);
    setFileName("קובץ_גיוס_דוגמה_שיבא.xlsx");
    setSelectedParticipantId("1001");
  };

  const clearLoadedData = () => {
    setRawRows([]);
    setHeaders([]);
    setParticipants([]);
    setHeaderMap({});
    setFileName("");
    setSelectedParticipantId(null);
  };

  const replacePlaceholders = (template: string, p: Participant) => {
    if (!p) return "";
    const monthsStr = p.monthsElapsed !== null ? Math.floor(p.monthsElapsed).toString() : targetFollowUpMonth.toString();
    const cleanFirstName = p.firstName || "משתתף/ת";
    const cleanLastName = p.lastName || "";
    const dateStr = p.recruitmentDateRaw ? String(p.recruitmentDateRaw) : "מועד הגיוס";
    const idStr = p.id || "";
    const tzStr = p.tz || "";

    return template
      .replace(/{שם פרטי}/g, cleanFirstName)
      .replace(/{שם משפחה}/g, cleanLastName)
      .replace(/{תאריך גיוס}/g, dateStr)
      .replace(/{חודשים שערו}/g, monthsStr) 
      .replace(/{חודשים שעברו}/g, monthsStr)
      .replace(/{מספר משתתף}/g, idStr)
      .replace(/{ת\.ז\.}/g, tzStr);
  };

  const getWhatsAppUrl = (p: Participant) => {
    const rawMsg = replacePlaceholders(templateText, p);
    const encoded = encodeURIComponent(rawMsg);
    const cleanPhone = p.formattedPhone;

    if (linkMode === "web") {
      return `https://web.whatsapp.com/send?phone=${cleanPhone}&text=${encoded}`;
    }
    return `https://wa.me/${cleanPhone}?text=${encoded}`;
  };

  const copyToClipboard = (text: string, pId: string) => {
    navigator.clipboard.writeText(text).then(() => {
      setCopySuccessId(pId);
      setTimeout(() => setCopySuccessId(null), 2550);
    });
  };

  const addPlaceholderToken = (token: string) => {
    if (token) {
      setTemplateText(prev => prev + ` {${token}}`);
    }
  };


  const updateMapping = (key: string, headerIndexValue: string) => {
    setHeaderMap(prev => ({
      ...prev,
      [key]: headerIndexValue === "" ? undefined : parseInt(headerIndexValue, 10)
    }));
  };

  const handleSaveNotes = (pId: string, notesContent: string) => {
    const idColIdx = headerMap["מספר משתתף"];
    const notesColIdx = headerMap["הערות"];
    
    if (idColIdx === undefined || notesColIdx === undefined) {
      alert("לא ניתן לשמור הערה: ודא שעמודות מספר משתתף והערות ממופות כראוי בהגדרות המתקדמות.");
      return;
    }

    const gridCpy = [...rawRows];
    let found = false;
    for (let i = 1; i < gridCpy.length; i++) {
      const row = gridCpy[i];
      if (row && String(row[idColIdx]).trim() === String(pId).trim()) {
        row[notesColIdx] = notesContent;
        found = true;
        break;
      }
    }

    if (found) {
      setRawRows(gridCpy);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 2000);
    } else {
      alert("שגיאה: המטופל לא נמצא במאגר הרשומות.");
    }
  };

  const handleSelectTemplate = (templateId: string) => {
    setSelectedTemplateId(templateId);
    const t = CLINICAL_TEMPLATES.find(c => c.id === templateId);
    if (t) {
      setTemplateText(t.text);
    }
  };

  const handleAddManualParticipant = (e: React.FormEvent) => {
    e.preventDefault();
    if (!manualFirstName || !manualLastName || !manualPhone) {
      alert("אנא מלא את שדות החובה: שם פרטי, שם משפחה ומספר טלפון.");
      return;
    }

    let currentHeaders = [...headers];
    let mappings = { ...headerMap };

    if (currentHeaders.length === 0) {
      currentHeaders = [
        "מספר משתתף", "תאריך גיוס", "ת.ז.", "שם פרטי", "שם משפחה", 
        "כתובת", "מספר טלפון", "כתב הסכמה", "צואה", "רוק", "סרום", 
        "דגימת קולונוסקופיה", "שאלון כללי", "שאלון תזונתי", "סטטוס במחקר", 
        "משתתף חוזר", "הערות"
      ];
      setHeaders(currentHeaders);
      currentHeaders.forEach((hdr, idx) => {
        mappings[hdr] = idx;
      });
      setHeaderMap(mappings);
    }

    let formattedRecDate = "02/07/2025";
    if (manualRecruitmentDate) {
      const dateParts = manualRecruitmentDate.split("-");
      if (dateParts.length === 3) {
        formattedRecDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
      }
    }

    const generatedId = manualId.trim() || String(Date.now()).slice(-4);

    const newRow = currentHeaders.map((hdr) => {
      switch(hdr) {
        case "מספר משתתף": return generatedId;
        case "תאריך גיוס": return formattedRecDate;
        case "ת.ז.": return manualTz.trim();
        case "שם פרטי": return manualFirstName.trim();
        case "שם משפחה": return manualLastName.trim();
        case "כתובת": return "גיוס ידני בקליניקה";
        case "מספר טלפון": return manualPhone.trim();
        case "כתב הסכמה": return "V";
        case "צואה": return manualStool;
        case "רוק": return manualSaliva;
        case "סרום": return manualSerum;
        case "דגימת קולונוסקופיה": return "";
        case "שאלון כללי": return "V";
        case "שאלון תזונתי": return "V";
        case "סטטוס במחקר": return "פעיל";
        case "משתתף חוזר": return "";
        case "הערות": return manualNotes.trim();
        default: return "";
      }
    });

    let newGrid: any[][] = [];
    if (rawRows.length === 0) {
      newGrid = [currentHeaders, newRow];
      setFileName("משתתפים_ייבוא_ידני.xlsx");
    } else {
      newGrid = [...rawRows, newRow];
    }

    setRawRows(newGrid);
    setSelectedParticipantId(generatedId);
    setShowManualForm(false);

    // Clear manual states
    setManualId("");
    setManualFirstName("");
    setManualLastName("");
    setManualTz("");
    setManualPhone("");
    setManualNotes("דגימות מעקב נרשמו ידנית בקליניקה");
    setManualStool("");
  };

  const handleExportToExcel = () => {
    if (rawRows.length < 2) {
      alert("אין נתונים לייצוא כרגע.");
      return;
    }
    try {
      const worksheet = XLSX.utils.aoa_to_sheet(rawRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "מעקב משתתפים מעודכן");
      XLSX.writeFile(workbook, `שיבא_מעקב_משתתפים_מעודכן_${referenceDate}.xlsx`);
    } catch(err) {
      console.error(err);
      alert("שגיאה בייצוא קובץ אקסל.");
    }
  };

  useEffect(() => {
    const selectedP = participants.find(p => p.id === selectedParticipantId);
    if (selectedP) {
      setEditingNotes(selectedP.notes || "");
    } else {
      setEditingNotes("");
    }
  }, [selectedParticipantId, participants]);

  const getTabParticipants = () => {
    let list: Participant[] = [];
    if (activeTab === "matching") {
      list = participants.filter(p => {
        if (p.monthsElapsed === null) return false;
        const isExactlyTargetMonth = exactMatchMode === "floor"
          ? Math.floor(p.monthsElapsed) === targetFollowUpMonth
          : Math.round(p.monthsElapsed) === targetFollowUpMonth;
        return isExactlyTargetMonth;
      });
    } else if (activeTab === "sent") {
      list = participants.filter(p => !!sentRecords[p.id]);
    } else {
      list = participants;
    }

    if (filterWithdrawn && quickFilterSample !== "withdrawn") {
      list = list.filter(p => !p.isWithdrawn);
    }

    if (searchQuery.trim() !== "") {
      const q = searchQuery.toLowerCase();
      list = list.filter(p => {
        const nameStr = `${p.firstName} ${p.lastName}`.toLowerCase();
        return nameStr.includes(q) || 
               String(p.id).toLowerCase().includes(q) || 
               String(p.phone).includes(q) || 
               String(p.tz).includes(q) ||
               String(p.notes).toLowerCase().includes(q);
      });
    }

    if (quickFilterSample === "missing_stool") {
      list = list.filter(p => {
        const hs = p.stool && (p.stool.toUpperCase() === "V" || p.stool.includes("כן") || p.stool.includes("✓") || p.stool.includes("הוגשה"));
        return !hs;
      });
    } else if (quickFilterSample === "missing_prev") {
      list = list.filter(p => p.missingPreviousStool);
    } else if (quickFilterSample === "returning") {
      list = list.filter(p => p.returningParticipant && p.returningParticipant.trim() !== "");
    } else if (quickFilterSample === "has_notes") {
      list = list.filter(p => p.notes && p.notes.trim() !== "");
    } else if (quickFilterSample === "withdrawn") {
      list = list.filter(p => p.isWithdrawn);
    }

    return list;
  };

  const activeList = getTabParticipants();
  const filteredParticipants = activeList;

  const selectNextCandidate = () => {
    const currentIndex = filteredParticipants.findIndex(p => p.id === selectedParticipantId);
    if (currentIndex !== -1 && currentIndex < filteredParticipants.length - 1) {
      setSelectedParticipantId(filteredParticipants[currentIndex + 1].id);
      return filteredParticipants[currentIndex + 1].id;
    }
    return null;
  };

  const selectPrevCandidate = () => {
    const currentIndex = filteredParticipants.findIndex(p => p.id === selectedParticipantId);
    if (currentIndex > 0) {
      setSelectedParticipantId(filteredParticipants[currentIndex - 1].id);
      return filteredParticipants[currentIndex - 1].id;
    }
    return null;
  };

  // Keyboard navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.target instanceof HTMLTextAreaElement || e.target instanceof HTMLInputElement) return;

      if (e.key === "ArrowLeft") {
        selectNextCandidate();
      } else if (e.key === "ArrowRight") {
        selectPrevCandidate();
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [selectedParticipantId, filteredParticipants]);

  const hasDateMappingIssue = participants.length > 0 && participants.filter(p => p.recruitmentDate !== null).length === 0;

  // Stats calculation
  const totalLoaded = participants.length;
  const totalSent = participants.filter(p => !!sentRecords[p.id]).length;

  const markAsSent = (id: string) => {
    setSentRecords(prev => ({
      ...prev,
      [id]: true
    }));

    if (autoAdvance) {
      setTimeout(() => {
        selectNextCandidate();
      }, 500);
    }
  };

  const removeSentMark = (id: string) => {
    setSentRecords(prev => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
  };

  const resetAllSentTags = () => {
    if (window.confirm("רוצה לאפס את כל רשומות השליחה השמורות בדפדפן?")) {
      setSentRecords({});
      try {
        localStorage.removeItem("sheba_sent_records_v1");
      } catch (e) {
        console.error("Error clearing localStorage", e);
      }
    }
  };

  const activeTargetCount = participants.filter(p => {
    if (p.monthsElapsed === null) return false;
    const isExactlyTarget = exactMatchMode === "floor"
      ? Math.floor(p.monthsElapsed) === targetFollowUpMonth
      : Math.round(p.monthsElapsed) === targetFollowUpMonth;
    return isExactlyTarget && (!filterWithdrawn || !p.isWithdrawn);
  }).length;

  const missingStoolTargetCount = participants.filter(p => {
    const isTarget = exactMatchMode === "floor" 
      ? Math.floor(p.monthsElapsed || 0) === targetFollowUpMonth 
      : Math.round(p.monthsElapsed || 0) === targetFollowUpMonth;
    const hasStool = p.stool && (p.stool.toUpperCase() === "V" || p.stool.includes("כן") || p.stool.includes("✓"));
    return isTarget && !hasStool && (!filterWithdrawn || !p.isWithdrawn);
  }).length;

  const selectedParticipant = filteredParticipants.find(p => p.id === selectedParticipantId) || filteredParticipants[0] || null;

  return (
    <div className="min-h-screen lg:h-screen w-full flex flex-col lg:overflow-hidden bg-slate-50 text-slate-800 font-sans selection:bg-teal-100 selection:text-teal-950" dir="rtl">
      
      {/* BRAND NEW LUXURIOUS HEADER */}
      <header className="h-auto py-3 lg:py-0 lg:h-14 bg-[#0E2232] text-white flex flex-col lg:flex-row items-center justify-between px-6 shrink-0 z-30 shadow-lg border-b border-white/5 gap-3 lg:gap-0">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-teal-500 rounded-xl flex items-center justify-center text-white font-extrabold text-xs shadow-md tracking-wider">
            שיבא
          </div>
          <div>
            <h1 className="text-base font-bold tracking-tight flex items-center gap-2">
              <span>פורטל מרכז המיקרוביום והתזונה</span>
              <span className="text-[10px] bg-teal-500/10 text-teal-300 px-2 py-0.5 rounded-full border border-teal-500/20 font-mono tracking-wider">
                ערוץ תקשורת מעקב {targetFollowUpMonth} חודשים
              </span>
            </h1>
            <p className="text-[10px] text-slate-400 font-medium">ניהול ממוחשב, סינון עיוור של פציאנטים ופניות ווטסאפ שבועיות | Sheba Medical Center</p>
          </div>
        </div>

        {/* Global Controls */}
        <div className="flex items-center gap-3">
          
          {/* Target Month */}
          <div className="flex items-center gap-2 bg-slate-800/60 px-3 py-1.5 rounded-lg text-xs border border-white/10">
            <Clock className="w-3.5 h-3.5 text-teal-400" />
            <span className="text-slate-300 font-medium">חודש יעד למעקב:</span>
            <select 
              value={targetFollowUpMonth}
              onChange={(e) => setTargetFollowUpMonth(parseInt(e.target.value, 10))}
              className="bg-transparent text-teal-300 font-bold outline-none text-xs cursor-pointer border-0 p-0 focus:ring-0"
            >
              {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24].map((m) => (
                <option key={m} value={m} className="text-slate-900 bg-white">חודש {m}</option>
              ))}
            </select>
          </div>

          {/* Clinician Reference Date key */}
          <div className="flex items-center gap-2 bg-slate-800/60 px-3 py-1.5 rounded-lg text-xs border border-white/10">
            <Calendar className="w-3.5 h-3.5 text-teal-400" />
            <span className="text-slate-300 font-medium font-sans">תאריך ציון קליני:</span>
            <input 
              type="date"
              value={referenceDate}
              onChange={(e) => setReferenceDate(e.target.value)}
              className="bg-transparent text-teal-300 font-bold outline-none font-mono text-xs w-[110px] cursor-pointer border-0 p-0 focus:ring-0"
            />
          </div>

          <button 
            onClick={handleExportToExcel}
            disabled={rawRows.length < 2}
            className="bg-teal-600 hover:bg-teal-500 disabled:opacity-30 px-4 py-2 rounded-lg text-xs font-bold text-white shadow-md transition-all flex items-center gap-2 cursor-pointer border-0"
          >
            <FileDown className="w-4 h-4" />
            ייצא אקסל מעודכן (.xlsx)
          </button>
        </div>
      </header>

      {/* THREE GLORIOUS HIGH-CONTRAST METRICS CORES */}
      <section className="bg-white border-b border-slate-200 px-6 py-2 shrink-0 grid grid-cols-1 md:grid-cols-3 gap-4 shadow-[0_1px_3px_rgba(0,0,0,0.02)]">
        
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 flex items-center justify-between hover:shadow-sm transition-all">
          <div className="space-y-0.5">
            <span className="block text-[9px] text-slate-550 font-black tracking-wide uppercase">סה"כ משתתפים שנטענו</span>
            <p className="text-xl font-mono font-black text-slate-900">{totalLoaded}</p>
          </div>
          <div className="p-2 bg-slate-100 rounded-lg text-slate-500 shrink-0">
            <Layers className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-teal-50/50 border border-teal-100 rounded-xl p-3 flex items-center justify-between hover:shadow-sm transition-all">
          <div className="space-y-0.5">
            <span className="block text-[9px] text-teal-700 font-black tracking-wide uppercase">זכאים למעקב חודש {targetFollowUpMonth}</span>
            <p className="text-xl font-mono font-black text-teal-900">{activeTargetCount}</p>
          </div>
          <div className="p-2 bg-teal-100/50 rounded-lg text-teal-700 shrink-0">
            <UserCheck className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-amber-55/30 border border-amber-150/50 rounded-xl p-3 flex items-center justify-between hover:shadow-sm transition-all">
          <div className="space-y-0.5">
            <span className="block text-[9px] text-amber-800 font-black tracking-wide uppercase">צואה חסרה {targetFollowUpMonth}ח׳</span>
            <p className="text-xl font-mono font-black text-amber-900">{missingStoolTargetCount}</p>
          </div>
          <div className="p-2 bg-amber-100 text-amber-700 shrink-0 animate-pulse rounded-lg">
            <AlertCircle className="w-5 h-5" />
          </div>
        </div>

      </section>

      {/* CORE FRAME LAYOUT */}
      <div className="flex-1 flex flex-col lg:flex-row lg:overflow-hidden">
        
        {/* RIGHT CONTROL PANEL - SIZED AND DESIGNED PERFECTLY */}
        <aside className="w-full lg:w-80 bg-white border-l border-slate-250 flex flex-col p-4 overflow-y-auto shrink-0 space-y-5">
          
          {/* Module 1: Clean Data import */}
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3 shadow-3xs">
            <div className="flex items-center justify-between pb-1 border-b border-slate-200">
              <h3 className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <FileSpreadsheet className="w-4 h-4 text-teal-600" />
                ייבוא וטעינת נתונים
              </h3>
              {fileName && (
                <button 
                  onClick={clearLoadedData}
                  className="text-[10px] text-red-600 hover:text-red-700 font-bold transition-colors cursor-pointer"
                >
                  נקה הכל
                </button>
              )}
            </div>

            {/* Upload method toggles */}
            <div className="flex gap-1.5 bg-slate-200/60 p-1 rounded-lg text-[10px] font-bold">
              <button 
                onClick={() => setUploadMethod("file")}
                className={`flex-1 py-1 px-2 rounded-md transition-all cursor-pointer ${uploadMethod === "file" ? "bg-white shadow-sm text-slate-800" : "text-slate-500 hover:text-slate-800"}`}
              >
                📁 קובץ Excel
              </button>
              <button 
                onClick={() => setUploadMethod("paste")}
                className={`flex-1 py-1 px-2 rounded-md transition-all cursor-pointer ${uploadMethod === "paste" ? "bg-white shadow-sm text-slate-800" : "text-slate-500 hover:text-slate-800"}`}
              >
                📋 הדבק נתונים
              </button>
            </div>

            {uploadMethod === "file" ? (
              <div className="space-y-2">
                <div 
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  onClick={() => fileInputRef.current?.click()}
                  className={`border-2 border-dashed rounded-xl p-4 text-center cursor-pointer transition-all ${
                    isDragOver ? "border-teal-500 bg-teal-50/20" : "border-slate-300 hover:border-teal-400 bg-white"
                  }`}
                >
                  <input 
                    type="file" 
                    ref={fileInputRef}
                    accept=".xlsx,.xls,.csv" 
                    onChange={handleFileUploadChange}
                    className="hidden" 
                  />
                  <div className="space-y-1.5 block">
                    <Upload className="w-7 h-7 mx-auto text-slate-400" />
                    <p className="text-xs font-bold text-slate-700 leading-tight">שחרר קובץ אקסל או לחץ לבחירה</p>
                    <p className="text-[9px] text-slate-400">קובצי .xlsx, .xls או מעקב שבועי</p>
                  </div>
                </div>

                {!fileName && (
                  <button 
                    onClick={handleLoadSample}
                    className="w-full bg-teal-55/10 text-teal-800 hover:bg-teal-55/20 border border-teal-200 text-[10.5px] font-bold py-2 rounded-lg transition-all cursor-pointer flex items-center justify-center gap-1.5"
                  >
                    <Sparkles className="w-3.5 h-3.5 text-teal-600 animate-pulse" />
                    טען בסיס נתונים רפואי לדוגמה
                  </button>
                )}
              </div>
            ) : (
              <div className="space-y-2">
                <textarea 
                  rows={4}
                  placeholder="העתק טבלה מתוך קובץ האקסל והדבק אותה כאן כולל שורת הכותרות..."
                  value={pasteText}
                  onChange={(e) => setPasteText(e.target.value)}
                  className="w-full bg-white border border-slate-250 rounded-lg p-2 font-mono text-[9px] focus:ring-1 focus:ring-teal-500 focus:border-teal-500 outline-none text-left"
                  dir="ltr"
                />
                <button 
                  onClick={() => handlePasteImport(pasteText)}
                  className="w-full bg-teal-600 hover:bg-teal-700 text-white font-bold py-1.5 rounded-lg text-xs transition-all cursor-pointer text-center"
                >
                  ⚡ מייבא נתוני הדבק
                </button>
              </div>
            )}
          </div>

          {/* Module 2: Fast clinician add */}
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 shadow-3xs space-y-2.5">
            <button 
              type="button"
              onClick={() => setShowManualForm(!showManualForm)}
              className="w-full flex items-center justify-between text-xs font-bold text-slate-800 cursor-pointer"
            >
              <span className="flex items-center gap-1.5">
                <UserPlus className="w-4 h-4 text-teal-600" />
                הוספת פציינט קליני מהירה
              </span>
              <span className="text-[10px] text-slate-400">{showManualForm ? "▲" : "▼"}</span>
            </button>

            <AnimatePresence>
              {showManualForm && (
                <motion.form 
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: "auto" }}
                  exit={{ opacity: 0, height: 0 }}
                  onSubmit={handleAddManualParticipant} 
                  className="space-y-3 pt-2.5 border-t border-slate-200 text-xs overflow-hidden"
                >
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[9px] text-slate-500 font-bold mb-0.5">מזהה פנימי:</label>
                      <input 
                        type="text" 
                        placeholder="ID" 
                        value={manualId}
                        onChange={(e) => setManualId(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded p-1.5 font-mono text-[10px]"
                      />
                    </div>
                    <div>
                      <label className="block text-[9px] text-slate-500 font-bold mb-0.5">תעודת זהות:</label>
                      <input 
                        type="text" 
                        placeholder="ת.ז." 
                        value={manualTz}
                        onChange={(e) => setManualTz(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded p-1.5 font-mono text-[10px]"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[9px] text-slate-500 font-bold mb-0.5">שם פרטי *:</label>
                      <input 
                        type="text" 
                        required
                        placeholder="שם משתתף" 
                        value={manualFirstName}
                        onChange={(e) => setManualFirstName(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded p-1.5 text-[10px] font-bold"
                      />
                    </div>
                    <div>
                      <label className="block text-[9px] text-slate-500 font-bold mb-0.5">שם משפחה *:</label>
                      <input 
                        type="text" 
                        required
                        placeholder="משפחה" 
                        value={manualLastName}
                        onChange={(e) => setManualLastName(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded p-1.5 text-[10px] font-bold"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-[9px] text-slate-500 font-bold mb-0.5">נייד משתתף *:</label>
                    <input 
                      type="text" 
                      required
                      placeholder="למשל: 0501234567" 
                      value={manualPhone}
                      onChange={(e) => setManualPhone(e.target.value)}
                      className="w-full bg-white border border-slate-300 rounded p-1.5 font-mono text-[10px] font-bold"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[9px] text-slate-500 font-bold mb-0.5">תאריך גיוס:</label>
                      <input 
                        type="date" 
                        value={manualRecruitmentDate}
                        onChange={(e) => setManualRecruitmentDate(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded p-1 font-mono text-[10px] cursor-pointer"
                      />
                    </div>
                    <div>
                      <label className="block text-[9px] text-slate-500 font-bold mb-0.5">דגימת צואה:</label>
                      <select 
                        value={manualStool} 
                        onChange={(e) => setManualStool(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded p-1 text-[10px] cursor-pointer"
                      >
                        <option value="">חסרה ⚠️</option>
                        <option value="V">הוגשה (V) ✓</option>
                      </select>
                    </div>
                  </div>

                  <button 
                    type="submit"
                    className="w-full bg-[#128C7E] hover:bg-[#075E54] hover:shadow text-white font-bold py-2 rounded-lg text-xs transition-colors cursor-pointer text-center"
                  >
                    💾 שמור והוסף לפאנל
                  </button>
                </motion.form>
              )}
            </AnimatePresence>
          </div>

          {/* Module 3: Exact threshold filters */}
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 shadow-3xs space-y-3">
            <span className="text-[10px] font-black text-slate-500 block uppercase tracking-wide">זמן חישוב ועיגולים</span>
            
            <div className="space-y-2">
              <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-700">
                <input 
                  type="radio" 
                  name="exactMode" 
                  value="floor"
                  checked={exactMatchMode === "floor"}
                  onChange={() => setExactMatchMode("floor")}
                  className="text-teal-600 focus:ring-teal-500 cursor-pointer w-4 h-4"
                />
                עיגול כלפי מטה (Floor)
              </label>

              <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-700">
                <input 
                  type="radio" 
                  name="exactMode" 
                  value="round"
                  checked={exactMatchMode === "round"}
                  onChange={() => setExactMatchMode("round")}
                  className="text-teal-600 focus:ring-teal-500 cursor-pointer w-4 h-4"
                />
                עיגול מתמטי קרוב (Round)
              </label>
            </div>

            <p className="text-[9.5px] text-slate-400 block font-semibold leading-relaxed">
              {exactMatchMode === "floor" 
                ? `מועמדים שהשלימו לפחות ${targetFollowUpMonth}.0 חודשים (טווח: ${targetFollowUpMonth}.0 עד ${targetFollowUpMonth}.9)` 
                : `מועמדים שהשלימו בין ${targetFollowUpMonth - 0.5} ל-${targetFollowUpMonth + 0.4} חודשים`}
            </p>
          </div>

          {/* Module 4: Manual mappings */}
          {headers.length > 0 && (
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 shadow-3xs space-y-3">
              <span className="text-[10px] font-black text-slate-500 block">זיהוי עמודות באקסל קיימות</span>
              <div className="space-y-2 max-h-[140px] overflow-y-auto pr-0.5">
                {[
                  { key: "שם פרטי", label: "שם פרטי" },
                  { key: "שם משפחה", label: "משפחה" },
                  { key: "ת.ז.", label: "תעודת זהות" },
                  { key: "מספר משתתף", label: "קוד תיק" },
                  { key: "מספר טלפון", label: "טלפון מעודכן" },
                  { key: "תאריך גיוס", label: "תאריך גיוס" },
                  { key: "צואה", label: "דגימת צואה" }
                ].map((item) => {
                  const value = headerMap[item.key] !== undefined ? headerMap[item.key] : "";
                  return (
                    <div key={item.key} className="flex items-center justify-between gap-1 text-[10px]">
                      <span className="text-slate-600 truncate max-w-[90px] font-medium">{item.label}:</span>
                      <select 
                        value={value}
                        onChange={(e) => updateMapping(item.key, e.target.value)}
                        className="bg-white border border-slate-250 rounded text-[9.5px] px-1 py-1 cursor-pointer outline-none max-w-[120px]"
                      >
                        <option value="">זיהוי אוטומטי</option>
                        {headers.map((hdr, idx) => (
                          <option key={idx} value={idx}>
                            {idx + 1}: {hdr || `עמודה ללא שם`}
                          </option>
                        ))}
                      </select>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Module 5: Activity indices */}
          <div className="bg-slate-900 text-white rounded-xl p-4 space-y-3 shadow-md mt-auto shrink-0 border border-slate-800">
            <span className="text-[10px] font-bold text-teal-400 block uppercase tracking-wider">מדדי פעילות מעקב נוכחיים</span>
            
            <div className="grid grid-cols-3 gap-2 text-center text-xs">
              <div className="bg-slate-800/80 p-2 rounded-lg border border-slate-700/50">
                <span className="text-[9px] text-slate-400 block font-bold leading-none">תיקים</span>
                <span className="text-sm font-black text-white font-mono block mt-1">{totalLoaded}</span>
              </div>
              <div className="bg-slate-800/80 p-2 rounded-lg border border-slate-700/50">
                <span className="text-[9px] text-slate-400 block font-bold leading-none">מיועדים</span>
                <span className="text-sm font-black text-teal-400 font-mono block mt-1">{filteredParticipants.length}</span>
              </div>
              <div className="bg-slate-800/80 p-2 rounded-lg border border-slate-700/50">
                <span className="text-[9px] text-slate-400 block font-bold leading-none">נשלחו</span>
                <span className="text-sm font-black text-emerald-400 font-mono block mt-1">{totalSent}</span>
              </div>
            </div>

            {totalSent > 0 && (
              <button 
                onClick={resetAllSentTags}
                className="w-full bg-slate-800 hover:bg-red-950 hover:text-red-200 hover:border-red-900 text-slate-300 font-bold text-[9.5px] py-1.5 rounded-lg transition-all border border-slate-700 cursor-pointer text-center"
              >
                איפוס סימוני פניות שליחה ({totalSent})
              </button>
            )}
          </div>

        </aside>

        {/* WORKSPACE PRESTIGE MAIN PORT - GORGEOUS GRID CARD */}
        <section className="flex-1 flex flex-col lg:overflow-hidden p-3 lg:p-4 gap-4">
          
          <div className="bg-white rounded-2xl border border-slate-200 flex-1 flex flex-col lg:overflow-hidden shadow-sm">
            
            {hasDateMappingIssue && (
              <div className="m-4 p-4 bg-rose-50 border border-rose-200 rounded-xl text-rose-950 flex flex-col md:flex-row items-start md:items-center justify-between gap-4 text-xs shadow-sm z-20">
                <div className="space-y-1">
                  <p className="font-bold flex items-center gap-1.5 text-rose-800">
                    <AlertCircle className="w-4 h-4 text-rose-600 animate-bounce" />
                    קליני: עמודת תאריך רישום/גיוס לא זוהתה באופן אוטומטי!
                  </p>
                  <p className="font-semibold text-rose-700 leading-normal">
                    סה"כ {participants.length} רשומות, אך עמודת התאריך המחושב ריקה. אנא הגדר/י עמודה במאפינג תחת ההגדרות או בחר/י ידנית:
                  </p>
                </div>
                <div className="flex items-center gap-2 bg-white/95 p-2 rounded-lg border border-rose-200 shrink-0 shadow-sm">
                  <span className="font-bold text-slate-700">בחר עמודת תאריך:</span>
                  <select 
                    onChange={(e) => updateMapping("תאריך גיוס", e.target.value)}
                    className="bg-white border border-slate-300 rounded px-2 py-1 font-bold text-slate-800 cursor-pointer text-xs focus:ring-1 focus:ring-rose-500 outline-none"
                  >
                    <option value="">-- בחר עמודה --</option>
                    {headers.map((h, i) => (
                      <option key={i} value={i}>{h}</option>
                    ))}
                  </select>
                </div>
              </div>
            )}

            {/* TAB-CONTROL BAR IN HIGH FIDELITY LAYOUT */}
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-col gap-4 shrink-0 justify-between select-none">
              
              <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                
                {/* Modern RTL Tab selection */}
                <div className="flex flex-wrap items-center gap-1.5 bg-slate-200/50 p-1 rounded-xl text-xs font-bold border border-slate-200/70 w-fit">
                  
                  <button 
                    type="button"
                    onClick={() => {
                      setActiveTab("matching");
                      setQuickFilterSample("all"); 
                    }}
                    className={`px-4 py-2 rounded-lg transition-all text-center flex items-center gap-2 cursor-pointer ${
                      activeTab === "matching" 
                        ? "bg-slate-900 text-white shadow-md font-bold scale-[1.02]" 
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <Clock className={`w-3.5 h-3.5 ${activeTab === "matching" ? "text-teal-400" : "text-slate-500"}`} />
                    <span>מיועדים למעקב ({targetFollowUpMonth} ח׳)</span>
                    <span className={`text-[9px] px-2 py-0.5 rounded-full font-bold font-mono ${
                      activeTab === "matching" ? "bg-teal-600 text-white" : "bg-slate-300 text-slate-800"
                    }`}>
                      {participants.filter(p => {
                        if (p.monthsElapsed === null) return false;
                        const isExactlyTargetMonth = exactMatchMode === "floor"
                          ? Math.floor(p.monthsElapsed) === targetFollowUpMonth
                          : Math.round(p.monthsElapsed) === targetFollowUpMonth;
                        return isExactlyTargetMonth && (!filterWithdrawn || !p.isWithdrawn);
                      }).length}
                    </span>
                  </button>

                  <button 
                    type="button"
                    onClick={() => setActiveTab("all")}
                    className={`px-4 py-2 rounded-lg transition-all text-center flex items-center gap-2 cursor-pointer ${
                      activeTab === "all" 
                        ? "bg-slate-900 text-white shadow-md font-bold scale-[1.02]" 
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <Database className={`w-3.5 h-3.5 ${activeTab === "all" ? "text-teal-400" : "text-slate-500"}`} />
                    <span>כל המטופלים</span>
                    <span className={`text-[9px] px-2 py-0.5 rounded-full font-bold font-mono ${
                      activeTab === "all" ? "bg-slate-800 text-teal-300" : "bg-slate-300 text-slate-800"
                    }`}>
                      {totalLoaded}
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      setActiveTab("matching");
                      setQuickFilterSample("withdrawn");
                    }}
                    className={`px-4 py-2 rounded-lg transition-all text-center flex items-center gap-2 cursor-pointer ${
                      quickFilterSample === "withdrawn"
                        ? "bg-slate-900 text-white shadow-md font-bold scale-[1.02]"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <UserX className={`w-3.5 h-3.5 ${quickFilterSample === "withdrawn" ? "text-rose-400" : "text-slate-500"}`} />
                    <span>פרשו מהמחקר</span>
                    <span className={`text-[9px] px-2 py-0.5 rounded-full font-bold font-mono ${
                      quickFilterSample === "withdrawn" ? "bg-rose-600 text-white" : "bg-slate-300 text-slate-800"
                    }`}>
                      {participants.filter(p => p.isWithdrawn).length}
                    </span>
                  </button>

                  <button 
                    type="button"
                    onClick={() => setActiveTab("sent")}
                    className={`px-4 py-2 rounded-lg transition-all text-center flex items-center gap-2 cursor-pointer ${
                      activeTab === "sent" 
                        ? "bg-slate-900 text-white shadow-md font-bold scale-[1.02]" 
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    <CheckSquare className={`w-3.5 h-3.5 ${activeTab === "sent" ? "text-teal-400" : "text-slate-500"}`} />
                    <span>שליחות שבוצעו</span>
                    <span className={`text-[9px] px-2 py-0.5 rounded-full font-bold font-mono ${
                      activeTab === "sent" ? "bg-emerald-600 text-white" : "bg-slate-300 text-slate-800"
                    }`}>
                      {totalSent}
                    </span>
                  </button>

                </div>

                {/* Left side: Pure Search Component */}
                <div className="relative w-full lg:w-80">
                  <Search className="w-4 h-4 absolute right-3.5 top-2.5 text-slate-400 stroke-[2.5]" />
                  <input 
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="חיפוש חופשי (שם, ת.ז. או נייד)..."
                    className="w-full pr-10 pl-8 py-2 bg-white border border-slate-250 rounded-xl text-xs outline-none focus:ring-2 focus:ring-teal-500 focus:border-teal-500 placeholder-slate-400 font-medium transition-all shadow-3xs"
                    dir="rtl"
                  />
                  {searchQuery && (
                    <button 
                      type="button"
                      onClick={() => setSearchQuery("")}
                      className="text-[10px] absolute left-3 top-2.5 text-slate-400 hover:text-slate-705 font-bold cursor-pointer"
                    >
                      נקה
                    </button>
                  )}
                </div>

              </div>

              {/* Sub filters */}
              <div className="flex flex-wrap items-center gap-2 pt-3 border-t border-slate-200/60 text-xs">
                <span className="text-slate-500 font-bold ml-1 flex items-center gap-1">
                  <Sliders className="w-3.5 h-3.5 text-teal-600" />
                  סינון לפי סטטוס משימה:
                </span>
                
                {[
                  { id: "all", label: `הכל (${filteredParticipants.length})` },
                  { id: "missing_stool", label: "🧫 חסרה דגימת צואה", count: participants.filter(p => {
                      const hasStool = p.stool && (p.stool.toUpperCase() === "V" || p.stool.includes("כן") || p.stool.includes("✓") || p.stool.includes("הוגשה"));
                      return !hasStool && (!filterWithdrawn || !p.isWithdrawn);
                    }).length
                  },
                  { id: "missing_prev", label: "⏮️ לא מסרו דגימה אשתקד", count: participants.filter(p => {
                      return p.missingPreviousStool && (!filterWithdrawn || !p.isWithdrawn);
                    }).length
                  },
                  { id: "returning", label: "🔄 משתתפים חוזרים" },
                  { id: "has_notes", label: "📝 עם הערות קליניות" }
                ].map((item: any) => (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => setQuickFilterSample(item.id)}
                    className={`px-3 py-1.5 rounded-full border transition-all cursor-pointer font-bold text-[11px] flex items-center gap-1.5 ${
                      quickFilterSample === item.id
                        ? "bg-teal-50 text-teal-800 border-teal-300 shadow-3xs"
                        : "bg-white hover:bg-slate-100 border-slate-200 text-slate-600"
                    }`}
                  >
                    <span>{item.label}</span>
                    {item.count !== undefined && (
                      <span className="bg-amber-100 text-amber-900 border border-amber-200 text-[9px] px-2 py-0.2 rounded-full font-bold">
                        {item.count}
                      </span>
                    )}
                  </button>
                ))}
              </div>

            </div>

            {/* PRESTIGE MEDICAL TABLE-BODY FLOW */}
            <div className="flex-1 overflow-auto">
              {filteredParticipants.length === 0 ? (
                <div className="py-20 text-center text-slate-500 space-y-4 select-none">
                  <div className="w-12 h-12 bg-amber-50 rounded-full flex items-center justify-center mx-auto text-amber-600 border border-amber-200">
                    <AlertCircle className="w-6 h-6 animate-pulse" />
                  </div>
                  <h4 className="text-sm font-bold text-slate-800">לא נמצאו מטופלים התואמים את הסינון הנוכחי</h4>
                  <div className="text-[11.5px] max-w-lg mx-auto text-slate-500 font-semibold leading-relaxed space-y-3">
                    {participants.length === 0 ? (
                      <p>הקובץ פתוח וריק כרגע. אנא טען קובץ אקסל שבועי או לחץ על הכפתור בצד ימין לטעינת נתוני סימולציה מהירה של שיבא.</p>
                    ) : (
                      <>
                        <p>לא נמצאו מטופלים העונים על תנאי חודש המעקב ({targetFollowUpMonth} ח׳) במצב החישוב הנוכחי.</p>
                        <div className="mt-3 bg-teal-50/70 border border-teal-200 p-3.5 rounded-xl text-teal-950 text-right leading-normal">
                          💡 <strong>מידע על הקובץ שלך:</strong> המטופלים במאגר זה נמצאים בחודשי מעקב הבאים:{" "}
                          <span className="font-mono font-bold text-teal-700 bg-white border border-teal-200 px-1.5 py-0.5 rounded">
                            {Array.from(new Set(participants.map(p => p.monthsElapsed !== null ? Math.floor(p.monthsElapsed) : null).filter((val): val is number => val !== null && val >= 0))).sort((a: number, b: number) => a - b).join(", ") || "אין תאריכים תקינים בקובץ"}
                          </span>
                          . נסה לשנות את <strong>"חודש יעד למעקב"</strong> בראש המסך לכל אחד החודשים הללו כדי לראותם מיידית!
                        </div>
                      </>
                    )}
                  </div>
                </div>
              ) : (
                <table className="w-full border-collapse select-text">
                  <thead className="bg-[#f8fafc] sticky top-0 border-b border-slate-200 text-[10.5px] font-black text-slate-500 uppercase z-10">
                    <tr>
                      <th className="p-3 text-center w-12">בחירה</th>
                      <th className="p-3 w-16">מזהה</th>
                      <th className="p-3">שם משתתף מלא</th>
                      <th className="p-3 w-28">תעודת זהות</th>
                      <th className="p-3 w-28 text-center">תאריך גיוס</th>
                      <th className="p-3 w-28 text-center">זמן שחלף (חלון)</th>
                      <th className="p-3 w-40">מעקב ביו-סמפלרז</th>
                      <th className="p-3 w-32 text-center">טלפון</th>
                      <th className="p-3 text-center w-28">סטטוס פנייה</th>
                    </tr>
                  </thead>
                  <tbody className="text-xs divide-y divide-slate-100">
                    {filteredParticipants.map((p) => {
                      const isSelected = p.id === selectedParticipantId;
                      const isSent = !!sentRecords[p.id];

                      const hasCompletedStool = p.stool && (p.stool.toUpperCase() === "V" || p.stool.includes("כן") || p.stool.includes("✓") || p.stool.includes("הוגשה"));
                      const hasCompletedSaliva = p.saliva && (p.saliva.toUpperCase() === "V" || p.saliva.includes("כן") || p.saliva.includes("✓") || p.saliva.includes("הוגשה"));
                      const hasCompletedSerum = p.serum && (p.serum.toUpperCase() === "V" || p.serum.includes("כן") || p.serum.includes("✓") || p.serum.includes("הוגשה"));

                      return (
                        <tr 
                          key={p.id}
                          ref={isSelected ? tableRowRef : null}
                          onClick={() => setSelectedParticipantId(p.id)}
                          className={`hover:bg-slate-50 transition-colors cursor-pointer ${
                            isSelected 
                              ? "bg-teal-50/50 font-black border-r-4 border-teal-600 scale-[0.99]" 
                              : isSent 
                              ? "bg-slate-100/80 text-slate-400"
                              : ""
                          }`}
                        >
                          <td className="p-3 text-center">
                            <span className={`w-3.5 h-3.5 rounded-full inline-block border ${
                              isSelected 
                                ? "bg-teal-650 border-teal-700 ring-2 ring-teal-500/20" 
                                : isSent 
                                ? "bg-emerald-500 border-emerald-600"
                                : "bg-white border-slate-300"
                            }`} />
                          </td>
                          <td className="p-3 font-mono text-[11px] text-slate-400 font-bold">#{p.id}</td>
                          <td className="p-3">
                            <span className="text-slate-900 font-bold">{p.firstName} {p.lastName}</span>
                            {p.returningParticipant && p.returningParticipant.trim() !== "" && (
                              <span className="bg-teal-50 border border-teal-150 text-teal-700 text-[8.5px] px-1.5 py-0.2 mr-2 rounded font-black select-none">משירה חוזרת</span>
                            )}
                            {p.missingPreviousStool && (
                              <span className="bg-amber-50 border border-amber-200 text-amber-700 text-[8.5px] px-1.5 py-0.2 mr-2 rounded font-black select-none">⚠️ חסרה דגימה משנה שעברה</span>
                            )}
                          </td>
                          <td className="p-3 font-mono text-slate-600">{p.tz || "-"}</td>
                          <td className="p-3 font-mono text-center text-slate-500">{p.recruitmentDateRaw}</td>
                          <td className="p-3 text-center font-mono font-bold text-teal-850">
                            {p.monthsElapsed ? `${p.monthsElapsed} ח׳` : "-"}
                          </td>
                          <td className="p-3">
                            {/* Gorgeous biological capsule indicators */}
                            <div className="flex gap-1 select-none">
                              <span 
                                title={`דגימת צואה: ${hasCompletedStool ? 'התקבלה במעבדה' : 'חסרה למעקב'}`}
                                className={`px-2 py-0.5 rounded-full text-[9px] font-black border transition-all ${
                                  hasCompletedStool 
                                    ? "bg-emerald-50 text-emerald-700 border-emerald-200" 
                                    : "bg-red-50 text-red-650 border-red-200"
                                }`}
                              >
                                צואה {hasCompletedStool ? '✓' : '✖'}
                              </span>
                              <span 
                                title={`דגימת רוק: ${hasCompletedSaliva ? 'התקבלה' : 'חסרה'}`}
                                className={`px-2 py-0.5 rounded-full text-[9px] font-black border transition-all ${
                                  hasCompletedSaliva 
                                    ? "bg-slate-100 text-slate-650 border-slate-200" 
                                    : "bg-red-50 text-red-650/80 border-red-150"
                                }`}
                              >
                                רוק {hasCompletedSaliva ? '✓' : '✖'}
                              </span>
                            </div>
                          </td>
                          <td className="p-3 font-mono text-center text-slate-650">{p.phone}</td>
                          <td className="p-3 text-center">
                            {isSent ? (
                              <span className="bg-emerald-100 border border-emerald-300 text-emerald-800 text-[10px] font-black px-2.5 py-0.5 rounded-full inline-flex items-center gap-1 shadow-sm">
                                <Check className="w-3.5 h-3.5 text-emerald-600 stroke-[3]" />
                                נשלח ✓
                              </span>
                            ) : (
                              <span className="bg-amber-50 border border-amber-200 text-amber-800 text-[10px] font-bold px-2.5 py-0.5 rounded-full inline-flex items-center gap-1 animate-pulse">
                                <Clock className="w-3.5 h-3.5 text-amber-600" />
                                ממתין
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>

          {/* DYNAMIC LOWER LAYOUT WORKSPACE FOR WHATSAPP COMMUNICATOR */}
          <div className="min-h-[400px] lg:h-[420px] shrink-0 grid grid-cols-1 md:grid-cols-12 bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm hover:shadow-md transition-shadow">
            
            {selectedParticipant ? (
              <>
                
                {/* Panel Left (2/3 width) - Detailed Preview, Selector and Whatsapp builder */}
                <div className="md:col-span-8 flex flex-col justify-between border-b md:border-b-0 md:border-l border-slate-200 p-4 h-full min-w-0">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-150 shrink-0 select-none">
                    
                    {/* Multi template swapper */}
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold text-slate-500">תבנית הודעה:</span>
                      <select
                        value={selectedTemplateId}
                        onChange={(e) => handleSelectTemplate(e.target.value)}
                        className="bg-slate-100 hover:bg-slate-200 font-bold text-slate-850 px-2 py-1 rounded-lg border border-slate-300 text-[11px] outline-none cursor-pointer focus:ring-1 focus:ring-teal-500"
                      >
                        {CLINICAL_TEMPLATES.map(ct => (
                          <option key={ct.id} value={ct.id}>{ct.name}</option>
                        ))}
                      </select>
                    </div>

                    <button 
                      type="button"
                      onClick={() => setIsEditingTemplate(!isEditingTemplate)}
                      className="text-[10.5px] text-teal-700 hover:text-teal-900 font-bold bg-teal-50/50 hover:bg-teal-50 py-1.5 px-3 rounded-lg border border-teal-200/50 cursor-pointer transition-colors"
                    >
                      {isEditingTemplate ? "🗄️ תצוגה מקדימה" : "📝 ערוך נוסח פנייה"}
                    </button>
                  </div>

                  {/* Operational Chat preview panel */}
                  <div className="flex-1 py-3 overflow-hidden flex flex-col justify-end">
                    {isEditingTemplate ? (
                      <div className="flex-1 flex flex-col gap-2 overflow-hidden">
                        <textarea 
                          value={templateText}
                          onChange={(e) => setTemplateText(e.target.value)}
                          className="w-full flex-1 bg-slate-900 text-slate-100 rounded-xl p-3 text-xs outline-none focus:ring-1 focus:ring-teal-500 border border-slate-800 resize-none leading-relaxed font-mono"
                          placeholder="נוסח הודעה חופשי..."
                        />
                        <div className="flex flex-wrap gap-1 items-center py-1 select-none shrink-0 text-xs">
                          <span className="text-[9px] text-slate-500 font-bold ml-1.5">הוסף פרמטר קליני:</span>
                          {[
                            { label: "שם פרטי", token: "שם פרטי" },
                            { label: "שם משפחה", token: "שם משפחה" },
                            { label: "תאריך רישום", token: "תאריך גיוס" },
                            { label: "קוד משתתף", token: "מספר משתתף" },
                            { label: "ת.ז", token: "ת.ז." },
                            { label: "חודשים שעברו", token: "חודשים שעברו" },
                          ].map(tk => (
                            <button 
                              key={tk.token}
                              onClick={() => addPlaceholderToken(tk.token)}
                              className="bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 rounded text-[9px] py-0.5 px-2 transition-all font-bold cursor-pointer"
                            >
                              +{tk.label}
                            </button>
                          ))}
                          <button 
                            onClick={() => {
                              if (window.confirm("לשחזר את נוסח הפנייה המקורי?")) {
                                setTemplateText(DEFAULT_TEMPLATE);
                              }
                            }}
                            className="text-[9.5px] text-red-650 hover:underline font-bold mr-auto cursor-pointer border-0 bg-transparent"
                          >
                            שחזר תבנית
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className="flex-1 bg-[#efeae2] border border-slate-300 rounded-xl p-4 overflow-y-auto relative flex flex-col justify-end shadow-inner">
                        <div className="absolute top-2 left-2 bg-white/90 text-[8.5px] font-bold text-[#128c7e] px-2 py-0.5 rounded-full select-none shadow-3xs border border-slate-200">
                          תצוגה מקדימה: סימולציית הודעת WhatsApp
                        </div>

                        <div className="bg-[#e2f4c5] border border-[#bcd39c] p-3 rounded-xl max-w-[85%] text-xs leading-relaxed text-slate-800 mr-auto ml-2 relative shadow-sm select-text">
                          <p className="whitespace-pre-line font-sans select-all font-medium">
                            {replacePlaceholders(templateText, selectedParticipant)}
                          </p>
                          <div className="text-left mt-2.5 text-[9px] text-slate-400 font-mono select-none flex items-center justify-end gap-1">
                            <span>{new Date().toLocaleTimeString('he-IL', {hour: '2-digit', minute:'2-digit'})}</span>
                            <span className="text-[#34b7f1] font-bold">✓✓</span>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Operational Launch WA CTAs - Simplified Preview Mode */}
                  <div className="space-y-3 shrink-0 select-none">
                    <button 
                      onClick={() => copyToClipboard(replacePlaceholders(templateText, selectedParticipant), selectedParticipant.id)}
                      className="w-full bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold text-[11px] py-2 rounded-lg border border-slate-200 transition-all flex items-center justify-center gap-1.5 cursor-pointer"
                    >
                      {copySuccessId === selectedParticipant.id ? (
                        <>
                          <Check className="w-3.5 h-3.5 text-teal-600 stroke-[2.5]" />
                          <span className="text-teal-700">התוכן הועתק!</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5 text-slate-500" />
                          <span>העתק תוכן הודעה</span>
                        </>
                      )}
                    </button>
                  </div>

                </div>

                {/* Panel Right (4/12 width) - Detailed clinician notes & paginator */}
                <div className="md:col-span-4 p-4 flex flex-col justify-between bg-[#f8fafc] h-full min-w-0">
                  
                  <div className="space-y-4">
                    <div className="flex items-center justify-between pb-1.5 border-b border-slate-200 select-none">
                      <span className="font-extrabold text-teal-850 text-xs">פרטי תיק פסיאנט</span>
                      <span className="text-[10px] font-bold text-slate-500 bg-slate-200/70 px-2 py-0.5 rounded-full font-mono">
                        #{selectedParticipant.id}
                      </span>
                    </div>

                    <div className="space-y-2.5 text-xs">
                      <div>
                        <span className="text-[9.5px] text-slate-400 block font-bold mb-0.5">שם מלא:</span>
                        <span className="font-black text-slate-900 text-sm leading-tight">{selectedParticipant.firstName} {selectedParticipant.lastName}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <span className="text-[9.5px] text-slate-400 block font-bold mb-0.5">תעודת זהות:</span>
                          <span className="font-bold text-slate-700 font-mono">{selectedParticipant.tz || "-"}</span>
                        </div>
                        <div>
                          <span className="text-[9.5px] text-slate-400 block font-bold mb-0.5">מספר טלפון:</span>
                          <span className="font-bold text-slate-800 font-mono">{selectedParticipant.phone || "-"}</span>
                        </div>
                      </div>
                      <div>
                        <span className="text-[9.5px] text-slate-400 block font-bold mb-0.5">תאריך ביקור אחרון:</span>
                        <span className="font-bold text-slate-800 font-mono">{selectedParticipant.recruitmentDateRaw || "-"}</span>
                      </div>
                    </div>

                    {/* Editor for clinical notes */}
                    <div className="space-y-1">
                      <span className="text-[9.5px] text-slate-400 block font-bold mb-0.5">הערות מעקב קליניות (ישמרו באקסל):</span>
                      <div className="relative">
                        <textarea
                          rows={2}
                          value={editingNotes}
                          onChange={(e) => setEditingNotes(e.target.value)}
                          placeholder="כתוב הערה כאן..."
                          className="w-full bg-white border border-slate-300 rounded-lg p-2 text-[10.5px] leading-snug outline-none focus:ring-1 focus:ring-teal-500 font-medium text-slate-800 shadow-3xs resize-none"
                        />
                        <button
                          onClick={() => handleSaveNotes(selectedParticipant.id, editingNotes)}
                          className="absolute bottom-2 left-2 bg-teal-600 hover:bg-teal-700 text-white font-bold text-[9px] px-2 py-1 rounded transition-colors shadow-3xs border-0 cursor-pointer"
                        >
                          שמור הערה 💾
                        </button>
                      </div>
                      {saveSuccess && (
                        <p className="text-[9px] text-emerald-600 font-bold block animate-bounce mt-1">ההערה עודכנה בגוף גליון האקסל בהצלחה!</p>
                      )}
                    </div>

                  </div>

                  {/* Operational navigation inside selection */}
                  <div className="space-y-3 pb-0.5 select-none shrink-0 border-t border-slate-200 pt-3">
                    <div className="flex items-center justify-between px-1">
                      <label className="flex items-center gap-2 cursor-pointer">
                        <div className="relative inline-flex items-center cursor-pointer">
                          <input
                            type="checkbox"
                            checked={autoAdvance}
                            onChange={(e) => setAutoAdvance(e.target.checked)}
                            className="sr-only peer"
                          />
                          <div className="w-7 h-3.5 bg-slate-300 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full rtl:peer-checked:after:-translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:start-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-2.5 after:w-2.5 after:transition-all peer-checked:bg-teal-600"></div>
                        </div>
                        <span className="text-[10px] font-bold text-slate-700">מעבר אוטומטי</span>
                      </label>

                      <div className="text-[10px] font-mono font-bold text-slate-500">
                         {filteredParticipants.findIndex(x => x.id === selectedParticipant.id) + 1} / {filteredParticipants.length}
                      </div>
                    </div>

                    <div className="flex gap-2 items-center">
                      <button
                        onClick={selectPrevCandidate}
                        disabled={filteredParticipants.findIndex(x => x.id === selectedParticipant.id) === 0}
                        className="bg-white hover:bg-slate-50 text-slate-700 p-2 rounded-xl border border-slate-200 transition-all disabled:opacity-20 cursor-pointer shadow-sm hover:border-teal-500 active:scale-95 group shrink-0"
                        title="הקודם (חץ ימינה)"
                      >
                        <ChevronRight className="w-5 h-5 group-hover:text-teal-600" />
                      </button>

                      <a
                        href={getWhatsAppUrl(selectedParticipant)}
                        target="_blank"
                        rel="noreferrer"
                        onClick={() => markAsSent(selectedParticipant.id)}
                        className="flex-1 bg-[#25D366] hover:bg-[#128C7E] active:bg-[#075E54] text-white font-black text-sm py-3 px-4 rounded-xl shadow-lg hover:shadow-[#25D366]/20 transition-all flex items-center justify-center gap-2 cursor-pointer text-center border-b-2 border-green-700 active:border-b-0 active:translate-y-0.5"
                      >
                        <Phone className="w-4 h-4 text-white fill-white" />
                        <span className="tracking-tight">שלח הודעה</span>
                      </a>

                      <button
                        onClick={selectNextCandidate}
                        disabled={filteredParticipants.findIndex(x => x.id === selectedParticipant.id) === filteredParticipants.length - 1}
                        className="bg-white hover:bg-slate-50 text-slate-700 p-2 rounded-xl border border-slate-200 transition-all disabled:opacity-20 cursor-pointer shadow-sm hover:border-teal-500 active:scale-95 group shrink-0"
                        title="הבא (חץ שמאלה)"
                      >
                        <ChevronLeft className="w-5 h-5 group-hover:text-teal-600" />
                      </button>
                    </div>

                    <div className="flex gap-2">
                      {sentRecords[selectedParticipant.id] ? (
                        <button 
                          onClick={() => removeSentMark(selectedParticipant.id)}
                          className="flex-1 bg-red-55/10 hover:bg-red-55/20 text-red-700 py-1.5 rounded-lg text-[10px] font-bold transition-all cursor-pointer border border-red-200 text-center"
                        >
                          בטל סימון "נשלח" ✖
                        </button>
                      ) : (
                        <button 
                          onClick={() => markAsSent(selectedParticipant.id)}
                          className="flex-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 py-1.5 rounded-lg text-[10px] font-bold transition-all cursor-pointer border border-emerald-250 text-center"
                        >
                          סמן כ"נשלח" באופן ידני ✓
                        </button>
                      )}
                    </div>
                  </div>

                </div>

              </>
            ) : (
              <div className="flex-1 flex flex-col items-center justify-center text-center p-8 space-y-2 select-none text-slate-400">
                <AlertCircle className="w-7 h-7 text-slate-300 animate-pulse" />
                <h4 className="text-slate-700 font-black text-xs">טרם נבחר פציינט מעקב מהטבלה</h4>
                <p className="text-[10.5px] text-slate-450 max-w-sm font-semibold leading-relaxed">אנא לחץ על אחד הבודקים המופיעים בטבלה מעלה על מנת להכין פניית ווטסאפ קלינית מעודכנת.</p>
              </div>
            )}

          </div>

        </section>

      </div>

    </div>
  );
}
