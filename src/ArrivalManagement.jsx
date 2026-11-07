import React, { useEffect, useMemo, useRef, useState } from "react";
import html2canvas from "html2canvas";
import * as XLSX from "xlsx";
import {
  collection,
  doc,
  getDocs,
  onSnapshot,
  writeBatch,
} from "firebase/firestore";
import { db } from "./firebase.js";
import "./ArrivalManagement.css";

const EMPTY_ROW = {
  date: "", supplier: "", customer: "", quality: "", width: "", weave: "",
  brt: "", meter: "", vehicle: "", remarks: "", type: "OWN",
};

const FIELD_LABELS = [
  ["supplier", "Supplier"], ["customer", "Customer"], ["quality", "Quality"],
  ["width", "Width"], ["weave", "Weave"], ["brt", "B / R / T"],
  ["meter", "Meter"], ["vehicle", "Vehicle #"], ["remarks", "Remarks"],
];

// Arrival Management ka data alag collection mein save hota hai.
// Existing Stock Manager collections (stock/current, stockBackups, etc.) touch nahi hotin.
const CLOUD_DOC_REF = doc(db, "arrivalManagementShared", "current");
const CLOUD_CHUNK_SIZE = 100;

const clean = (v) => String(v ?? "").trim();
const norm = (v) => clean(v).toLowerCase().replace(/[^a-z0-9]/g, "");
const number = (v) => Number(String(v ?? "").replace(/,/g, "")) || 0;

function normalizeDateForFilter(value) {
  const text = clean(value);
  if (!text) return "";

  // YYYY-MM-DD / YYYY/MM/DD
  let match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) {
    return `${match[1]}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;
  }

  // DD-MMM-YYYY (for example 06-Nov-2026)
  match = text.match(/^(\d{1,2})[-\s/.]([A-Za-z]{3,})[-\s/.](\d{2,4})$/);
  if (match) {
    const monthNames = [
      "jan", "feb", "mar", "apr", "may", "jun",
      "jul", "aug", "sep", "oct", "nov", "dec",
    ];
    const monthIndex = monthNames.indexOf(match[2].slice(0, 3).toLowerCase());
    if (monthIndex >= 0) {
      const year = match[3].length === 2 ? `20${match[3]}` : match[3];
      return `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(match[1]).padStart(2, "0")}`;
    }
  }

  // UNLOADING date string is M/D/YYYY (e.g. 11/7/2026 = 7 Nov 2026).
  match = text.match(/^(\d{1,2})[/ .\-](\d{1,2})[/ .\-](\d{2,4})$/);
  if (match) {
    const month = Number(match[1]);
    const day = Number(match[2]);
    const year = match[3].length === 2 ? `20${match[3]}` : match[3];
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
  }

  return "";
}

// UNLOADING J2 date uses M/D/YYYY format.
// Example: 11/7/2026 means 7 November 2026 (month first, day second).
function parseDailyArrivalDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
  }

  // Excel kabhi date ko serial number ki form mein return karta hai.
  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 80000) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
  }

  const text = clean(value);
  if (!text) return "";

  // M/D/YYYY (e.g. 11/7/2026 = 7 November 2026).
  // Date ke sath cell mein extra text ho tab bhi date extract ho jaye.
  const match = text.match(/(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/);
  if (match) {
    const month = Number(match[1]);
    const day = Number(match[2]);
    const year = Number(match[3]);

    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  return normalizeDateForFilter(text);
}

function getDateFromDailyHeading(
  row,
  headingIndex,
  worksheet,
  worksheetRowIndex,
  worksheetHeadingColumn
) {
  if (!Array.isArray(row)) return "";

  const candidates = [];
  const merges = worksheet?.["!merges"] || [];
  const range = worksheet?.["!ref"]
    ? XLSX.utils.decode_range(worksheet["!ref"])
    : { s: { r: 0, c: 0 }, e: { r: worksheetRowIndex, c: 8 } };

  // Pehle original worksheet mein heading ka exact row/column dhoondhein.
  // Is se sheet_to_json ke row-offset/blank-row differences ka asar nahi hota.
  let actualRow = worksheetRowIndex;
  let actualHeadingColumn = worksheetHeadingColumn;
  let nearestDistance = Number.POSITIVE_INFINITY;

  if (worksheet) {
    for (let r = range.s.r; r <= range.e.r; r += 1) {
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const cell = worksheet[XLSX.utils.encode_cell({ r, c })];
        const cellText = clean(cell?.v ?? cell?.w).replace(/\s+/g, " ");

        if (/DAILY\s+GREIGE\s+ARRIVAL/i.test(cellText)) {
          const distance = Math.abs(r - worksheetRowIndex);
          if (distance < nearestDistance) {
            nearestDistance = distance;
            actualRow = r;
            actualHeadingColumn = c;
          }
        }
      }
    }
  }

  // Aapki sheet mein title A:F merge hai aur date usi row mein G:I merge hai.
  // G cell (column index 6) ko explicitly check karte hain; merge metadata
  // browser/parser ki wajah se unavailable ho tab bhi date mil sakti hai.
  const candidateColumns = new Set([6]);

  const titleMerge = merges.find(
    (merge) =>
      merge.s.r === actualRow &&
      merge.s.c <= actualHeadingColumn &&
      merge.e.c >= actualHeadingColumn
  );

  if (titleMerge) {
    candidateColumns.add(titleMerge.e.c + 1);

    // Same heading row mein date merge ya koi aur merged date cell.
    merges
      .filter(
        (merge) =>
          merge.s.r === actualRow &&
          merge.s.c > titleMerge.e.c
      )
      .sort((a, b) => a.s.c - b.s.c)
      .forEach((merge) => candidateColumns.add(merge.s.c));
  }

  // Heading ke right side ke saare actual worksheet cells bhi check hon.
  for (
    let c = Math.max(actualHeadingColumn + 1, range.s.c);
    c <= range.e.c;
    c += 1
  ) {
    candidateColumns.add(c);
  }

  const tryCandidate = (value) => {
    const parsed = parseDailyArrivalDate(value);
    if (parsed) candidates.push(parsed);
  };

  // Sab se pehle direct G cell aur merged-date anchor ko parse karein.
  Array.from(candidateColumns)
    .filter((c) => c >= 0 && c <= range.e.c)
    .sort((a, b) => a - b)
    .forEach((c) => {
      const cell = worksheet?.[XLSX.utils.encode_cell({ r: actualRow, c })];
      if (cell) {
        tryCandidate(cell.v);
        tryCandidate(cell.w);
        tryCandidate(cell.h);
      }
    });

  if (candidates.length) return candidates[0];

  // Fallback: sheet_to_json ki usi row ke cells mein date dhoondhein.
  row.forEach((value, index) => {
    if (index > headingIndex) tryCandidate(value);
  });
  if (candidates.length) return candidates[0];

  // Kuch workbooks mein date merge G:I ke bajaye adjacent right cell se start hota hai.
  // Original raw cells ki poori row ko scan karna final fallback hai.
  if (worksheet) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: actualRow, c })];
      if (!cell) continue;
      const parsed =
        parseDailyArrivalDate(cell.v) ||
        parseDailyArrivalDate(cell.w) ||
        parseDailyArrivalDate(cell.h);
      if (parsed) return parsed;
    }
  }

  return parseDailyArrivalDate(row[headingIndex]);
}

function formatStockDate(dateKey) {
  if (!dateKey || dateKey === "undated") return "Date Not Available";

  const [year, month, day] = dateKey.split("-").map(Number);
  if (!year || !month || !day) return "Date Not Available";

  const date = new Date(year, month - 1, day, 12);
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "long",
    year: "numeric",
  }).format(date);
}

async function readCloudChunks(subcollectionName, chunkCount) {
  if (!chunkCount) return [];

  const snapshot = await getDocs(
    collection(db, "arrivalManagementShared", "current", subcollectionName)
  );

  return snapshot.docs
    .sort((a, b) => Number(a.id) - Number(b.id))
    .slice(0, chunkCount)
    .flatMap((item) => {
      const savedRows = item.data()?.rows;
      return Array.isArray(savedRows) ? savedRows : [];
    });
}

function parseSections(rows, kind) {
  const result = [];
  let type = "OWN";
  let headers = null;

  rows.forEach((row) => {
    const first = clean(row?.[0]).toUpperCase();

    if (first === "OWN" || first === "OWN GREIGE") {
      type = "OWN";
      headers = null;
      return;
    }

    if (
      first === "CMT" ||
      first === "COMMERCIAL" ||
      first === "COMMERCIAL GREIGE"
    ) {
      type = "CMT";
      headers = null;
      return;
    }

    if (!headers && row?.some((v) => /supplier/i.test(clean(v)))) {
      headers = row.map(norm);
      return;
    }

    if (!headers || !row?.some((v) => clean(v))) return;

    const get = (...names) => {
      const i = headers.findIndex((h) => names.includes(h));
      return i >= 0 ? row[i] : "";
    };

    const supplier = get("supplier");
    if (!clean(supplier) || /total/i.test(clean(supplier))) return;

    result.push({
      ...EMPTY_ROW,
      type,
      date: clean(get("date", "arrivaldate", "stockdate", "receiveddate", "receivingdate", "receiptdate", "dateofreceipt", "dateofarrival", "inwarddate", "challandate", "entrydate", "dated")),
      supplier: clean(supplier),
      customer: clean(get("customer")),
      quality: clean(get("quality")),
      width: clean(get("width")),
      weave: clean(get("weave")),
      brt: clean(get("brt")),
      meter: clean(get("meter", "meters")),
      vehicle: clean(get("vehicle", "vehicle#", "vehicleno")),
      remarks: clean(get("remarks")),
      source: kind,
    });
  });

  return result;
}

function parseStockRows(rows, worksheet) {
  const out = [];
  let type = "OWN";
  let headers = null;
  let currentDate = "";
  const worksheetRange = worksheet?.["!ref"]
    ? XLSX.utils.decode_range(worksheet["!ref"])
    : { s: { r: 0, c: 0 } };

  rows.forEach((row, rowIndex) => {
    const first = clean(row?.[0]).toUpperCase();
    const dailyHeadingIndex = Array.isArray(row)
      ? row.findIndex((value) => /DAILY\s+GREIGE\s+ARRIVAL/i.test(clean(value)))
      : -1;

    // Date DAILY GREIGE ARRIVAL wali merged heading row se uthti hai.
    if (dailyHeadingIndex >= 0) {
      currentDate = getDateFromDailyHeading(
        row,
        dailyHeadingIndex,
        worksheet,
        worksheetRange.s.r + rowIndex,
        worksheetRange.s.c + dailyHeadingIndex
      );
      headers = null;
      return;
    }

    if (first === "OWN" || first === "OWN GREIGE") {
      type = "OWN";
      headers = null;
      return;
    }

    if (
      first === "CMT" ||
      first === "COMMERCIAL" ||
      first === "COMMERCIAL GREIGE"
    ) {
      type = "CMT";
      headers = null;
      return;
    }

    if (!headers && row?.some((v) => /supplier/i.test(clean(v)))) {
      headers = row.map(norm);
      return;
    }

    if (!headers || !row?.some((v) => clean(v))) return;

    const get = (...names) => {
      const i = headers.findIndex((h) => names.includes(h));
      return i >= 0 ? row[i] : "";
    };

    const supplier = clean(get("supplier"));
    if (!supplier || /total/i.test(supplier)) return;

    out.push({
      ...EMPTY_ROW,
      type,
      date: currentDate || parseDailyArrivalDate(get(
        "date",
        "arrivaldate",
        "stockdate",
        "receiveddate",
        "receivingdate",
        "receiptdate",
        "dateofreceipt",
        "dateofarrival",
        "inwarddate",
        "challandate",
        "entrydate",
        "dated"
      )),
      supplier,
      quality: clean(get("quality")),
      width: clean(get("width")),
      weave: clean(get("weave")),
      brt: clean(get("brt")),
      meter: clean(get("meter")),
      vehicle: clean(get("vehicle")),
      remarks: clean(get("remarks")),
      customer: "",
      stockId: `${type}-${out.length}-${supplier}-${clean(get("vehicle"))}`,
    });
  });

  return out;
}

function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {
          type: "array",
          cellDates: true,
        });

        const sheets = {};
        const sheetObjects = {};

        wb.SheetNames.forEach((name) => {
          const key = name.toUpperCase();
          const worksheet = wb.Sheets[name];

          sheets[key] = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: "",
            raw: false,
          });
          // Keep raw worksheet too, because merged cells are not reliably
          // represented by sheet_to_json. It lets us read the date in merged G:I.
          sheetObjects[key] = worksheet;
        });

        resolve({ name: file.name, sheets, sheetObjects });
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () =>
      reject(new Error("Excel file could not be read."));

    reader.readAsArrayBuffer(file);
  });
}

function signature(r) {
  return [
    r.type,
    r.supplier,
    r.quality,
    r.width,
    r.weave,
    r.brt,
    r.meter,
    r.vehicle,
  ]
    .map((x) => norm(x))
    .join("|");
}

function vehicleCount(rows) {
  const all = rows.flatMap((r) =>
    clean(r.vehicle)
      .split(/[,;/]+/)
      .map((x) => x.trim())
      .filter(Boolean)
  );

  return new Set(all.map((x) => x.toUpperCase())).size;
}

function dateLabel(date) {
  if (!date) return "";

  const [y, m, d] = date.split("-");
  return `${Number(d)}/${Number(m)}/${y}`;
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");

  a.href = url;
  a.download = name;

  document.body.appendChild(a);
  a.click();
  a.remove();

  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// --------------------------------------------------
// REPORT TABLE
// --------------------------------------------------


function ReportTable({
  rows,
  type,
  summary = false,
  vehicles = 0,
}) {
  const filtered = rows.filter((r) => r.type === type);

  const meterTotal = filtered.reduce(
    (s, r) => s + number(r.meter),
    0
  );

  const columns = summary
    ? [
        ["supplier", "Supplier"],
        ["quality", "Quality"],
        ["width", "Width"],
        ["weave", "Weave"],
        ["brt", "B / R / T"],
        ["meter", "Meter"],
        ["vehicle", "Vehicle"],
        ["remarks", "Remarks"],
      ]
    : FIELD_LABELS;

  return (
    <div className="arrival-print-section">
      <div className="arrival-print-band">
        {type === "OWN" ? "Own Greige" : "Commercial Greige"}
      </div>

      <table className="arrival-report-table">
        <thead>
          <tr>
            <th>S#</th>
            {columns.map(([key, label]) => (
              <th key={key}>{label}</th>
            ))}
          </tr>
        </thead>

        <tbody>
          {filtered.map((r, i) => (
            <tr key={r.id || `${type}-${i}`}>
              <td>{i + 1}</td>
              {columns.map(([key]) => (
                <td key={key}>{r[key]}</td>
              ))}
            </tr>
          ))}

          <tr className="arrival-blank-row">
            <td colSpan={columns.length + 1}></td>
          </tr>
        </tbody>

        <tfoot>
          <tr className="arrival-section-total">
            {summary ? (
              <>
                <td colSpan={6}>
                  {type === "OWN" ? "Own Total" : "Cmt Total"}
                </td>
                <td>{meterTotal.toLocaleString()}</td>
                <td colSpan={2}>{vehicles} VEHICLES</td>
              </>
            ) : (
              <>
                <td colSpan={7}>
                  {type === "OWN" ? "Own Total" : "Cmt Total"}
                </td>
                <td>{meterTotal.toLocaleString()}</td>
                <td colSpan={2}>{vehicles} VEHICLES</td>
              </>
            )}
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

// --------------------------------------------------
// MAIN COMPONENT
// --------------------------------------------------

export default function ArrivalManagement({ onHome }) {
  const [section, setSection] = useState("home");

  // Home screen se sirf aik shared workbook upload hoti hai (STOCK + UNLOADING).
  const [stockFile, setStockFile] = useState(null);
  const [stockBook, setStockBook] = useState(null);

  const [stockRows, setStockRows] = useState([]);
  const [rows, setRows] = useState([]);

  const [reportDate, setReportDate] = useState(
    new Date().toLocaleDateString("en-CA")
  );

  const [showAuto, setShowAuto] = useState(false);
  const [selected, setSelected] = useState([]);

  // Stock filters apply only when Search is pressed.
  const EMPTY_STOCK_FILTERS = {
    date: "",
    quality: "",
    width: "",
    supplier: "",
    vehicle: "",
  };
  const [stockFilterDraft, setStockFilterDraft] = useState({ ...EMPTY_STOCK_FILTERS });
  const [stockFilters, setStockFilters] = useState({ ...EMPTY_STOCK_FILTERS });
  const [showStockDateWiseTotals, setShowStockDateWiseTotals] = useState(false);

  const [manual, setManual] = useState({ ...EMPTY_ROW });

  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [cloudReady, setCloudReady] = useState(false);
  const [cloudLoading, setCloudLoading] = useState(true);
  const [cloudSaving, setCloudSaving] = useState(false);

  const lastSavedRevisionRef = useRef("");
  const lastAppliedRevisionRef = useRef("");
  const loadingRevisionRef = useRef("");
  const lastPersistedPayloadRef = useRef("");
  const saveQueueRef = useRef(Promise.resolve());

  // Manual vehicle count fields. Blank means automatic count.
  const [manualOwnVehicles, setManualOwnVehicles] = useState("");
  const [manualCmtVehicles, setManualCmtVehicles] = useState("");

  const ownRows = useMemo(
    () => rows.filter((r) => r.type === "OWN"),
    [rows]
  );

  const cmtRows = useMemo(
    () => rows.filter((r) => r.type === "CMT"),
    [rows]
  );

  const ownMeters = ownRows.reduce(
    (s, r) => s + number(r.meter),
    0
  );

  const cmtMeters = cmtRows.reduce(
    (s, r) => s + number(r.meter),
    0
  );

  const totalMeters = ownMeters + cmtMeters;

  const autoOwnVehicles = vehicleCount(ownRows);
  const autoCmtVehicles = vehicleCount(cmtRows);

  const ownVehicles =
    manualOwnVehicles.trim() !== ""
      ? Math.max(0, Math.floor(number(manualOwnVehicles)))
      : autoOwnVehicles;

  const cmtVehicles =
    manualCmtVehicles.trim() !== ""
      ? Math.max(0, Math.floor(number(manualCmtVehicles)))
      : autoCmtVehicles;

  const vehicles = ownVehicles + cmtVehicles;

  const usedSignatures = useMemo(
    () => new Set(rows.map(signature)),
    [rows]
  );

  const availableRows = stockRows.filter(
    (r) => !usedSignatures.has(signature(r))
  );

  const filteredStockRows = useMemo(() => {
    const dateFilter = stockFilters.date;
    const qualityFilter = norm(stockFilters.quality);
    const widthFilter = norm(stockFilters.width);
    const supplierFilter = norm(stockFilters.supplier);
    const vehicleFilter = norm(stockFilters.vehicle);

    return stockRows.filter((r) => {
      if (dateFilter && normalizeDateForFilter(r.date) !== dateFilter) {
        return false;
      }

      // Text filters use includes(), so text from the start, middle or end matches.
      if (qualityFilter && !norm(r.quality).includes(qualityFilter)) return false;
      if (widthFilter && !norm(r.width).includes(widthFilter)) return false;
      if (supplierFilter && !norm(r.supplier).includes(supplierFilter)) return false;
      if (vehicleFilter && !norm(r.vehicle).includes(vehicleFilter)) return false;

      return true;
    });
  }, [stockRows, stockFilters]);

  const filteredStockGroups = useMemo(() => {
    const groups = new Map();

    filteredStockRows.forEach((row) => {
      const dateKey = normalizeDateForFilter(row.date) || "undated";
      if (!groups.has(dateKey)) groups.set(dateKey, []);
      groups.get(dateKey).push(row);
    });

    return Array.from(groups.entries())
      .map(([dateKey, groupRows]) => ({
        dateKey,
        label: formatStockDate(dateKey),
        ownRows: groupRows.filter((row) => row.type === "OWN"),
        cmtRows: groupRows.filter((row) => row.type === "CMT"),
      }))
      .sort((a, b) => {
        if (a.dateKey === "undated") return 1;
        if (b.dateKey === "undated") return -1;
        return b.dateKey.localeCompare(a.dateKey);
      });
  }, [filteredStockRows]);

  // Latest date poore stock data se nikalti hai, filters se nahi.
  // Is liye purani date filter karne par usay ghalti se "Latest" mark nahi kiya jayega.
  const latestStockDate = useMemo(() => {
    return stockRows.reduce((latest, row) => {
      const dateKey = normalizeDateForFilter(row.date);
      if (!dateKey) return latest;
      return !latest || dateKey > latest ? dateKey : latest;
    }, "");
  }, [stockRows]);

  // Arrival Stock ke overall OWN / Commercial / Grand meter totals.
  const stockMeterTotals = useMemo(() => {
    const own = stockRows
      .filter((row) => row.type === "OWN")
      .reduce((sum, row) => sum + number(row.meter), 0);
    const commercial = stockRows
      .filter((row) => row.type !== "OWN")
      .reduce((sum, row) => sum + number(row.meter), 0);

    return { own, commercial, total: own + commercial };
  }, [stockRows]);

  // Date-wise OWN / Commercial / Grand totals; latest date appears first.
  const stockDateWiseTotals = useMemo(() => {
    const grouped = new Map();

    stockRows.forEach((row) => {
      const dateKey = normalizeDateForFilter(row.date) || "undated";
      if (!grouped.has(dateKey)) {
        grouped.set(dateKey, { dateKey, own: 0, commercial: 0, total: 0 });
      }

      const item = grouped.get(dateKey);
      const meters = number(row.meter);
      if (row.type === "OWN") item.own += meters;
      else item.commercial += meters;
      item.total += meters;
    });

    return Array.from(grouped.values()).sort((a, b) => {
      if (a.dateKey === "undated") return 1;
      if (b.dateKey === "undated") return -1;
      return b.dateKey.localeCompare(a.dateKey);
    });
  }, [stockRows]);

  // Auto Add list bhi Stock ki tarah date-wise latest-first rahegi.
  const availableDateGroups = useMemo(() => {
    const groups = new Map();

    availableRows.forEach((row) => {
      const dateKey = normalizeDateForFilter(row.date) || "undated";
      if (!groups.has(dateKey)) groups.set(dateKey, []);
      groups.get(dateKey).push(row);
    });

    return Array.from(groups.entries())
      .map(([dateKey, groupRows]) => ({
        dateKey,
        label: formatStockDate(dateKey),
        ownRows: groupRows.filter((row) => row.type === "OWN"),
        cmtRows: groupRows.filter((row) => row.type !== "OWN"),
      }))
      .sort((a, b) => {
        if (a.dateKey === "undated") return 1;
        if (b.dateKey === "undated") return -1;
        return b.dateKey.localeCompare(a.dateKey);
      });
  }, [availableRows]);

  const latestAvailableDate = useMemo(() => {
    return availableRows.reduce((latest, row) => {
      const dateKey = normalizeDateForFilter(row.date);
      if (!dateKey) return latest;
      return !latest || dateKey > latest ? dateKey : latest;
    }, "");
  }, [availableRows]);

  // Unloading cards ko bhi date ke hisaab se group aur latest-first sort karein.
  const unloadingDateGroups = useMemo(() => {
    const groups = new Map();

    rows.forEach((row) => {
      const dateKey = normalizeDateForFilter(row.date) || "undated";
      if (!groups.has(dateKey)) groups.set(dateKey, []);
      groups.get(dateKey).push(row);
    });

    return Array.from(groups.entries())
      .map(([dateKey, groupRows]) => ({
        dateKey,
        label: formatStockDate(dateKey),
        ownRows: groupRows.filter((row) => row.type === "OWN"),
        cmtRows: groupRows.filter((row) => row.type !== "OWN"),
      }))
      .sort((a, b) => {
        if (a.dateKey === "undated") return 1;
        if (b.dateKey === "undated") return -1;
        return b.dateKey.localeCompare(a.dateKey);
      });
  }, [rows]);

  const latestUnloadingDate = useMemo(() => {
    return rows.reduce((latest, row) => {
      const dateKey = normalizeDateForFilter(row.date);
      if (!dateKey) return latest;
      return !latest || dateKey > latest ? dateKey : latest;
    }, "");
  }, [rows]);

  function applyStockFilters() {
    setStockFilters({
      date: stockFilterDraft.date,
      quality: stockFilterDraft.quality,
      width: stockFilterDraft.width,
      supplier: stockFilterDraft.supplier,
      vehicle: stockFilterDraft.vehicle,
    });
  }

  function clearStockFilters() {
    const empty = { date: "", quality: "", width: "", supplier: "", vehicle: "" };
    setStockFilterDraft(empty);
    setStockFilters(empty);
  }

  // --------------------------------------------------
  // FIREBASE LOAD: shared data sab users ko milega
  // --------------------------------------------------

  useEffect(() => {
    let active = true;

    const unsubscribe = onSnapshot(
      CLOUD_DOC_REF,
      async (snapshot) => {
        if (!active) return;

        if (!snapshot.exists()) {
          setCloudLoading(false);
          setCloudReady(true);
          return;
        }

        const meta = snapshot.data();
        const revision = String(meta.revision || "");

        // Apni hi save ki hui update dobara state par apply na karein.
        if (
          revision &&
          (revision === lastSavedRevisionRef.current ||
            revision === lastAppliedRevisionRef.current ||
            revision === loadingRevisionRef.current)
        ) {
          return;
        }

        loadingRevisionRef.current = revision;
        setCloudLoading(true);

        try {
          const [remoteStockRows, remoteRows] = await Promise.all([
            readCloudChunks("stockChunks", Number(meta.stockChunkCount) || 0),
            readCloudChunks("unloadingChunks", Number(meta.unloadingChunkCount) || 0),
          ]);

          if (!active || revision !== loadingRevisionRef.current) return;

          const remotePayload = {
            stockRows: remoteStockRows,
            rows: remoteRows,
            reportDate: meta.reportDate || new Date().toLocaleDateString("en-CA"),
            manualOwnVehicles: meta.manualOwnVehicles ?? "",
            manualCmtVehicles: meta.manualCmtVehicles ?? "",
            stockFileName: meta.stockFileName || "",
            previousFileName: "",
          };

          lastAppliedRevisionRef.current = revision;
          lastPersistedPayloadRef.current = JSON.stringify(remotePayload);

          setStockRows(remotePayload.stockRows);
          setRows(remotePayload.rows);
          setReportDate(remotePayload.reportDate);
          setManualOwnVehicles(remotePayload.manualOwnVehicles);
          setManualCmtVehicles(remotePayload.manualCmtVehicles);
          setStockFile(null);
          setStockBook(
            remotePayload.stockFileName
              ? { name: remotePayload.stockFileName }
              : null
          );
          setError("");
        } catch (e) {
          console.error("Firebase load error:", e);
          if (active && revision === loadingRevisionRef.current) {
            setError(
              "Firebase se saved Arrival data load nahi ho saka. Firestore Rules / internet check karein."
            );
          }
        } finally {
          if (active && revision === loadingRevisionRef.current) {
            setCloudLoading(false);
            setCloudReady(true);
            loadingRevisionRef.current = "";
          }
        }
      },
      (e) => {
        console.error("Firebase listener error:", e);
        if (!active) return;
        setError(
          "Firebase connection nahi ho saki. Firestore Rules / internet check karein."
        );
        setCloudLoading(false);
        setCloudReady(true);
      }
    );

    return () => {
      active = false;
      unsubscribe();
    };
  }, []);

  // Firebase mein wahi app data save hoga; purane Stock Manager collections alag hain.
  const cloudPayload = useMemo(
    () => ({
      stockRows,
      rows,
      reportDate,
      manualOwnVehicles,
      manualCmtVehicles,
      stockFileName: stockBook?.name || stockFile?.name || "",
      previousFileName: "",
    }),
    [
      stockRows,
      rows,
      reportDate,
      manualOwnVehicles,
      manualCmtVehicles,
      stockBook?.name,
      stockFile?.name,
    ]
  );

  const cloudPayloadText = useMemo(
    () => JSON.stringify(cloudPayload),
    [cloudPayload]
  );

  // --------------------------------------------------
  // FIREBASE SAVE: atomic batch, pehle se saved state safe rahegi
  // --------------------------------------------------

  useEffect(() => {
    if (!cloudReady || cloudLoading) return;
    if (lastPersistedPayloadRef.current === cloudPayloadText) return;

    const timer = setTimeout(() => {
      const payloadToSave = cloudPayload;
      const payloadTextToSave = cloudPayloadText;

      const saveTask = async () => {
        const revision = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        const stockChunkCount = Math.ceil(
          payloadToSave.stockRows.length / CLOUD_CHUNK_SIZE
        );
        const unloadingChunkCount = Math.ceil(
          payloadToSave.rows.length / CLOUD_CHUNK_SIZE
        );

        // Firestore batch limit 500 writes hai. Is check se old data safe rahega.
        const totalWrites = stockChunkCount + unloadingChunkCount + 1;
        if (totalWrites > 500) {
          throw new Error(
            "Arrival data Firestore batch limit se zyada hai. Data save nahi hua; pehle se saved data safe hai."
          );
        }

        const batch = writeBatch(db);

        for (let i = 0; i < stockChunkCount; i += 1) {
          const chunkRef = doc(
            db,
            "arrivalManagementShared",
            "current",
            "stockChunks",
            String(i)
          );
          batch.set(chunkRef, {
            rows: payloadToSave.stockRows.slice(
              i * CLOUD_CHUNK_SIZE,
              (i + 1) * CLOUD_CHUNK_SIZE
            ),
          });
        }

        for (let i = 0; i < unloadingChunkCount; i += 1) {
          const chunkRef = doc(
            db,
            "arrivalManagementShared",
            "current",
            "unloadingChunks",
            String(i)
          );
          batch.set(chunkRef, {
            rows: payloadToSave.rows.slice(
              i * CLOUD_CHUNK_SIZE,
              (i + 1) * CLOUD_CHUNK_SIZE
            ),
          });
        }

        batch.set(
          CLOUD_DOC_REF,
          {
            revision,
            stockChunkCount,
            unloadingChunkCount,
            stockFileName: payloadToSave.stockFileName,
            previousFileName: payloadToSave.previousFileName,
            reportDate: payloadToSave.reportDate,
            manualOwnVehicles: payloadToSave.manualOwnVehicles,
            manualCmtVehicles: payloadToSave.manualCmtVehicles,
            updatedAt: new Date().toISOString(),
          },
          { merge: true }
        );

        lastSavedRevisionRef.current = revision;
        setCloudSaving(true);
        await batch.commit();
        lastPersistedPayloadRef.current = payloadTextToSave;
        setError("");
      };

      saveQueueRef.current = saveQueueRef.current
        .catch(() => {})
        .then(saveTask)
        .catch((e) => {
          console.error("Firebase save error:", e);
          setError(
            e.message ||
              "Arrival data Firebase par save nahi ho saka. Purana saved data delete nahi hua."
          );
        })
        .finally(() => setCloudSaving(false));
    }, 900);

    return () => clearTimeout(timer);
  }, [cloudReady, cloudLoading, cloudPayload, cloudPayloadText]);

  // --------------------------------------------------
  // LOAD ONE WORKBOOK: STOCK + UNLOADING
  // --------------------------------------------------

  async function loadWorkbook(file) {
    if (!file) return;

    setError("");

    try {
      setBusy(true);

      const book = await parseWorkbook(file);
      const stockSheet = book.sheets.STOCK;
      const unloadingSheet = book.sheets.UNLOADING;

      // Dono required sheets honi chahiyein; invalid file se saved data replace nahi hoga.
      if (!stockSheet || !unloadingSheet) {
        setError(
          "Excel file mein STOCK aur UNLOADING dono sheets honi chahiyein. Purana saved data safe hai."
        );
        return;
      }

      const parsedStock = parseStockRows(
        stockSheet,
        book.sheetObjects?.STOCK
      );
      if (!parsedStock.length) {
        setError(
          "STOCK sheet mein valid rows nahi milin. OWN/CMT headings aur Supplier header check karein. Purana saved data safe hai."
        );
        return;
      }

      const hasUnloadingHeader = unloadingSheet.some((row) =>
        row?.some((value) => /supplier/i.test(clean(value)))
      );

      if (!hasUnloadingHeader) {
        setError(
          "UNLOADING sheet mein Supplier header nahi mila. Purana saved data safe hai."
        );
        return;
      }

      // Same file ki STOCK sheet Arrival Stock mein aur UNLOADING sheet
      // Unloading / Summary mein load hoti hai. Empty unloading rows allowed hain.
 const unloadingJ2 = book.sheetObjects?.UNLOADING?.J2;
 const unloadingDateRaw = [
  unloadingJ2?.v,
  unloadingJ2?.w,
  unloadingJ2?.h,
  unloadingSheet?.[1]?.[9],
].find((value) => value !== undefined && value !== null && clean(value) !== "") ?? "";

const unloadingSheetDate =
  parseDailyArrivalDate(unloadingDateRaw) ||
  normalizeDateForFilter(unloadingDateRaw) ||
  "";

console.log("UNLOADING J2 DATE:", unloadingDateRaw);
console.log("PARSED UNLOADING DATE:", unloadingSheetDate);

const parsedUnloading = parseSections(unloadingSheet, "current").map(
  (row, index) => ({
    ...row,
    date: normalizeDateForFilter(row.date) || unloadingSheetDate,
    id: `sheet-${Date.now()}-${index}`,
  })
);
      setStockFile(file);
      setStockBook({ name: book.name });
      setStockRows(parsedStock);
      setRows(parsedUnloading);
      setSelected([]);
      clearStockFilters();
    } catch (e) {
      setError(e.message || "Arrival workbook could not be loaded.");
    } finally {
      setBusy(false);
    }
  }

  // --------------------------------------------------
  // ADD / UPDATE / REMOVE
  // --------------------------------------------------

  function addSelected() {
    const additions = availableRows.filter((r) =>
      selected.includes(r.stockId)
    );

    setRows((old) => [
      ...old,
      ...additions.map((r, i) => ({
        ...r,
        id: `auto-${Date.now()}-${i}`,
      })),
    ]);

    setSelected([]);
    setShowAuto(false);
  }

  function addManual(e) {
    e.preventDefault();

    if (!manual.supplier.trim() || !manual.vehicle.trim()) {
      setError(
        "Manual entry ke liye Supplier aur Vehicle # zaroori hain."
      );
      return;
    }

    setRows((old) => [
      ...old,
      {
        ...manual,
        id: `manual-${Date.now()}`,
      },
    ]);

    setManual({ ...EMPTY_ROW });
    setError("");
  }

  function updateRow(id, key, value) {
    setRows((old) =>
      old.map((r) =>
        r.id === id ? { ...r, [key]: value } : r
      )
    );
  }

  function removeRow(id) {
    setRows((old) => old.filter((r) => r.id !== id));
  }

  // --------------------------------------------------
  // IMAGE EXPORT
  // --------------------------------------------------

  async function exportImage(kind) {
    const target = document.getElementById(
      `arrival-${kind}-print`
    );

    if (!target) return;

    let clone = null;

    try {
      setError("");

      clone = target.cloneNode(true);

      clone.style.width = kind === "summary" ? "1200px" : "max-content";
      clone.style.minWidth = kind === "summary" ? "1200px" : "1100px";
      clone.style.maxWidth = "none";
      clone.style.padding = "24px";
      clone.style.background = "#fff";
      clone.style.position = "fixed";
      clone.style.display = "inline-block";
      clone.style.left = "-12000px";
      clone.style.top = "0";
      clone.style.transform = "none";

      document.body.appendChild(clone);

      // Tables ko content ke mutabiq auto width dein
      clone.querySelectorAll(".arrival-report-table").forEach((table) => {
        table.style.tableLayout = kind === "summary" ? "fixed" : "auto";
        table.style.width = kind === "summary" ? "100%" : "max-content";
        table.style.minWidth = kind === "summary" ? "0" : "100%";
        table.style.maxWidth = "none";
        table.style.borderCollapse = "collapse";

        table.querySelectorAll("th, td").forEach((cell) => {
          cell.style.whiteSpace = kind === "summary" ? "normal" : "nowrap";
          cell.style.wordBreak = kind === "summary" ? "break-word" : "normal";
          cell.style.overflowWrap = kind === "summary" ? "anywhere" : "normal";
          cell.style.width = "auto";
          cell.style.height = "auto";
          cell.style.padding = "4px 8px";
          cell.style.lineHeight = "1.2";
          cell.style.fontSize = "13px";
        });

        // Summary aur Unloading mein Weave column ki position alag hai
        const isSummary = kind === "summary";

        const widthColumn = isSummary ? 4 : 5;
        const weaveColumn = isSummary ? 5 : 6;
        const brtColumn = isSummary ? 6 : 7;
        const meterColumn = isSummary ? 7 : 8;

        // Summary image: columns ko fixed proportions dein taa-ke
        // long supplier/remarks/weave text poori image ko wide na kare.
        if (kind === "summary") {
          const summaryColumnWidths = ["3%", "15%", "12%", "6%", "18%", "8%", "10%", "12%", "16%"];
          table.querySelectorAll("thead tr, tbody tr").forEach((row) => {
            row.querySelectorAll("th, td").forEach((cell, index) => {
              const width = summaryColumnWidths[index];
              if (width) {
                cell.style.width = width;
                cell.style.minWidth = "0";
                cell.style.maxWidth = width;
              }
            });
          });
        }

        // Width compact
        table.querySelectorAll(
          `th:nth-child(${widthColumn}), td:nth-child(${widthColumn})`
        ).forEach((cell) => {
          cell.style.width = "55px";
          cell.style.minWidth = "55px";
          cell.style.maxWidth = "70px";
        });

        // B/R/T compact
        table.querySelectorAll(
          `th:nth-child(${brtColumn}), td:nth-child(${brtColumn})`
        ).forEach((cell) => {
          cell.style.width = "75px";
          cell.style.minWidth = "65px";
        });

        // Meter compact
        table.querySelectorAll(
          `th:nth-child(${meterColumn}), td:nth-child(${meterColumn})`
        ).forEach((cell) => {
          cell.style.width = "80px";
          cell.style.minWidth = "70px";
        });

        // Weave ko zyada space
        table.querySelectorAll(
          `th:nth-child(${weaveColumn}), td:nth-child(${weaveColumn})`
        ).forEach((cell) => {
          cell.style.width = "220px";
          cell.style.minWidth = "220px";
          cell.style.maxWidth = "350px";
          cell.style.whiteSpace = "normal";
          cell.style.wordBreak = "break-word";
          cell.style.overflowWrap = "anywhere";
        });
      });

      // Summary totals ko reference image ki tarah compact aur center mein rakhein.
      // Inline !important styles existing app CSS ke full-width table rules ko override karte hain.
      if (kind === "summary") {
        const totalsTable = clone.querySelector(".arrival-summary-totals");

        if (totalsTable) {
          totalsTable.style.setProperty("display", "table", "important");
          totalsTable.style.setProperty("width", "58%", "important");
          totalsTable.style.setProperty("min-width", "0", "important");
          totalsTable.style.setProperty("max-width", "58%", "important");
          totalsTable.style.setProperty("margin-left", "auto", "important");
          totalsTable.style.setProperty("margin-right", "auto", "important");
          totalsTable.style.setProperty("table-layout", "fixed", "important");
          totalsTable.style.setProperty("border-collapse", "collapse", "important");

          const summaryWidths = ["31%", "39%", "13%", "17%"];
          totalsTable.querySelectorAll("tr").forEach((row) => {
            row.querySelectorAll("td, th").forEach((cell, index) => {
              if (summaryWidths[index]) {
                cell.style.setProperty("width", summaryWidths[index], "important");
                cell.style.setProperty("min-width", "0", "important");
                cell.style.setProperty("max-width", summaryWidths[index], "important");
                cell.style.setProperty("white-space", "nowrap", "important");
              }
            });
          });
        }
      }

      clone.querySelectorAll(".arrival-summary-totals td, .arrival-summary-totals th").forEach((cell) => {
        cell.style.fontSize = "12px";
      });

      clone.querySelectorAll(".arrival-grand-total-row td").forEach((cell) => {
        cell.style.fontSize = "13px";
        cell.style.fontWeight = "900";
      });

      clone.querySelectorAll(".arrival-section-total td").forEach((cell) => {
        cell.style.fontSize = "13px";
        cell.style.fontWeight = "900";
      });

      clone.querySelectorAll("tr").forEach((row) => {
        row.style.height = "auto";
      });

      // Summary image ko browser ki canvas limit se bachane ke liye
      // output dimensions ke mutabiq scale automatically adjust karein.
      const exportWidth = Math.max(clone.scrollWidth, clone.offsetWidth, 1);
      const exportHeight = Math.max(clone.scrollHeight, clone.offsetHeight, 1);
      const summaryScale = Math.min(
        2,
        2400 / exportWidth,
        16000 / exportHeight,
        Math.sqrt(80000000 / (exportWidth * exportHeight))
      );

      const canvas = await html2canvas(clone, {
        scale: kind === "summary" ? summaryScale : 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        logging: false,
        scrollX: 0,
        scrollY: 0,
        windowWidth: Math.max(1200, exportWidth),
        windowHeight: exportHeight,
      });

      if (kind === "summary") {
        const png = await new Promise((resolve) => {
          canvas.toBlob(resolve, "image/png");
        });

        if (!png) {
          throw new Error("Summary image create nahi ho saki.");
        }

        downloadBlob(png, `Arrival-${kind}-${reportDate}.png`);
      } else {
        canvas.toBlob((png) => {
          if (png) {
            downloadBlob(
              png,
              `Arrival-${kind}-${reportDate}.png`
            );
          } else {
            setError("Image create nahi ho saki.");
          }
        }, "image/png");
      }
    } catch (err) {
      console.error("Image export error:", err);
      setError("Image download nahi ho saki. Dobara try karein.");
    } finally {
      if (clone) clone.remove();
    }
  }

  // --------------------------------------------------
  // PRINT REPORT - A4 PORTRAIT ONE PAGE
  // --------------------------------------------------

  function printReport(kind) {
    const target = document.getElementById(
      `arrival-${kind}-print`
    );

    if (!target) return;

    const win = window.open(
      "",
      "_blank",
      "width=1000,height=1200"
    );

    if (!win) {
      setError("Pop-up allow karein, phir PDF print karein.");
      return;
    }

    const isSummary = kind === "summary";

    const printCss = `
      @page {
        size: A4 portrait;
        margin: 8mm;
      }

      * {
        box-sizing: border-box;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }

      html, body {
        margin: 0;
        padding: 0;
        color: #111;
        background: #fff;
        font-family: Arial, Helvetica, sans-serif;
      }

      body {
        width: 194mm;
      }

      .arrival-fit-area {
        width: 194mm;
        height: 281mm;
        overflow: hidden;
        position: relative;
      }

      .arrival-print-report {
        width: 194mm;
        max-width: none;
        min-width: 0;
        margin: 0;
        padding: 0;
        background: #fff;
        transform-origin: top left;
      }

      .arrival-report-title {
        text-align: center;
        font-size: 16px;
        line-height: 1.15;
        font-weight: 700;
        margin: 0 0 2px;
      }

      .arrival-report-subtitle {
        text-align: center;
        font-size: 11px;
        font-weight: 700;
        margin: 0 0 3px;
      }

      .arrival-report-date {
        text-align: right;
        font-size: 10px;
        font-weight: 700;
        margin: 0 0 2px;
      }

      .arrival-print-section {
        width: 100%;
        margin: 0;
        padding: 0;
        break-inside: avoid;
        page-break-inside: avoid;
      }

      .arrival-print-band {
        width: 100%;
        text-align: center;
        border: 1.5px solid #111;
        padding: 3px 1px;
        font-size: 10px;
        line-height: 1.05;
        font-weight: 700;
      }

      .arrival-report-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        margin: 0;
        font-size: 9px;
      }

      .arrival-report-table th,
      .arrival-report-table td {
        border: 1px solid #222;
        padding: 2px 1px;
        text-align: center;
        vertical-align: middle;
        line-height: 1.05;
        overflow-wrap: anywhere;
        word-break: normal;
      }

      .arrival-report-table th {
        background: #b9cde5 !important;
        font-weight: 700;
      }

      .arrival-report-table tfoot td {
        font-size: 9px;
        font-weight: 900;
        padding-top: 3px;
        padding-bottom: 3px;
      }

      .arrival-section-total td {
        font-size: 9px;
        font-weight: 900;
        border-top: 1.5px solid #111;
        border-bottom: 1.5px solid #111;
      }

      .arrival-blank-row td {
        height: 8px;
      }

      .arrival-grand-total-row td {
        font-size: 9px;
        font-weight: 900;
      }

      .arrival-summary-totals {
        width: 58%;
        margin: 0 auto;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 12px;
      }

      .arrival-summary-totals td,
      .arrival-summary-totals th {
        border: 1px solid #222;
        padding: 2px 4px;
        text-align: center;
        line-height: 1.05;
      }

      .arrival-summary-totals tr:nth-child(3) {
        border-top: 1.5px solid #111;
        border-bottom: 1.5px solid #111;
        font-weight: 700;
      }

      .arrival-summary-totals tr:nth-child(6) {
        font-weight: 700;
      }

      ${isSummary ? `
        .arrival-report-table th:nth-child(2),
        .arrival-report-table td:nth-child(2) { width: 15%; }

        .arrival-report-table th:nth-child(3),
        .arrival-report-table td:nth-child(3) { width: 12%; }

        .arrival-report-table th:nth-child(4),
        .arrival-report-table td:nth-child(4) { width: 5%; }

        .arrival-report-table th:nth-child(5),
        .arrival-report-table td:nth-child(5) { width: 22%; }

        .arrival-report-table th:nth-child(6),
        .arrival-report-table td:nth-child(6) { width: 7%; }

        .arrival-report-table th:nth-child(7),
        .arrival-report-table td:nth-child(7) { width: 10%; }

        .arrival-report-table th:nth-child(8),
        .arrival-report-table td:nth-child(8) { width: 14%; }

        .arrival-report-table th:nth-child(9),
        .arrival-report-table td:nth-child(9) { width: 12%; }
      ` : `
        .arrival-report-table th:nth-child(2),
        .arrival-report-table td:nth-child(2) { width: 10%; }

        .arrival-report-table th:nth-child(3),
        .arrival-report-table td:nth-child(3) { width: 10%; }

        .arrival-report-table th:nth-child(4),
        .arrival-report-table td:nth-child(4) { width: 11%; }

        .arrival-report-table th:nth-child(5),
        .arrival-report-table td:nth-child(5) { width: 5%; }

        .arrival-report-table th:nth-child(6),
        .arrival-report-table td:nth-child(6) { width: 21%; }

        .arrival-report-table th:nth-child(7),
        .arrival-report-table td:nth-child(7) { width: 6%; }

        .arrival-report-table th:nth-child(8),
        .arrival-report-table td:nth-child(8) { width: 7%; }

        .arrival-report-table th:nth-child(9),
        .arrival-report-table td:nth-child(9) { width: 12%; }

        .arrival-report-table th:nth-child(10),
        .arrival-report-table td:nth-child(10) { width: 15%; }
      `}

      @media print {
        html, body {
          width: 194mm;
          height: 281mm;
          margin: 0;
          padding: 0;
          overflow: hidden;
        }

        .arrival-fit-area {
          width: 194mm;
          height: 281mm;
          overflow: hidden;
        }

        .arrival-print-report,
        .arrival-print-section,
        table {
          break-inside: avoid !important;
          page-break-inside: avoid !important;
        }
      }
    `;

    win.document.open();

    win.document.write(`
      <!doctype html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>Arrival ${kind}</title>
          <style>${printCss}</style>
        </head>

        <body>
          <div class="arrival-fit-area">
            <div class="arrival-print-report" id="fit-content">
              ${target.innerHTML}
            </div>
          </div>

          <script>
            window.addEventListener("load", async function () {
              try {
                if (document.fonts && document.fonts.ready) {
                  await document.fonts.ready;
                }
              } catch(e) {}

              const area = document.querySelector(".arrival-fit-area");
              const content = document.getElementById("fit-content");

              content.style.transform = "none";

              const availableW = area.clientWidth;
              const availableH = area.clientHeight;

              const contentW = content.scrollWidth;
              const contentH = content.scrollHeight;

              // Fit complete report inside one A4 page.
              const scale = Math.min(
                1,
                availableW / contentW,
                availableH / contentH
              );

              content.style.transform = "scale(" + scale + ")";

              setTimeout(function () {
                window.focus();
                window.print();
              }, 500);
            });

          <\/script>
        </body>
      </html>
    `);

    win.document.close();
  }

  // --------------------------------------------------
  // REPORT COMPONENT
  // --------------------------------------------------

  const Report = ({ kind }) => {
    const summary = kind === "summary";

    return (
      <div
        className="arrival-print-report"
        id={`arrival-${kind}-print`}
      >
        <div className="arrival-report-title">
          Faisal Spinning Mills # 3 (Finishing Unit)
        </div>

        <div className="arrival-report-subtitle">
          Daily Greige Arrival {summary ? "Summary " : ""}Report
        </div>

        <div className="arrival-report-date">
          {dateLabel(reportDate)}
        </div>

        <ReportTable
          rows={rows}
          type="OWN"
          summary={summary}
          vehicles={ownVehicles}
        />

        <ReportTable
          rows={rows}
          type="CMT"
          summary={summary}
          vehicles={cmtVehicles}
        />

 {summary && (
          <table className="arrival-summary-totals">
            <tbody>
              <tr>
                <td>Own Mtr</td>
                <td>{ownMeters.toLocaleString()}</td>
                <td>{ownVehicles}</td>
                <td>VEHICLES</td>
              </tr>

              <tr>
                <td>Cmt Mtr</td>
                <td>{cmtMeters.toLocaleString()}</td>
                <td>{cmtVehicles}</td>
                <td>VEHICLES</td>
              </tr>

              <tr>
                <th>Total</th>
                <th>{totalMeters.toLocaleString()}</th>
                <th>{vehicles}</th>
                <th>VEHICLES</th>
              </tr>

              <tr>
                <td>Offloaded Mtr</td>
                <td>{totalMeters.toLocaleString()}</td>
                <td>{vehicles}</td>
                <td>VEHICLES</td>
              </tr>

              <tr>
                <td>Pending Mtr</td>
                <td>0</td>
                <td>0</td>
                <td>VEHICLES</td>
              </tr>

              <tr>
                <th></th>
                <th>{totalMeters.toLocaleString()}</th>
                <th>{vehicles}</th>
                <th>VEHICLES</th>
              </tr>
            </tbody>
          </table>
        )}
      </div>
    );
  };

  // --------------------------------------------------
  // MAIN UI
  // --------------------------------------------------

  return (
    <main className="arrival-management">
      <header className="arrival-header">
        <div>
          <div className="arrival-eyebrow">
            GREIGE OPERATIONS
          </div>

          <h1>ARRIVAL MANAGEMENT SYSTEM</h1>

          <p>Daily Arrival Stock, Unloading and Summary</p>
        </div>

        <button
          className="arrival-back"
          type="button"
          onClick={() => {
            if (section === "home" && onHome) {
              onHome();
            } else {
              setSection("home");
            }
          }}
        >
          ← Back to menu
        </button>
      </header>

      {section === "home" ? (
        <>
          <div className="arrival-panel">
            <h3>Upload Daily Arrival Excel (Single File)</h3>

            <div className="arrival-two-files">
              <label className="arrival-file-picker">
                <span className="arrival-file-picker__label">
                  One Excel File — STOCK + UNLOADING Sheets
                </span>

                <span className="arrival-file-picker__control">
                  <span className="arrival-file-picker__button">
                    {busy ? "Reading Excel…" : "Choose Excel file"}
                  </span>

                  <span className="arrival-file-picker__name">
                    {stockFile?.name || stockBook?.name || "No file selected"}
                  </span>

                  <input
                    type="file"
                    accept=".xlsx,.xls,.xlsm"
                    disabled={busy || cloudLoading}
                    onChange={(e) => {
                      const file = e.target.files?.[0] || null;
                      if (file) loadWorkbook(file);
                      e.target.value = "";
                    }}
                  />
                </span>
              </label>
            </div>

            <p className="arrival-help">
              Isi aik file ki STOCK sheet Arrival Stock mein aur UNLOADING sheet
              Unloading / Summary mein load hogi. Previous Day ki alag file ki
              zaroorat nahi.
            </p>

            {cloudLoading && (
              <div className="arrival-status">
                Loading shared Arrival data from Firebase…
              </div>
            )}

            {cloudSaving && (
              <div className="arrival-status">
                Saving Arrival data to Firebase…
              </div>
            )}

            {busy && (
              <div className="arrival-status">Reading workbook…</div>
            )}

            {error && (
              <div className="arrival-error" role="alert">
                {error}
              </div>
            )}

            {stockBook && (
              <div className="arrival-status">
                Workbook loaded: {stockRows.length} STOCK rows and {rows.length}
                UNLOADING rows.
              </div>
            )}
          </div>

        <section className="arrival-home">
          <button
            className="arrival-module-card arrival-module-card--stock"
            onClick={() => setSection("stock")}
            type="button"
          >
            <span className="arrival-card-icon">▤</span>
            <span className="arrival-card-title">ARRIVAL STOCK</span>
            <span className="arrival-card-description">
              View available stock cards from the uploaded workbook
            </span>
            <span className="arrival-card-open">Open section →</span>
          </button>

          <button
            className="arrival-module-card arrival-module-card--unloading"
            onClick={() => setSection("unloading")}
            type="button"
          >
            <span className="arrival-card-icon">▦</span>
            <span className="arrival-card-title">
              ARRIVAL UNLOADING / SUMMARY
            </span>
            <span className="arrival-card-description">
              Prepare unloading and matching summary reports
            </span>
            <span className="arrival-card-open">Open section →</span>
          </button>
        </section>
        </>
      ) : (
        <section className="arrival-workspace">
          <div className="arrival-section-heading">
            <div className="arrival-eyebrow">WORKSPACE</div>

            <h2>
              {section === "stock"
                ? "ARRIVAL STOCK"
                : "ARRIVAL UNLOADING / SUMMARY"}
            </h2>

            <p>
              {section === "stock"
                ? "Available stock entries from the STOCK sheet."
                : "Manage unloading rows once; both reports use the same data."}
            </p>
          </div>

          <div className="arrival-panel">
            <h3>Shared Arrival Workbook</h3>

            <div className="arrival-status">
              {stockBook?.name || "No workbook uploaded yet. Back to menu to choose one Excel file."}
            </div>

            {section === "unloading" && (
              <label className="arrival-date-field">
                <span>Report date</span>
                <input
                  type="date"
                  value={reportDate}
                  onChange={(e) => setReportDate(e.target.value)}
                />
              </label>
            )}

            {cloudLoading && (
              <div className="arrival-status">
                Loading shared Arrival data from Firebase…
              </div>
            )}

            {cloudSaving && (
              <div className="arrival-status">
                Saving Arrival data to Firebase…
              </div>
            )}

            {busy && (
              <div className="arrival-status">Reading workbook…</div>
            )}

            {error && (
              <div className="arrival-error" role="alert">
                {error}
              </div>
            )}

            {stockBook && (
              <div className="arrival-status">
                {stockRows.length} STOCK rows loaded; {rows.length} UNLOADING rows loaded.
              </div>
            )}
          </div>

          {section === "stock" && (
            <div className="arrival-panel">
              <div
                className="arrival-stock-total-summary"
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))",
                  gap: "12px",
                  marginBottom: "16px",
                }}
              >
                {[
                  { label: "OWN TOTAL METERS", value: stockMeterTotals.own, color: "#1d4ed8", bg: "#eff6ff" },
                  { label: "COMMERCIAL TOTAL METERS", value: stockMeterTotals.commercial, color: "#c2410c", bg: "#fff7ed" },
                  { label: "GRAND TOTAL METERS", value: stockMeterTotals.total, color: "#166534", bg: "#f0fdf4" },
                ].map((item) => (
                  <div
                    key={item.label}
                    style={{
                      padding: "15px 16px",
                      borderRadius: "12px",
                      border: `2px solid ${item.color}33`,
                      background: item.bg,
                      minWidth: 0,
                    }}
                  >
                    <div style={{ color: item.color, fontSize: "0.78rem", fontWeight: 900, letterSpacing: "0.03em" }}>
                      {item.label}
                    </div>
                    <div style={{ color: item.color, fontSize: "1.55rem", fontWeight: 900, marginTop: "5px" }}>
                      {item.value.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                    </div>
                  </div>
                ))}
              </div>

              <div style={{ marginBottom: "18px" }}>
                <button
                  type="button"
                  className="arrival-action arrival-action--primary"
                  onClick={() => setShowStockDateWiseTotals((value) => !value)}
                  aria-expanded={showStockDateWiseTotals}
                >
                  {showStockDateWiseTotals ? "Hide Date Wise Total" : "Date Wise Total"}
                </button>
              </div>

              {showStockDateWiseTotals && (
                <div
                  className="arrival-stock-date-total-wrap"
                  style={{
                    overflowX: "auto",
                    marginBottom: "20px",
                    border: "1px solid #cbd5e1",
                    borderRadius: "12px",
                  }}
                >
                  <table style={{ width: "100%", minWidth: "560px", borderCollapse: "collapse", background: "#ffffff" }}>
                    <thead>
                      <tr style={{ background: "#334155", color: "#ffffff" }}>
                        {["Date", "OWN Meters", "Commercial Meters", "Total Meters"].map((heading) => (
                          <th key={heading} style={{ padding: "11px 12px", textAlign: heading === "Date" ? "left" : "right", whiteSpace: "nowrap" }}>
                            {heading}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {stockDateWiseTotals.map((item) => (
                        <tr key={item.dateKey} style={{ borderTop: "1px solid #e2e8f0" }}>
                          <td style={{ padding: "10px 12px", fontWeight: 800 }}>
                            {item.dateKey === "undated" ? "Date Not Available" : formatStockDate(item.dateKey)}
                          </td>
                          <td style={{ padding: "10px 12px", textAlign: "right", color: "#1d4ed8", fontWeight: 700 }}>
                            {item.own.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                          </td>
                          <td style={{ padding: "10px 12px", textAlign: "right", color: "#c2410c", fontWeight: 700 }}>
                            {item.commercial.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                          </td>
                          <td style={{ padding: "10px 12px", textAlign: "right", color: "#166534", fontWeight: 900 }}>
                            {item.total.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr style={{ background: "#f1f5f9", borderTop: "2px solid #94a3b8" }}>
                        <td style={{ padding: "11px 12px", fontWeight: 900 }}>Grand Total</td>
                        <td style={{ padding: "11px 12px", textAlign: "right", color: "#1d4ed8", fontWeight: 900 }}>
                          {stockMeterTotals.own.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                        </td>
                        <td style={{ padding: "11px 12px", textAlign: "right", color: "#c2410c", fontWeight: 900 }}>
                          {stockMeterTotals.commercial.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                        </td>
                        <td style={{ padding: "11px 12px", textAlign: "right", color: "#166534", fontWeight: 900 }}>
                          {stockMeterTotals.total.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              )}

              <h3>
                Available Stock Cards ({filteredStockRows.length} / {stockRows.length})
              </h3>

              <div className="arrival-stock-filters arrival-manual-form">
                <label>
                  Date
                  <input
                    type="date"
                    value={stockFilterDraft.date}
                    onChange={(e) =>
                      setStockFilterDraft((old) => ({
                        ...old,
                        date: e.target.value,
                      }))
                    }
                  />
                </label>

                <label>
                  Quality
                  <input
                    type="text"
                    placeholder="Quality ka koi hissa likhein"
                    value={stockFilterDraft.quality}
                    onChange={(e) =>
                      setStockFilterDraft((old) => ({
                        ...old,
                        quality: e.target.value,
                      }))
                    }
                  />
                </label>

                <label>
                  Width
                  <input
                    type="text"
                    placeholder="Width ka koi hissa likhein"
                    value={stockFilterDraft.width}
                    onChange={(e) =>
                      setStockFilterDraft((old) => ({
                        ...old,
                        width: e.target.value,
                      }))
                    }
                  />
                </label>

                <label>
                  Supplier
                  <input
                    type="text"
                    placeholder="Supplier ka koi hissa likhein"
                    value={stockFilterDraft.supplier}
                    onChange={(e) =>
                      setStockFilterDraft((old) => ({
                        ...old,
                        supplier: e.target.value,
                      }))
                    }
                  />
                </label>

                <label>
                  Vehicle #
                  <input
                    type="text"
                    placeholder="Vehicle number ka koi hissa likhein"
                    value={stockFilterDraft.vehicle}
                    onChange={(e) =>
                      setStockFilterDraft((old) => ({
                        ...old,
                        vehicle: e.target.value,
                      }))
                    }
                  />
                </label>

                <div className="arrival-stock-filter-actions">
                  <button
                    type="button"
                    className="arrival-action arrival-action--primary"
                    onClick={applyStockFilters}
                  >
                    Search
                  </button>

                  <button
                    type="button"
                    className="arrival-action"
                    onClick={clearStockFilters}
                  >
                    Clear Filters
                  </button>
                </div>
              </div>

              {filteredStockRows.length === 0 ? (
                <div className="arrival-status">
                  Koi matching stock record nahi mila.
                </div>
              ) : (
                <div className="arrival-stock-date-groups">
                  {filteredStockGroups.map((group) => {
                    const isLatestDate =
                      group.dateKey !== "undated" &&
                      group.dateKey === latestStockDate;

                    return (
                      <div
                        className={`arrival-panel arrival-stock-date-group${
                          isLatestDate ? " arrival-stock-date-group--latest" : ""
                        }`}
                        key={group.dateKey}
                        style={{
                          background: isLatestDate ? "#fffbeb" : "#ffffff",
                          border: isLatestDate
                            ? "3px solid #d97706"
                            : "1px solid #d6dce5",
                          borderRadius: "16px",
                          padding: isLatestDate ? "18px" : "14px",
                          marginBottom: "22px",
                          boxShadow: isLatestDate
                            ? "0 0 0 4px rgba(245, 158, 11, 0.16), 0 8px 22px rgba(146, 64, 14, 0.12)"
                            : "0 3px 12px rgba(15, 23, 42, 0.05)",
                        }}
                      >
                        <div
                          className="arrival-stock-date-heading-wrap"
                          style={{
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "space-between",
                            flexWrap: "wrap",
                            gap: "10px",
                            padding: "10px 14px",
                            marginBottom: "14px",
                            borderRadius: "10px",
                            background: isLatestDate ? "#78350f" : "#334155",
                            color: "#ffffff",
                            borderLeft: isLatestDate ? "6px solid #fbbf24" : "6px solid #94a3b8",
                          }}
                        >
                          <h3
                            className="arrival-stock-date-heading"
                            style={{
                              color: "#ffffff",
                              margin: 0,
                              fontSize: "1.15rem",
                              fontWeight: 800,
                            }}
                          >
                            {group.label}
                          </h3>

                          {isLatestDate && (
                            <span
                              className="arrival-stock-latest-badge"
                              style={{
                                display: "inline-flex",
                                alignItems: "center",
                                gap: "6px",
                                padding: "6px 11px",
                                borderRadius: "999px",
                                background: "#fbbf24",
                                color: "#451a03",
                                fontSize: "0.75rem",
                                fontWeight: 900,
                                letterSpacing: "0.04em",
                              }}
                            >
                              ★ LATEST ARRIVAL
                            </span>
                          )}
                        </div>

                        {group.ownRows.length > 0 && (
                          <section
                            className="arrival-stock-type-group arrival-stock-type-group--own"
                            style={{
                              background: "#eff6ff",
                              border: "2px solid #93c5fd",
                              borderRadius: "12px",
                              padding: "12px",
                              marginBottom: group.cmtRows.length > 0 ? "14px" : 0,
                            }}
                          >
                            <h4
                              className="arrival-stock-type-heading"
                              style={{
                                color: "#1d4ed8",
                                background: "#dbeafe",
                                borderLeft: "5px solid #2563eb",
                                borderRadius: "7px",
                                padding: "9px 12px",
                                margin: "0 0 12px",
                                fontWeight: 900,
                              }}
                            >
                              OWN ARRIVAL
                            </h4>
                            <div className="arrival-card-grid">
                              {group.ownRows.map((r) => (
                                <article
                                  className="arrival-data-card arrival-stock-card--own"
                                  key={r.stockId}
                                  style={{
                                    background: "#f8fbff",
                                    border: "1px solid #bfdbfe",
                                    borderLeft: "5px solid #2563eb",
                                  }}
                                >
                                  <div className="arrival-data-card__top">
                                    <b style={{ color: "#1d4ed8" }}>OWN</b>
                                    <strong>{r.meter} Mtr</strong>
                                  </div>

                                  {FIELD_LABELS.filter(([key]) => key !== "customer").map(([key, label]) => (
                                    <div className="arrival-data-field" key={key}>
                                      <span>{label}</span>
                                      <b>{r[key] || "—"}</b>
                                    </div>
                                  ))}
                                </article>
                              ))}
                            </div>
                          </section>
                        )}

                        {group.cmtRows.length > 0 && (
                          <section
                            className="arrival-stock-type-group arrival-stock-type-group--commercial"
                            style={{
                              background: "#fff7ed",
                              border: "2px solid #fdba74",
                              borderRadius: "12px",
                              padding: "12px",
                            }}
                          >
                            <h4
                              className="arrival-stock-type-heading"
                              style={{
                                color: "#c2410c",
                                background: "#ffedd5",
                                borderLeft: "5px solid #ea580c",
                                borderRadius: "7px",
                                padding: "9px 12px",
                                margin: "0 0 12px",
                                fontWeight: 900,
                              }}
                            >
                              COMMERCIAL ARRIVAL
                            </h4>
                            <div className="arrival-card-grid">
                              {group.cmtRows.map((r) => (
                                <article
                                  className="arrival-data-card arrival-stock-card--commercial"
                                  key={r.stockId}
                                  style={{
                                    background: "#fffaf5",
                                    border: "1px solid #fed7aa",
                                    borderLeft: "5px solid #ea580c",
                                  }}
                                >
                                  <div className="arrival-data-card__top">
                                    <b style={{ color: "#c2410c" }}>COMMERCIAL</b>
                                    <strong>{r.meter} Mtr</strong>
                                  </div>

                                  {FIELD_LABELS.filter(([key]) => key !== "customer").map(([key, label]) => (
                                    <div className="arrival-data-field" key={key}>
                                      <span>{label}</span>
                                      <b>{r[key] || "—"}</b>
                                    </div>
                                  ))}
                                </article>
                              ))}
                            </div>
                          </section>
                        )}

                        {isLatestDate && (
                          <div
                            className="arrival-stock-latest-end"
                            style={{
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
                              gap: "9px",
                              marginTop: "16px",
                              padding: "9px 12px",
                              borderTop: "2px dashed #d97706",
                              color: "#92400e",
                              fontWeight: 900,
                              fontSize: "0.78rem",
                              letterSpacing: "0.06em",
                            }}
                          >
                            <span>END OF LATEST ARRIVAL SECTION</span>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}

          {section === "unloading" && (
            <>
              <div className="arrival-toolbar">
                <button
                  type="button"
                  onClick={() => setShowAuto((v) => !v)}
                  className="arrival-action arrival-action--primary"
                >
                  {showAuto ? "Close Auto Add" : "Auto Add from Stock"}
                </button>

                <button
                  type="button"
                  className="arrival-action"
                  onClick={() =>
                    document
                      .getElementById("arrival-manual-form")
                      ?.scrollIntoView({ behavior: "smooth" })
                  }
                >
                  + Manual Add
                </button>

                <button
                  type="button"
                  className="arrival-action"
                  onClick={() => printReport("unloading")}
                >
                  Print / Save PDF — Unloading
                </button>

                <button
                  type="button"
                  className="arrival-action"
                  onClick={() => exportImage("unloading")}
                >
                  Save Unloading Image
                </button>

                <button
                  type="button"
                  className="arrival-action"
                  onClick={() => printReport("summary")}
                >
                  Print / Save PDF — Summary
                </button>

                <button
                  type="button"
                  className="arrival-action"
                  onClick={() => exportImage("summary")}
                >
                  Save Summary Image
                </button>
              </div>

              {showAuto && (
                <div className="arrival-panel">
                  <h3>
                    Available Stock — select cards to add ({availableRows.length})
                  </h3>

                  <p className="arrival-help">
                    Rows already present in the workbook's UNLOADING sheet
                    are excluded using their stock details and vehicle.
                  </p>

                  <button
                    type="button"
                    className="arrival-action arrival-action--primary"
                    disabled={!selected.length}
                    onClick={addSelected}
                    style={{
                      marginBottom: "14px",
                      fontWeight: 900,
                      minHeight: "44px",
                      width: "100%",
                      position: "sticky",
                      bottom: "10px",
                      zIndex: 20,
                      boxShadow: "0 4px 14px rgba(0,0,0,0.22)",
                    }}
                  >
                    Add Selected to Unloading Sheet ({selected.length})
                  </button>

                  {availableDateGroups.length === 0 ? (
                    <div className="arrival-status">
                      Koi available stock record nahi mila.
                    </div>
                  ) : (
                    <div className="arrival-stock-date-groups">
                      {availableDateGroups.map((group) => {
                        const isLatestDate =
                          group.dateKey !== "undated" &&
                          group.dateKey === latestAvailableDate;

                        return (
                          <div
                            className={`arrival-panel arrival-stock-date-group${
                              isLatestDate ? " arrival-stock-date-group--latest" : ""
                            }`}
                            key={group.dateKey}
                            style={{
                              background: isLatestDate ? "#fffbeb" : "#ffffff",
                              border: isLatestDate ? "3px solid #d97706" : "1px solid #d6dce5",
                              borderRadius: "16px",
                              padding: isLatestDate ? "18px" : "14px",
                              marginBottom: "22px",
                              boxShadow: isLatestDate
                                ? "0 0 0 4px rgba(245, 158, 11, 0.16), 0 8px 22px rgba(146, 64, 14, 0.12)"
                                : "0 3px 12px rgba(15, 23, 42, 0.05)",
                            }}
                          >
                            <div
                              className="arrival-stock-date-heading-wrap"
                              style={{
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "space-between",
                                flexWrap: "wrap",
                                gap: "10px",
                                padding: "10px 14px",
                                marginBottom: "14px",
                                borderRadius: "10px",
                                background: isLatestDate ? "#78350f" : "#334155",
                                color: "#ffffff",
                                borderLeft: isLatestDate ? "6px solid #fbbf24" : "6px solid #94a3b8",
                              }}
                            >
                              <h3
                                className="arrival-stock-date-heading"
                                style={{
                                  color: "#ffffff",
                                  margin: 0,
                                  fontSize: "1.15rem",
                                  fontWeight: 800,
                                }}
                              >
                                {group.label}
                              </h3>

                              {isLatestDate && (
                                <span
                                  className="arrival-stock-latest-badge"
                                  style={{
                                    display: "inline-flex",
                                    alignItems: "center",
                                    gap: "6px",
                                    padding: "6px 11px",
                                    borderRadius: "999px",
                                    background: "#fbbf24",
                                    color: "#451a03",
                                    fontSize: "0.75rem",
                                    fontWeight: 900,
                                    letterSpacing: "0.04em",
                                  }}
                                >
                                  ★ LATEST ARRIVAL
                                </span>
                              )}
                            </div>

                            {group.ownRows.length > 0 && (
                              <section
                                className="arrival-stock-type-group arrival-stock-type-group--own"
                                style={{
                                  background: "#eff6ff",
                                  border: "2px solid #93c5fd",
                                  borderRadius: "12px",
                                  padding: "12px",
                                  marginBottom: group.cmtRows.length > 0 ? "14px" : 0,
                                }}
                              >
                                <h4
                                  className="arrival-stock-type-heading"
                                  style={{
                                    color: "#1d4ed8",
                                    background: "#dbeafe",
                                    borderLeft: "5px solid #2563eb",
                                    borderRadius: "7px",
                                    padding: "9px 12px",
                                    margin: "0 0 12px",
                                    fontWeight: 900,
                                  }}
                                >
                                  OWN ARRIVAL
                                </h4>
                                <div className="arrival-card-grid">
                                  {group.ownRows.map((r) => (
                                    <label
                                      className="arrival-data-card arrival-select-card arrival-stock-card--own"
                                      key={r.stockId}
                                      style={{
                                        background: "#f8fbff",
                                        border: "1px solid #bfdbfe",
                                        borderLeft: "5px solid #2563eb",
                                      }}
                                    >
                                      <div className="arrival-data-card__top">
                                        <b style={{ color: "#1d4ed8" }}>OWN</b>
                                        <input
                                          type="checkbox"
                                          checked={selected.includes(r.stockId)}
                                          onChange={(e) =>
                                            setSelected((old) =>
                                              e.target.checked
                                                ? [...old, r.stockId]
                                                : old.filter((x) => x !== r.stockId)
                                            )
                                          }
                                        />
                                      </div>
                                      {FIELD_LABELS.filter(([k]) => k !== "customer").map(([k, label]) => (
                                        <div className="arrival-data-field" key={k}>
                                          <span>{label}</span>
                                          <b>{r[k] || "—"}</b>
                                        </div>
                                      ))}
                                    </label>
                                  ))}
                                </div>
                              </section>
                            )}

                            {group.cmtRows.length > 0 && (
                              <section
                                className="arrival-stock-type-group arrival-stock-type-group--commercial"
                                style={{
                                  background: "#fff7ed",
                                  border: "2px solid #fdba74",
                                  borderRadius: "12px",
                                  padding: "12px",
                                }}
                              >
                                <h4
                                  className="arrival-stock-type-heading"
                                  style={{
                                    color: "#c2410c",
                                    background: "#ffedd5",
                                    borderLeft: "5px solid #ea580c",
                                    borderRadius: "7px",
                                    padding: "9px 12px",
                                    margin: "0 0 12px",
                                    fontWeight: 900,
                                  }}
                                >
                                  COMMERCIAL ARRIVAL
                                </h4>
                                <div className="arrival-card-grid">
                                  {group.cmtRows.map((r) => (
                                    <label
                                      className="arrival-data-card arrival-select-card arrival-stock-card--commercial"
                                      key={r.stockId}
                                      style={{
                                        background: "#fffaf5",
                                        border: "1px solid #fed7aa",
                                        borderLeft: "5px solid #ea580c",
                                      }}
                                    >
                                      <div className="arrival-data-card__top">
                                        <b style={{ color: "#c2410c" }}>COMMERCIAL</b>
                                        <input
                                          type="checkbox"
                                          checked={selected.includes(r.stockId)}
                                          onChange={(e) =>
                                            setSelected((old) =>
                                              e.target.checked
                                                ? [...old, r.stockId]
                                                : old.filter((x) => x !== r.stockId)
                                            )
                                          }
                                        />
                                      </div>
                                      {FIELD_LABELS.filter(([k]) => k !== "customer").map(([k, label]) => (
                                        <div className="arrival-data-field" key={k}>
                                          <span>{label}</span>
                                          <b>{r[k] || "—"}</b>
                                        </div>
                                      ))}
                                    </label>
                                  ))}
                                </div>
                              </section>
                            )}

                            {isLatestDate && (
                              <div
                                className="arrival-stock-latest-end"
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  justifyContent: "center",
                                  marginTop: "16px",
                                  padding: "9px 12px",
                                  borderTop: "2px dashed #d97706",
                                  color: "#92400e",
                                  fontWeight: 900,
                                  fontSize: "0.78rem",
                                  letterSpacing: "0.06em",
                                }}
                              >
                                END OF LATEST ARRIVAL SECTION
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}

                </div>
              )}

              <div className="arrival-summary-strip">
                <div>
                  <span>Total Rows</span>
                  <b>{rows.length}</b>
                </div>

                <div>
                  <span>Own Meters</span>
                  <b>{ownMeters.toLocaleString()}</b>
                </div>

                <div>
                  <span>Commercial Meters</span>
                  <b>{cmtMeters.toLocaleString()}</b>
                </div>

                <div>
                  <span>Total Meters</span>
                  <b>{totalMeters.toLocaleString()}</b>
                </div>

                <div>
                  <span>Vehicles</span>
                  <b>{vehicles}</b>
                </div>
              </div>

              {/* MANUAL VEHICLE COUNT */}
              <div className="arrival-panel">
                <h3>Vehicle Count Adjustment</h3>

                <p className="arrival-help">
                  Automatic vehicle count neeche show ho raha hai.
                  Agar count manually change karna ho to number enter karein.
                  Field blank karne par automatic count use hoga.
                </p>

                <div className="arrival-manual-form">
                  <label>
                    Own Vehicles (Auto: {autoOwnVehicles})
                    <input
                      type="number"
                      min="0"
                      step="1"
                      placeholder={`Automatic: ${autoOwnVehicles}`}
                      value={manualOwnVehicles}
                      onChange={(e) =>
                        setManualOwnVehicles(e.target.value)
                      }
                    />
                  </label>

                  <label>
                    Commercial Vehicles (Auto: {autoCmtVehicles})
                    <input
                      type="number"
                      min="0"
                      step="1"
                      placeholder={`Automatic: ${autoCmtVehicles}`}
                      value={manualCmtVehicles}
                      onChange={(e) =>
                        setManualCmtVehicles(e.target.value)
                      }
                    />
                  </label>

                  <div>
                    <b>Grand Total Vehicles: {vehicles}</b>
                  </div>
                </div>
              </div>

              {/* MANUAL ADD */}
              <div className="arrival-panel">
                <h3>Manual Add Entry</h3>

                <form
                  id="arrival-manual-form"
                  className="arrival-manual-form"
                  onSubmit={addManual}
                >
                  <label>
                    Type
                    <select
                      value={manual.type}
                      onChange={(e) =>
                        setManual({
                          ...manual,
                          type: e.target.value,
                        })
                      }
                    >
                      <option value="OWN">Own</option>
                      <option value="CMT">Commercial</option>
                    </select>
                  </label>

                  {FIELD_LABELS.map(([key, label]) => (
                    <label key={key}>
                      {label}
                      <input
                        value={manual[key]}
                        onChange={(e) =>
                          setManual({
                            ...manual,
                            [key]: e.target.value,
                          })
                        }
                      />
                    </label>
                  ))}

                  <button
                    className="arrival-action arrival-action--primary"
                    type="submit"
                  >
                    Add Row
                  </button>
                </form>
              </div>

              {/* UNLOADING CARDS */}
              <div className="arrival-panel">
                <h3>Unloading Cards ({rows.length})</h3>

                {unloadingDateGroups.length === 0 ? (
                  <div className="arrival-status">
                    Abhi Unloading mein koi card available nahi hai.
                  </div>
                ) : (
                  <div className="arrival-stock-date-groups">
                    {unloadingDateGroups.map((group) => {
                      const isLatestDate =
                        group.dateKey !== "undated" &&
                        group.dateKey === latestUnloadingDate;

                      return (
                        <div
                          className={`arrival-panel arrival-stock-date-group${
                            isLatestDate ? " arrival-stock-date-group--latest" : ""
                          }`}
                          key={group.dateKey}
                          style={{
                            background: isLatestDate ? "#fffbeb" : "#ffffff",
                            border: isLatestDate ? "3px solid #d97706" : "1px solid #d6dce5",
                            borderRadius: "16px",
                            padding: isLatestDate ? "18px" : "14px",
                            marginBottom: "22px",
                            boxShadow: isLatestDate
                              ? "0 0 0 4px rgba(245, 158, 11, 0.16), 0 8px 22px rgba(146, 64, 14, 0.12)"
                              : "0 3px 12px rgba(15, 23, 42, 0.05)",
                          }}
                        >


                          {group.ownRows.length > 0 && (
                            <section
                              className="arrival-stock-type-group arrival-stock-type-group--own"
                              style={{
                                background: "#eff6ff",
                                border: "2px solid #93c5fd",
                                borderRadius: "12px",
                                padding: "12px",
                                marginBottom: group.cmtRows.length > 0 ? "14px" : 0,
                              }}
                            >
                              <h4
                                className="arrival-stock-type-heading"
                                style={{
                                  color: "#1d4ed8",
                                  background: "#dbeafe",
                                  borderLeft: "5px solid #2563eb",
                                  borderRadius: "7px",
                                  padding: "9px 12px",
                                  margin: "0 0 12px",
                                  fontWeight: 900,
                                }}
                              >
                                OWN UNLOADING
                              </h4>
                              <div className="arrival-card-grid">
                                {group.ownRows.map((r) => (
                                  <article
                                    className="arrival-data-card arrival-stock-card--own"
                                    key={r.id}
                                    style={{
                                      background: "#f8fbff",
                                      border: "1px solid #bfdbfe",
                                      borderLeft: "5px solid #2563eb",
                                    }}
                                  >
                                    <div className="arrival-data-card__top">
                                      <b style={{ color: "#1d4ed8" }}>OWN · {r.supplier}</b>
                                      <button type="button" onClick={() => removeRow(r.id)}>Remove</button>
                                    </div>
                                    {FIELD_LABELS.map(([k, label]) => (
                                      <label className="arrival-data-field arrival-edit-field" key={k}>
                                        <span>{label}</span>
                                        <input
                                          value={r[k] || ""}
                                          onChange={(e) => updateRow(r.id, k, e.target.value)}
                                        />
                                      </label>
                                    ))}
                                  </article>
                                ))}
                              </div>
                            </section>
                          )}

                          {group.cmtRows.length > 0 && (
                            <section
                              className="arrival-stock-type-group arrival-stock-type-group--commercial"
                              style={{
                                background: "#fff7ed",
                                border: "2px solid #fdba74",
                                borderRadius: "12px",
                                padding: "12px",
                              }}
                            >
                              <h4
                                className="arrival-stock-type-heading"
                                style={{
                                  color: "#c2410c",
                                  background: "#ffedd5",
                                  borderLeft: "5px solid #ea580c",
                                  borderRadius: "7px",
                                  padding: "9px 12px",
                                  margin: "0 0 12px",
                                  fontWeight: 900,
                                }}
                              >
                                COMMERCIAL UNLOADING
                              </h4>
                              <div className="arrival-card-grid">
                                {group.cmtRows.map((r) => (
                                  <article
                                    className="arrival-data-card arrival-stock-card--commercial"
                                    key={r.id}
                                    style={{
                                      background: "#fffaf5",
                                      border: "1px solid #fed7aa",
                                      borderLeft: "5px solid #ea580c",
                                    }}
                                  >
                                    <div className="arrival-data-card__top">
                                      <b style={{ color: "#c2410c" }}>COMMERCIAL · {r.supplier}</b>
                                      <button type="button" onClick={() => removeRow(r.id)}>Remove</button>
                                    </div>
                                    {FIELD_LABELS.map(([k, label]) => (
                                      <label className="arrival-data-field arrival-edit-field" key={k}>
                                        <span>{label}</span>
                                        <input
                                          value={r[k] || ""}
                                          onChange={(e) => updateRow(r.id, k, e.target.value)}
                                        />
                                      </label>
                                    ))}
                                  </article>
                                ))}
                              </div>
                            </section>
                          )}

                          {isLatestDate && (
                            <div
                              className="arrival-stock-latest-end"
                              style={{
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                marginTop: "16px",
                                padding: "9px 12px",
                                borderTop: "2px dashed #d97706",
                                color: "#92400e",
                                fontWeight: 900,
                                fontSize: "0.78rem",
                                letterSpacing: "0.06em",
                              }}
                            >
                              END OF LATEST ARRIVAL SECTION
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* PRINT PREVIEW */}
              <div className="arrival-panel">
                <h3>Print Preview — Reports</h3>

                <div className="arrival-report-preview">
                  <Report kind="unloading" />
                  <Report kind="summary" />
                </div>
              </div>
            </>
          )}
        </section>
      )}
    </main>
  );
}
