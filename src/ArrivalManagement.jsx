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

  // Common Excel date format in Pakistan: DD/MM/YYYY or DD-MM-YYYY.
  match = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
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

function parseStockRows(rows) {
  const out = [];
  let type = "OWN";
  let headers = null;

  rows.forEach((row) => {
    const first = clean(row?.[0]).toUpperCase();

    if (first === "OWN") {
      type = "OWN";
      headers = null;
      return;
    }

    if (first === "CMT" || first === "COMMERCIAL") {
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
      date: clean(get("date", "arrivaldate", "stockdate", "receiveddate", "receivingdate", "receiptdate", "dateofreceipt", "dateofarrival", "inwarddate", "challandate", "entrydate", "dated")),
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

        wb.SheetNames.forEach((name) => {
          sheets[name.toUpperCase()] = XLSX.utils.sheet_to_json(
            wb.Sheets[name],
            {
              header: 1,
              defval: "",
              raw: false,
            }
          );
        });

        resolve({ name: file.name, sheets });
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
  return `${Number(m)}/${Number(d)}/${y}`;
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
            <td colSpan={7}>
              {type === "OWN" ? "Own Total" : "Cmt Total"}
            </td>
            <td>{meterTotal.toLocaleString()}</td>
            <td colSpan={2}>
              {vehicles} VEHICLES
            </td>
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

  const [stockFile, setStockFile] = useState(null);
  const [previousFile, setPreviousFile] = useState(null);

  const [stockBook, setStockBook] = useState(null);
  const [previousBook, setPreviousBook] = useState(null);

  const [stockRows, setStockRows] = useState([]);
  const [rows, setRows] = useState([]);

  const [reportDate, setReportDate] = useState(
    new Date().toLocaleDateString("en-CA")
  );

  const [showAuto, setShowAuto] = useState(false);
  const [selected, setSelected] = useState([]);

  // Stock filters apply only when Search is pressed.
  const EMPTY_STOCK_FILTERS = { date: "", quality: "", supplier: "" };
  const [stockFilterDraft, setStockFilterDraft] = useState({ ...EMPTY_STOCK_FILTERS });
  const [stockFilters, setStockFilters] = useState({ ...EMPTY_STOCK_FILTERS });

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
    const qualityFilter = stockFilters.quality.trim().toLowerCase();
    const supplierFilter = stockFilters.supplier.trim().toLowerCase();

    return stockRows.filter((r) => {
      if (
        dateFilter &&
        normalizeDateForFilter(r.date) !== dateFilter
      ) {
        return false;
      }

      if (
        qualityFilter &&
        !clean(r.quality).toLowerCase().includes(qualityFilter)
      ) {
        return false;
      }

      if (
        supplierFilter &&
        !clean(r.supplier).toLowerCase().includes(supplierFilter)
      ) {
        return false;
      }

      return true;
    });
  }, [stockRows, stockFilters]);

  function applyStockFilters() {
    setStockFilters({
      date: stockFilterDraft.date,
      quality: stockFilterDraft.quality,
      supplier: stockFilterDraft.supplier,
    });
  }

  function clearStockFilters() {
    const empty = { date: "", quality: "", supplier: "" };
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
            previousFileName: meta.previousFileName || "",
          };

          lastAppliedRevisionRef.current = revision;
          lastPersistedPayloadRef.current = JSON.stringify(remotePayload);

          setStockRows(remotePayload.stockRows);
          setRows(remotePayload.rows);
          setReportDate(remotePayload.reportDate);
          setManualOwnVehicles(remotePayload.manualOwnVehicles);
          setManualCmtVehicles(remotePayload.manualCmtVehicles);
          setStockFile(null);
          setPreviousFile(null);
          setStockBook(
            remotePayload.stockFileName
              ? { name: remotePayload.stockFileName }
              : null
          );
          setPreviousBook(
            remotePayload.previousFileName
              ? { name: remotePayload.previousFileName }
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
      previousFileName: previousBook?.name || previousFile?.name || "",
    }),
    [
      stockRows,
      rows,
      reportDate,
      manualOwnVehicles,
      manualCmtVehicles,
      stockBook?.name,
      stockFile?.name,
      previousBook?.name,
      previousFile?.name,
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
  // LOAD STOCK
  // --------------------------------------------------

  async function loadStock(file) {
    if (!file) return;

    setError("");

    try {
      setBusy(true);

      const book = await parseWorkbook(file);
      const data = book.sheets.STOCK || [];
      const parsed = parseStockRows(data);

      // Invalid / empty workbook se pehle wala saved stock replace nahi hoga.
      if (!parsed.length) {
        setError(
          "STOCK sheet mein rows nahi milin. Sheet ke OWN/CMT headings aur Supplier header check karein. Purana saved stock safe hai."
        );
        return;
      }

      setStockFile(file);
      setStockBook(book);
      setStockRows(parsed);

      if (!previousFile && !previousBook && book.sheets.UNLOADING) {
        const prev = parseSections(
          book.sheets.UNLOADING,
          "previous"
        );

        if (prev.length) {
          setPreviousBook({
            name: `${file.name} — UNLOADING sheet`,
          });

          setRows(
            prev.map((r, i) => ({
              ...r,
              id: `previous-${i}`,
            }))
          );
        }
      }
    } catch (e) {
      setError(
        e.message || "Stock file could not be loaded."
      );
    } finally {
      setBusy(false);
    }
  }

  // --------------------------------------------------
  // LOAD PREVIOUS UNLOADING
  // --------------------------------------------------

  async function loadPrevious(file) {
    if (!file) return;

    setError("");

    try {
      setBusy(true);

      const book = await parseWorkbook(file);

      const data =
        book.sheets.UNLOADING ||
        book.sheets.SUMMARY ||
        book.sheets[Object.keys(book.sheets)[0]] ||
        [];

      const parsed = parseSections(data, "previous");

      if (!parsed.length) {
        setError(
          "Previous unloading file mein valid rows nahi milin. Pehle se saved unloading data safe hai."
        );
        return;
      }

      setPreviousFile(file);
      setPreviousBook(book);

      setRows(
        parsed.map((r, i) => ({
          ...r,
          id: `previous-${i}`,
        }))
      );
    } catch (e) {
      setError(
        e.message || "Previous unloading file could not be loaded."
      );
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

      clone.style.width = "max-content";
      clone.style.minWidth = "1100px";
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
        table.style.tableLayout = "auto";
        table.style.width = "max-content";
        table.style.minWidth = "100%";
        table.style.maxWidth = "none";
        table.style.borderCollapse = "collapse";

        table.querySelectorAll("th, td").forEach((cell) => {
          cell.style.whiteSpace = "nowrap";
          cell.style.wordBreak = "normal";
          cell.style.overflowWrap = "normal";
          cell.style.width = "auto";
          cell.style.height = "auto";
          cell.style.padding = "4px 8px";
          cell.style.lineHeight = "1.2";
          cell.style.fontSize = "11px";
        });

        // Summary aur Unloading mein Weave column ki position alag hai
        const isSummary = kind === "summary";

        const widthColumn = isSummary ? 4 : 5;
        const weaveColumn = isSummary ? 5 : 6;
        const brtColumn = isSummary ? 6 : 7;
        const meterColumn = isSummary ? 7 : 8;

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

      clone.querySelectorAll("tr").forEach((row) => {
        row.style.height = "auto";
      });

      const canvas = await html2canvas(clone, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        logging: false,
        scrollX: 0,
        scrollY: 0,
        windowWidth: Math.max(1200, clone.scrollWidth),
        windowHeight: clone.scrollHeight,
      });

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
        font-size: 7px;
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
        font-weight: 700;
        padding-top: 3px;
        padding-bottom: 3px;
      }

      .arrival-section-total td {
        border-top: 1.5px solid #111;
        border-bottom: 1.5px solid #111;
      }

      .arrival-blank-row td {
        height: 8px;
      }

      .arrival-summary-totals {
        width: 58%;
        margin: 0 auto;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 10px;
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

        {!summary && (
          <table className="arrival-report-table arrival-grand-total">
            <colgroup>
              <col style={{ width: "3%" }} />
              <col style={{ width: "10%" }} />
              <col style={{ width: "10%" }} />
              <col style={{ width: "11%" }} />
              <col style={{ width: "5%" }} />
              <col style={{ width: "21%" }} />
              <col style={{ width: "6%" }} />
              <col style={{ width: "7%" }} />
              <col style={{ width: "12%" }} />
              <col style={{ width: "15%" }} />
            </colgroup>
            <tbody>
              <tr className="arrival-grand-total-row">
                <td colSpan={6}>Grand Total</td>
                <td colSpan={2}>{totalMeters.toLocaleString()}</td>
                <td colSpan={2}>{vehicles} VEHICLES</td>
              </tr>
            </tbody>
          </table>
        )}

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
        <section className="arrival-home">
          <button
            className="arrival-module-card arrival-module-card--stock"
            onClick={() => setSection("stock")}
            type="button"
          >
            <span className="arrival-card-icon">▤</span>
            <span className="arrival-card-title">ARRIVAL STOCK</span>
            <span className="arrival-card-description">
              Upload stock workbook and view available stock cards
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
            <h3>
              {section === "stock"
                ? "Upload Arrival Excel"
                : "Required Excel files"}
            </h3>

            <div className="arrival-two-files">
              <label className="arrival-file-picker">
                <span className="arrival-file-picker__label">
                  Current Arrival Excel / Stock
                </span>

                <span className="arrival-file-picker__control">
                  <span className="arrival-file-picker__button">
                    Choose Excel file
                  </span>

                  <span className="arrival-file-picker__name">
                    {stockFile?.name || stockBook?.name || "No file selected"}
                  </span>

                  <input
                    type="file"
                    accept=".xlsx,.xls,.xlsm"
                    onChange={(e) =>
                      loadStock(e.target.files?.[0] || null)
                    }
                  />
                </span>
              </label>

              {section === "unloading" && (
                <label className="arrival-file-picker">
                  <span className="arrival-file-picker__label">
                    Previous day's Unloading (optional)
                  </span>

                  <span className="arrival-file-picker__control">
                    <span className="arrival-file-picker__button">
                      Choose Excel file
                    </span>

                    <span className="arrival-file-picker__name">
                      {previousFile?.name ||
                        previousBook?.name ||
                        "No file selected"}
                    </span>

                    <input
                      type="file"
                      accept=".xlsx,.xls,.xlsm"
                      onChange={(e) =>
                        loadPrevious(e.target.files?.[0] || null)
                      }
                    />
                  </span>
                </label>
              )}
            </div>

            <label className="arrival-date-field">
              <span>Report date</span>
              <input
                type="date"
                value={reportDate}
                onChange={(e) => setReportDate(e.target.value)}
              />
            </label>

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
              <div className="arrival-status">
                Reading workbook…
              </div>
            )}

            {error && (
              <div className="arrival-error" role="alert">
                {error}
              </div>
            )}

            {stockBook && (
              <div className="arrival-status">
                Stock loaded: {stockRows.length} rows from STOCK sheet.
              </div>
            )}
          </div>

          {section === "stock" && (
            <div className="arrival-panel">
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
                    Clear
                  </button>
                </div>
              </div>

              <div className="arrival-card-grid">
                {filteredStockRows.map((r) => (
                  <article
                    className="arrival-data-card"
                    key={r.stockId}
                  >
                    <div className="arrival-data-card__top">
                      <b>
                        {r.type === "OWN" ? "OWN" : "COMMERCIAL"}
                      </b>
                      <strong>{r.meter} Mtr</strong>
                    </div>

                    {FIELD_LABELS.filter(
                      ([k]) => k !== "customer"
                    ).map(([k, label]) => (
                      <div
                        className="arrival-data-field"
                        key={k}
                      >
                        <span>{label}</span>
                        <b>{r[k] || "—"}</b>
                      </div>
                    ))}
                  </article>
                ))}
              </div>

              {filteredStockRows.length === 0 && (
                <div className="arrival-status">
                  Koi matching stock record nahi mila.
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
                    Rows already present in current/previous unloading
                    are excluded using their stock details and vehicle.
                  </p>

                  <div className="arrival-card-grid">
                    {availableRows.map((r) => (
                      <label
                        className="arrival-data-card arrival-select-card"
                        key={r.stockId}
                      >
                        <div className="arrival-data-card__top">
                          <b>
                            {r.type === "OWN" ? "OWN" : "COMMERCIAL"}
                          </b>

                          <input
                            type="checkbox"
                            checked={selected.includes(r.stockId)}
                            onChange={(e) =>
                              setSelected((old) =>
                                e.target.checked
                                  ? [...old, r.stockId]
                                  : old.filter(
                                      (x) => x !== r.stockId
                                    )
                              )
                            }
                          />
                        </div>

                        {FIELD_LABELS.filter(
                          ([k]) => k !== "customer"
                        ).map(([k, label]) => (
                          <div
                            className="arrival-data-field"
                            key={k}
                          >
                            <span>{label}</span>
                            <b>{r[k] || "—"}</b>
                          </div>
                        ))}
                      </label>
                    ))}
                  </div>

                  <button
                    type="button"
                    className="arrival-action arrival-action--primary"
                    disabled={!selected.length}
                    onClick={addSelected}
                  >
                    Add selected ({selected.length}) to Unloading
                  </button>
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

              {/* OWN UNLOADING CARDS */}
              <div className="arrival-panel">
                <h3>
                  Unloading Cards — Own ({ownRows.length})
                </h3>

                <div className="arrival-card-grid">
                  {ownRows.map((r) => (
                    <article
                      className="arrival-data-card"
                      key={r.id}
                    >
                      <div className="arrival-data-card__top">
                        <b>OWN · {r.supplier}</b>

                        <button
                          type="button"
                          onClick={() => removeRow(r.id)}
                        >
                          Remove
                        </button>
                      </div>

                      {FIELD_LABELS.map(([k, label]) => (
                        <label
                          className="arrival-data-field arrival-edit-field"
                          key={k}
                        >
                          <span>{label}</span>

                          <input
                            value={r[k] || ""}
                            onChange={(e) =>
                              updateRow(r.id, k, e.target.value)
                            }
                          />
                        </label>
                      ))}
                    </article>
                  ))}
                </div>
              </div>

              {/* COMMERCIAL UNLOADING CARDS */}
              <div className="arrival-panel">
                <h3>
                  Unloading Cards — Commercial ({cmtRows.length})
                </h3>

                <div className="arrival-card-grid">
                  {cmtRows.map((r) => (
                    <article
                      className="arrival-data-card"
                      key={r.id}
                    >
                      <div className="arrival-data-card__top">
                        <b>COMMERCIAL · {r.supplier}</b>

                        <button
                          type="button"
                          onClick={() => removeRow(r.id)}
                        >
                          Remove
                        </button>
                      </div>

                      {FIELD_LABELS.map(([k, label]) => (
                        <label
                          className="arrival-data-field arrival-edit-field"
                          key={k}
                        >
                          <span>{label}</span>

                          <input
                            value={r[k] || ""}
                            onChange={(e) =>
                              updateRow(r.id, k, e.target.value)
                            }
                          />
                        </label>
                      ))}
                    </article>
                  ))}
                </div>
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
