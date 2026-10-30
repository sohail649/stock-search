import { useEffect, useMemo, useState, useRef } from "react";
import * as XLSX from "xlsx";
import {
  collection,
  doc,
  getDoc,
  getDocs,
  setDoc,
  deleteDoc,
} from "firebase/firestore";
import { db } from "./firebase.js";
import "./App.css";

const UPLOAD_PASSWORD = "Stock@2025";

const bytesToBase64 = (bytes) => {
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i++)
    binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
};
const base64ToBytes = (base64) => {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

const SHEETS = [
  "Daily Stock Report (Own)",
  "Daily Stock Report (Commercial)",
  "Free For Sale (Own)",
  "Free Of Cost (Own)",
];

const HIDDEN_COMMERCIAL_SHEETS = ["Stock Grey Finished 205"];

const COMMERCIAL_SHEETS = [
  "Daily Stock Report (Commercial)",
  "Stock Grey Finished 205",
];

const clean = (value) =>
  String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
const normalize = (value) => clean(value).toLowerCase();
const normalizeDate = (value) => {
  const text = clean(value);
  if (!text || text === "—") return "";
  const match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (match) {
    let [, day, month, year] = match;
    if (year.length === 2) year = `20${year}`;
    return `${day.padStart(2, "0")}-${month.padStart(2, "0")}-${year}`;
  }
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    const day = String(parsed.getDate()).padStart(2, "0");
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const year = parsed.getFullYear();
    return `${day}-${month}-${year}`;
  }
  return text;
};

const getWorkbookStockDate = (workbook) => {
  for (const sheetName of SHEETS) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const cell = sheet.A3;
    if (
      !cell ||
      cell.v === undefined ||
      cell.v === null ||
      clean(cell.v) === ""
    )
      continue;
    let value = cell.v;
    if (cell.t === "n" && typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) {
        value = `${String(parsed.d).padStart(2, "0")}-${String(parsed.m).padStart(2, "0")}-${parsed.y}`;
        return value;
      }
    }
    const formatted = cell.w ? clean(cell.w) : clean(value);
    const normalized = normalizeDate(formatted);
    if (normalized) return normalized;
  }
  return "";
};

const getSheetTotals = (workbook, sheetName) => {
  if (!workbook || !workbook.Sheets[sheetName])
    return { items: 0, rolls: 0, aGrade: 0, bGrade: 0 };
  const sheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  const headerIndex = matrix.findIndex((row) => {
    const values = row.map(normalize);
    return (
      values.some((value) => /quality/.test(value)) &&
      values.some((value) => /width|witdh/.test(value))
    );
  });
  if (headerIndex === -1) return { items: 0, rolls: 0, aGrade: 0, bGrade: 0 };
  const headers = matrix[headerIndex].map(clean);
  const findColumn = (patterns) =>
    headers.findIndex((header) =>
      patterns.some((pattern) => pattern.test(header)),
    );
  const columns = {
    quality: findColumn([/^quality$/i, /^quality\s*\/\s*construction$/i]),
    width: findColumn([
      /^width$/i,
      /^witdh$/i,
      /^g\.?\s*width$/i,
      /^g\.?\s*witdh$/i,
      /witdh/i,
      /width/i,
    ]),
    weaveBlendSelvedge: findColumn([/^weave\s*,?\s*blend\s*,?\s*selvedge$/i]),
  };
  let closingCol = -1;
  for (let r = 0; r < Math.min(matrix.length, 15); r++) {
    const row = matrix[r];
    const idx = row.findIndex((cell) => /closing\s*balance/i.test(clean(cell)));
    if (idx !== -1) {
      closingCol = idx;
      break;
    }
  }
  let stockColumns = null;
  if (closingCol !== -1) {
    stockColumns = {
      rolls: closingCol,
      aGrade: closingCol + 1,
      bGrade: closingCol + 2,
    };
  }
  let totalRolls = 0,
    totalAGrade = 0,
    totalBGrade = 0,
    itemCount = 0;
  matrix.slice(headerIndex + 1).forEach((row) => {
    if (
      row.some(
        (c) =>
          /^\s*total\s*$/i.test(clean(c)) || /grand\s*total/i.test(clean(c)),
      )
    )
      return;
    const quality = clean(row[columns.quality]);
    const width = clean(row[columns.width]);
    const weave = clean(row[columns.weaveBlendSelvedge]);
    if (!quality && !width && !weave) return;
    if (/total/i.test(quality)) return;
    itemCount++;
    if (stockColumns) {
      const r = Number(clean(row[stockColumns.rolls]).replace(/,/g, ""));
      const a = Number(clean(row[stockColumns.aGrade]).replace(/,/g, ""));
      const b = Number(clean(row[stockColumns.bGrade]).replace(/,/g, ""));
      if (Number.isFinite(r)) totalRolls += r;
      if (Number.isFinite(a)) totalAGrade += a;
      if (Number.isFinite(b)) totalBGrade += b;
    }
  });
  return {
    items: itemCount,
    rolls: totalRolls,
    aGrade: totalAGrade,
    bGrade: totalBGrade,
  };
};

const getExactIssuanceTotals = (workbook) => {
  if (!workbook) return null;
  const result = {};
  SHEETS.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
    });
    let totalIssuanceCol = -1,
      totalReceivingCol = -1;
    for (let r = 0; r < Math.min(matrix.length, 15); r++) {
      const row = matrix[r];
      for (let c = 0; c < row.length; c++) {
        const cell = clean(row[c]);
        if (/total\s*issuance\s*to\s*date/i.test(cell)) totalIssuanceCol = c;
        if (/total\s*receiving\s*to\s*date/i.test(cell)) totalReceivingCol = c;
      }
    }
    const getTotalFromRow = (col, possibleRows) => {
      for (const r of possibleRows) {
        if (r >= matrix.length) continue;
        const row = matrix[r];
        if (col < row.length) {
          const val = clean(row[col]).replace(/,/g, "");
          const num = Number(val);
          if (Number.isFinite(num) && num !== 0) return num;
        }
      }
      return 0;
    };
    if (totalIssuanceCol !== -1) {
      const rolls = getTotalFromRow(totalIssuanceCol, [2, 4, 3, 5]);
      const aGrade = getTotalFromRow(totalIssuanceCol + 1, [2, 4, 3, 5]);
      const bGrade = getTotalFromRow(totalIssuanceCol + 2, [2, 4, 3, 5]);
      result[sheetName] = result[sheetName] || {};
      result[sheetName].issuance = {
        rolls,
        aGrade,
        bGrade,
        total: aGrade + bGrade,
      };
    }
    if (totalReceivingCol !== -1) {
      const rolls = getTotalFromRow(totalReceivingCol, [2, 4, 3, 5]);
      const aGrade = getTotalFromRow(totalReceivingCol + 1, [2, 4, 3, 5]);
      const bGrade = getTotalFromRow(totalReceivingCol + 2, [2, 4, 3, 5]);
      result[sheetName] = result[sheetName] || {};
      result[sheetName].receiving = {
        rolls,
        aGrade,
        bGrade,
        total: aGrade + bGrade,
      };
    }
  });
  return result;
};

const getTransactions = (workbook, type) => {
  if (!workbook) return [];
  const allData = [];
  SHEETS.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
    });
    const headerIndex = matrix.findIndex((row) => {
      const values = row.map(normalize);
      return (
        values.some((value) => /quality/.test(value)) &&
        values.some((value) => /width|witdh/.test(value))
      );
    });
    if (headerIndex === -1) return;
    const headers = matrix[headerIndex].map(clean);
    const findColumn = (patterns) =>
      headers.findIndex((header) =>
        patterns.some((pattern) => pattern.test(header)),
      );
    let mainHeadingRow = -1;
    let mainHeadingCol = -1;
    const searchPattern =
      type === "received"
        ? /total\s*receiving\s*to\s*date/i
        : /total\s*issuance\s*to\s*date/i;
    const fallbackPattern = type === "received" ? /receiving/i : /issuance/i;
    for (let r = 0; r < Math.min(matrix.length, 15); r++) {
      const row = matrix[r];
      for (let c = 0; c < row.length; c++) {
        if (searchPattern.test(clean(row[c]))) {
          mainHeadingRow = r;
          mainHeadingCol = c;
          break;
        }
      }
      if (mainHeadingRow !== -1) break;
    }
    if (mainHeadingRow === -1) {
      for (let r = 0; r < Math.min(matrix.length, 15); r++) {
        const row = matrix[r];
        for (let c = 0; c < row.length; c++) {
          if (fallbackPattern.test(clean(row[c]))) {
            if (/total/i.test(clean(row[c]))) {
              mainHeadingRow = r;
              mainHeadingCol = c;
              break;
            }
          }
        }
        if (mainHeadingRow !== -1) break;
      }
    }
    if (mainHeadingRow === -1 || mainHeadingCol === -1) return;
    let subHeaderRow = -1;
    let rbCol = -1;
    for (
      let r = mainHeadingRow + 1;
      r <= Math.min(mainHeadingRow + 3, matrix.length - 1);
      r++
    ) {
      const row = matrix[r];
      for (
        let c = mainHeadingCol;
        c < Math.min(mainHeadingCol + 5, row.length);
        c++
      ) {
        const cell = clean(row[c]).toLowerCase();
        if (
          /^r\/b$/i.test(cell) ||
          /^r\s*\/\s*b$/i.test(cell) ||
          /r\/b/i.test(cell)
        ) {
          subHeaderRow = r;
          rbCol = c;
          break;
        }
      }
      if (rbCol !== -1) break;
    }
    if (rbCol === -1) {
      rbCol = mainHeadingCol;
    }
    let aGradeCol = rbCol + 1;
    let bGradeCol = rbCol + 2;
    if (subHeaderRow !== -1) {
      const subRow = matrix[subHeaderRow];
      for (let c = rbCol; c < Math.min(rbCol + 4, subRow.length); c++) {
        const cell = clean(subRow[c]).toLowerCase();
        if (/a\.?grade/i.test(cell) && !/b\.?grade/i.test(cell)) {
          if (c === rbCol + 1) aGradeCol = c;
        }
        if (/b\.?grade/i.test(cell)) {
          if (c === rbCol + 2 || c === rbCol + 1) bGradeCol = c;
        }
      }
    }
    const transColumns = { rolls: rbCol, aGrade: aGradeCol, bGrade: bGradeCol };
    const cols = {
      quality: findColumn([/^quality$/i, /^quality\s*\/\s*construction$/i]),
      width: findColumn([
        /^width$/i,
        /^witdh$/i,
        /^g\.?\s*width$/i,
        /^g\.?\s*witdh$/i,
        /witdh/i,
        /width/i,
      ]),
      date: findColumn([
        /^last\s*received$/i,
        /^date$/i,
        /received\s*date/i,
        /issue\s*date/i,
        /received/i,
      ]),
      itemCode: findColumn([/^item\s*code$/i, /^item\s*#$/i]),
      serial: findColumn([/^sr\.?[-\s]*$/i, /^serial/i]),
      weave: findColumn([/^weave\s*,?\s*blend\s*,?\s*selvedge$/i]),
      supplier: findColumn([/^greige\s*supplier$/i]),
      customer: findColumn([/^processing\s*customer$/i]),
      type: findColumn([/^type$/i]),
      igp: findColumn([/\bigp\b/i, /\bgrn\b/i]),
      po: findColumn([/\bpo\b/i, /contract/i]),
      remarks: findColumn([/^remarks?$/i]),
    };
    matrix.slice(headerIndex + 1).forEach((row) => {
      if (
        row.some(
          (c) =>
            /^\s*total\s*$/i.test(clean(c)) || /grand\s*total/i.test(clean(c)),
        )
      )
        return;
      const quality = clean(row[cols.quality]);
      if (
        !quality ||
        /total/i.test(quality) ||
        /closing\s*balance/i.test(quality) ||
        /receiving/i.test(quality) ||
        /issuance/i.test(quality) ||
        /issuing/i.test(quality) ||
        /^r\/b$/i.test(quality) ||
        /a\.?grade/i.test(quality) ||
        /b\.?grade/i.test(quality)
      )
        return;
      const r = Number(clean(row[transColumns.rolls]).replace(/,/g, "")) || 0;
      const a = Number(clean(row[transColumns.aGrade]).replace(/,/g, "")) || 0;
      const b = Number(clean(row[transColumns.bGrade]).replace(/,/g, "")) || 0;
      if (r === 0 && a === 0 && b === 0) return;
      allData.push({
        sheet: sheetName,
        date: normalizeDate(clean(row[cols.date])),
        quality: clean(row[cols.quality]) || "—",
        width: clean(row[cols.width]) || "—",
        itemCode: clean(row[cols.itemCode]) || "—",
        serial: clean(row[cols.serial]) || "—",
        weave: clean(row[cols.weave]) || "—",
        supplier: clean(row[cols.supplier]) || "—",
        customer: clean(row[cols.customer]) || "—",
        type: clean(row[cols.type]) || "—",
        igp: clean(row[cols.igp]) || "—",
        po: clean(row[cols.po]) || "—",
        remarks: clean(row[cols.remarks]) || "—",
        rolls: r,
        aGrade: a,
        bGrade: b,
        totalMeters: a + b,
      });
    });
  });
  return allData;
};

function App() {
  const [workbook, setWorkbook] = useState(null);
  const [fileName, setFileName] = useState("");
  const [stockUpdatedTill, setStockUpdatedTill] = useState("");
  const [selectedSheet, setSelectedSheet] = useState(SHEETS[0]);
  const [qualityFilter, setQualityFilter] = useState("");
  const [widthFilter, setWidthFilter] = useState("");
  const [igpFilter, setIgpFilter] = useState("");
  const [poFilter, setPoFilter] = useState("");
  const [supplierFilter, setSupplierFilter] = useState("");
  const [customerFilter, setCustomerFilter] = useState("");
  const [serialFilter, setSerialFilter] = useState("");
  const [lastReceivedFilter, setLastReceivedFilter] = useState("");
  const [error, setError] = useState("");
  const [hideZeroStock, setHideZeroStock] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [currentView, setCurrentView] = useState("home");
  const [issuedDateFilter, setIssuedDateFilter] = useState("");
  const [receivedDateFilter, setReceivedDateFilter] = useState("");
  // === ADDED: Quality & Width filters for Issued & Received ===
  const [issuedQualityFilter, setIssuedQualityFilter] = useState("");
  const [issuedWidthFilter, setIssuedWidthFilter] = useState("");
  const [receivedQualityFilter, setReceivedQualityFilter] = useState("");
  const [receivedWidthFilter, setReceivedWidthFilter] = useState("");
  const uploadRef = useRef(null);

  useEffect(() => {
    let cancelled = false;
    const loadSavedWorkbook = async () => {
      try {
        const currentRef = doc(db, "stock", "current");
        const currentSnap = await getDoc(currentRef);
        if (!currentSnap.exists()) {
          setIsLoading(false);
          return;
        }
        const saved = currentSnap.data();
        const chunksSnap = await getDocs(
          collection(db, "stockBackups", saved.uploadId, "chunks"),
        );
        const chunks = chunksSnap.docs
          .map((item) => item.data())
          .sort((a, b) => a.index - b.index);
        if (
          cancelled ||
          chunks.length !== saved.chunkCount ||
          chunks.length === 0
        ) {
          setIsLoading(false);
          return;
        }
        const base64 = chunks.map((item) => item.data).join("");
        const bytes = base64ToBytes(base64);
        const savedWorkbook = XLSX.read(bytes, {
          type: "array",
          cellDates: true,
        });
        if (cancelled) return;
        setWorkbook(savedWorkbook);
        setFileName(saved.fileName || "Saved Excel file");
        setStockUpdatedTill(saved.stockUpdatedTill || "");
      } catch (err) {
        console.error("Load failed:", err);
      } finally {
        if (!cancelled) setIsLoading(false);
      }
    };
    loadSavedWorkbook();
    return () => {
      cancelled = true;
    };
  }, []);

  const allTotals = useMemo(() => {
    if (!workbook) return null;
    const perSheet = SHEETS.map((name) => ({
      name,
      ...getSheetTotals(workbook, name),
    }));
    const hiddenTotals = HIDDEN_COMMERCIAL_SHEETS.map((name) => ({
      name,
      ...getSheetTotals(workbook, name),
      isHidden: true,
    }));

    const ownSheets = perSheet.filter(
      (s) => !COMMERCIAL_SHEETS.includes(s.name),
    );
    const commercialVisible = perSheet.filter((s) =>
      COMMERCIAL_SHEETS.includes(s.name),
    );
    const commercialHidden = hiddenTotals.filter(
      (s) => COMMERCIAL_SHEETS.includes(s.name) && s.items > 0,
    );

    const ownTotal = ownSheets.reduce(
      (acc, cur) => ({
        items: acc.items + cur.items,
        rolls: acc.rolls + cur.rolls,
        aGrade: acc.aGrade + cur.aGrade,
        bGrade: acc.bGrade + cur.bGrade,
      }),
      { items: 0, rolls: 0, aGrade: 0, bGrade: 0 },
    );

    const allCommercialForTotal = [...commercialVisible, ...commercialHidden];
    const commercialTotal = allCommercialForTotal.reduce(
      (acc, cur) => ({
        items: acc.items + cur.items,
        rolls: acc.rolls + cur.rolls,
        aGrade: acc.aGrade + cur.aGrade,
        bGrade: acc.bGrade + cur.bGrade,
      }),
      { items: 0, rolls: 0, aGrade: 0, bGrade: 0 },
    );

    const grand = [...perSheet, ...commercialHidden].reduce(
      (acc, cur) => ({
        items: acc.items + cur.items,
        rolls: acc.rolls + cur.rolls,
        aGrade: acc.aGrade + cur.aGrade,
        bGrade: acc.bGrade + cur.bGrade,
      }),
      { items: 0, rolls: 0, aGrade: 0, bGrade: 0 },
    );

    return {
      perSheet,
      ownTotal,
      commercialTotal,
      commercialSheets: allCommercialForTotal,
      ownSheets,
      grand,
      hiddenTotals: commercialHidden,
    };
  }, [workbook]);

  const exactTotals = useMemo(
    () => getExactIssuanceTotals(workbook),
    [workbook],
  );

  const issuedData = useMemo(
    () => getTransactions(workbook, "issued"),
    [workbook],
  );
  const receivedData = useMemo(
    () => getTransactions(workbook, "received"),
    [workbook],
  );

  const filteredIssued = useMemo(() => {
    return issuedData.filter((item) => {
      const dateMatch =
        !issuedDateFilter ||
        normalizeDate(item.date) === normalizeDate(issuedDateFilter);
      const qualityMatch =
        !issuedQualityFilter ||
        normalize(item.quality).includes(normalize(issuedQualityFilter));
      const widthMatch =
        !issuedWidthFilter ||
        normalize(item.width).includes(normalize(issuedWidthFilter));
      return dateMatch && qualityMatch && widthMatch;
    });
  }, [issuedData, issuedDateFilter, issuedQualityFilter, issuedWidthFilter]);

  const filteredReceived = useMemo(() => {
    return receivedData.filter((item) => {
      const dateMatch =
        !receivedDateFilter ||
        normalizeDate(item.date) === normalizeDate(receivedDateFilter);
      const qualityMatch =
        !receivedQualityFilter ||
        normalize(item.quality).includes(normalize(receivedQualityFilter));
      const widthMatch =
        !receivedWidthFilter ||
        normalize(item.width).includes(normalize(receivedWidthFilter));
      return dateMatch && qualityMatch && widthMatch;
    });
  }, [
    receivedData,
    receivedDateFilter,
    receivedQualityFilter,
    receivedWidthFilter,
  ]);

  const handleShowUpload = () => {
    const enteredPassword = prompt("Upload ke liye password likho:");
    if (enteredPassword === null) return;
    if (enteredPassword !== UPLOAD_PASSWORD) {
      setError("Ghalat password!");
      return;
    }
    setError("");
    setCurrentView("upload");
    setTimeout(
      () =>
        uploadRef.current?.scrollIntoView({
          behavior: "smooth",
          block: "start",
        }),
      100,
    );
  };

  const handleUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setError("");
    setIsSaving(true);
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const bytes = new Uint8Array(e.target.result);
        const newWorkbook = XLSX.read(bytes, {
          type: "array",
          cellDates: true,
        });
        const missingSheets = SHEETS.filter(
          (sheetName) => !newWorkbook.SheetNames.includes(sheetName),
        );
        if (missingSheets.length) {
          setError(
            `Excel file mein required sheet nahi mili: ${missingSheets.join(", ")}`,
          );
          setIsSaving(false);
          return;
        }
        const newStockDate = getWorkbookStockDate(newWorkbook);
        const base64Data = bytesToBase64(bytes);
        const uploadId = Date.now().toString();
        const CHUNK_SIZE = 800000;
        const totalChunks = Math.ceil(base64Data.length / CHUNK_SIZE);
        const currentRef = doc(db, "stock", "current");
        const currentSnap = await getDoc(currentRef);
        if (currentSnap.exists()) {
          const oldId = currentSnap.data().uploadId;
          try {
            const oldChunks = await getDocs(
              collection(db, "stockBackups", oldId, "chunks"),
            );
            for (const d of oldChunks.docs) await deleteDoc(d.ref);
          } catch {}
        }
        for (let i = 0; i < totalChunks; i++) {
          const chunkData = base64Data.slice(
            i * CHUNK_SIZE,
            (i + 1) * CHUNK_SIZE,
          );
          await setDoc(
            doc(db, "stockBackups", uploadId, "chunks", i.toString()),
            { index: i, data: chunkData },
          );
        }
        await setDoc(doc(db, "stock", "current"), {
          uploadId,
          fileName: file.name,
          stockUpdatedTill: newStockDate,
          chunkCount: totalChunks,
          createdAt: new Date().toISOString(),
        });
        setWorkbook(newWorkbook);
        setFileName(file.name);
        setStockUpdatedTill(newStockDate);
        setSelectedSheet(SHEETS[0]);
        setQualityFilter("");
        setWidthFilter("");
        setIgpFilter("");
        setPoFilter("");
        setSupplierFilter("");
        setCustomerFilter("");
        setSerialFilter("");
        setLastReceivedFilter("");
        setError("File kamyabi se upload ho gayi!");
        setCurrentView("home");
      } catch {
        setError("Excel file read nahi ho saki.");
      } finally {
        setIsSaving(false);
      }
    };
    reader.onerror = () => {
      setError("File read nahi ho saki.");
      setIsSaving(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const data = useMemo(() => {
    if (!workbook || !workbook.Sheets[selectedSheet])
      return { rows: [], columns: {}, stockColumns: null, headers: [] };
    const sheet = workbook.Sheets[selectedSheet];
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
    });
    const headerIndex = matrix.findIndex((row) => {
      const values = row.map(normalize);
      return (
        values.some((value) => /quality/.test(value)) &&
        values.some((value) => /width|witdh/.test(value))
      );
    });
    if (headerIndex === -1)
      return { rows: [], columns: {}, stockColumns: null, headers: [] };
    const headers = matrix[headerIndex].map(clean);
    const findColumn = (patterns) =>
      headers.findIndex((header) =>
        patterns.some((pattern) => pattern.test(header)),
      );
    const columns = {
      lastReceived: findColumn([/^last\s*received$/i]),
      quality: findColumn([/^quality$/i, /^quality\s*\/\s*construction$/i]),
      width: findColumn([
        /^width$/i,
        /^witdh$/i,
        /^g\.?\s*width$/i,
        /^g\.?\s*witdh$/i,
        /witdh/i,
        /width/i,
      ]),
      igp: findColumn([/\bigp\b/i, /\bgrn\b/i]),
      po: findColumn([/\bpo\b/i, /contract/i]),
      supplier: findColumn([/^greige\s*supplier$/i]),
      customer: findColumn([/^processing\s*customer$/i]),
      serial: findColumn([/^sr\.?[-\s]*$/i, /^serial(?:\s*number)?$/i]),
      itemCode: findColumn([/^item\s*code$/i, /^item\s*#$/i]),
      type: findColumn([/^type$/i]),
      weaveBlendSelvedge: findColumn([
        /^weave\s*,?\s*blend\s*,?\s*selvedge$/i,
        /^weave\s*blend\s*selvedge$/i,
      ]),
    };
    const closingRowIndex = matrix.findIndex((row) =>
      row.some((cell) => /closing\s*balance/i.test(clean(cell))),
    );
    let stockColumns = null;
    if (closingRowIndex !== -1) {
      const closingRow = matrix[closingRowIndex];
      const closingColumn = closingRow.findIndex((cell) =>
        /closing\s*balance/i.test(clean(cell)),
      );
      if (closingColumn !== -1)
        stockColumns = {
          rolls: closingColumn,
          aGrade: closingColumn + 1,
          bGrade: closingColumn + 2,
        };
    }
    const rows = matrix
      .slice(headerIndex + 1)
      .filter((row) => {
        const quality = clean(row[columns.quality]);
        const width = clean(row[columns.width]);
        const weaveBlendSelvedge = clean(row[columns.weaveBlendSelvedge]);
        return Boolean(quality || width || weaveBlendSelvedge);
      })
      .map((row, index) => ({ id: index, values: row }));
    return { rows, columns, stockColumns, headers };
  }, [workbook, selectedSheet]);

  const getCell = (row, patterns) => {
    const headerIndex = data.headers.findIndex((header) =>
      patterns.some((pattern) => pattern.test(clean(header))),
    );
    if (headerIndex === -1) return "—";
    const value = clean(row.values[headerIndex]);
    return value || "—";
  };
  const getColumnCell = (row, columnIndex) => {
    if (columnIndex === undefined || columnIndex === null || columnIndex < 0)
      return "—";
    const value = clean(row.values[columnIndex]);
    return value || "—";
  };
  const getStockCell = (row, columnIndex) => {
    if (columnIndex === undefined || columnIndex === null) return "—";
    const value = clean(row.values[columnIndex]);
    return value || "0";
  };

  const filteredRows = useMemo(() => {
    return data.rows.filter((row) => {
      const valueAt = (column) =>
        column >= 0 ? clean(row.values[column]) : "";
      const quality = normalize(valueAt(data.columns.quality)),
        width = normalize(valueAt(data.columns.width)),
        igp = normalize(valueAt(data.columns.igp)),
        po = normalize(valueAt(data.columns.po)),
        supplier = normalize(valueAt(data.columns.supplier)),
        customer = normalize(valueAt(data.columns.customer)),
        serial = normalize(valueAt(data.columns.serial));
      const lastReceived = normalizeDate(valueAt(data.columns.lastReceived)),
        requestedDate = normalizeDate(lastReceivedFilter);
      const parseStock = (columnIndex) => {
        if (
          !data.stockColumns ||
          columnIndex === undefined ||
          columnIndex === null
        )
          return 0;
        const rawValue = clean(row.values[columnIndex]).replace(/,/g, "");
        const numberValue = Number(rawValue);
        return rawValue === "" || !Number.isFinite(numberValue)
          ? 0
          : numberValue;
      };
      const rolls = parseStock(data.stockColumns?.rolls),
        meters = parseStock(data.stockColumns?.aGrade),
        hasStock = rolls !== 0 || meters !== 0;
      return (
        (!hideZeroStock || hasStock) &&
        (!qualityFilter || quality.includes(normalize(qualityFilter))) &&
        (!widthFilter || width.includes(normalize(widthFilter))) &&
        (!igpFilter || igp.includes(normalize(igpFilter))) &&
        (!poFilter || po.includes(normalize(poFilter))) &&
        (!supplierFilter || supplier.includes(normalize(supplierFilter))) &&
        (!customerFilter || customer.includes(normalize(customerFilter))) &&
        (!serialFilter || serial.includes(normalize(serialFilter))) &&
        (!lastReceivedFilter ||
          normalize(lastReceived) === normalize(requestedDate))
      );
    });
  }, [
    data,
    qualityFilter,
    widthFilter,
    igpFilter,
    poFilter,
    supplierFilter,
    customerFilter,
    serialFilter,
    lastReceivedFilter,
    hideZeroStock,
  ]);

  const qualitySuggestions = useMemo(() => {
    if (!data.rows.length) return [];
    const all = [
      ...new Set(
        data.rows
          .map((r) => clean(r.values[data.columns.quality]))
          .filter(Boolean)
      ),
    ];
    if (!qualityFilter) return all.slice(0, 8);
    return all
      .filter((q) => normalize(q).includes(normalize(qualityFilter)))
      .slice(0, 8);
  }, [data.rows, data.columns.quality, qualityFilter]);

  const qualityWiseTotal = useMemo(() => {
    if (!qualityFilter) return null;
    const matched = data.rows.filter((row) =>
      normalize(clean(row.values[data.columns.quality])).includes(
        normalize(qualityFilter)
      )
    );
    if (!matched.length) return null;
    let rolls = 0,
      aGrade = 0,
      bGrade = 0;
    matched.forEach((row) => {
      rolls +=
        Number(clean(row.values[data.stockColumns?.rolls]).replace(/,/g, "")) ||
        0;
      aGrade +=
        Number(
          clean(row.values[data.stockColumns?.aGrade]).replace(/,/g, "")
        ) || 0;
      bGrade +=
        Number(
          clean(row.values[data.stockColumns?.bGrade]).replace(/,/g, "")
        ) || 0;
    });
    return {
      rolls,
      aGrade,
      bGrade,
      total: aGrade + bGrade,
      count: matched.length,
    };
  }, [data.rows, data.columns, data.stockColumns, qualityFilter]);

  const filteredAvailableTotal = useMemo(() => {
    if (!filteredRows.length)
      return { total: 0, rolls: 0, aGrade: 0, bGrade: 0 };
    let rolls = 0,
      aGrade = 0,
      bGrade = 0;
    filteredRows.forEach((row) => {
      rolls +=
        Number(clean(row.values[data.stockColumns?.rolls]).replace(/,/g, "")) ||
        0;
      aGrade +=
        Number(
          clean(row.values[data.stockColumns?.aGrade]).replace(/,/g, "")
        ) || 0;
      bGrade +=
        Number(
          clean(row.values[data.stockColumns?.bGrade]).replace(/,/g, "")
        ) || 0;
    });
    return { rolls, aGrade, bGrade, total: aGrade + bGrade };
  }, [filteredRows, data.stockColumns]);

  const clearFilters = () => {
    setQualityFilter("");
    setWidthFilter("");
    setIgpFilter("");
    setPoFilter("");
    setSupplierFilter("");
    setCustomerFilter("");
    setSerialFilter("");
    setLastReceivedFilter("");
  };

  const NavButtons = () => (
    <nav className="nav-buttons" aria-label="Stock pages">
      <button
        className="nav-home"
        onClick={() => setCurrentView("home")}
        type="button"
      >
        ⌂ Home
      </button>
    </nav>
  );

  const FullNavButtons = () => (
    <nav className="nav-buttons" aria-label="Stock pages">
      <button
        className="nav-home"
        onClick={() => setCurrentView("home")}
        type="button"
      >
        ⌂ Home
      </button>
      <button
        className={`nav-available ${currentView === "dashboard" ? "is-active" : ""}`}
        onClick={() => setCurrentView("dashboard")}
        type="button"
      >
        Available Stock
      </button>
      <button
        className={`nav-issued ${currentView === "issued" ? "is-active" : ""}`}
        onClick={() => setCurrentView("issued")}
        type="button"
      >
        Issued Stock
      </button>
      <button
        className={`nav-received ${currentView === "received" ? "is-active" : ""}`}
        onClick={() => setCurrentView("received")}
        type="button"
      >
        Received Stock
      </button>
      <button
        className={`nav-summary ${currentView === "summary" ? "is-active" : ""}`}
        onClick={() => setCurrentView("summary")}
        type="button"
      >
        Stock Summary
      </button>
      <button className="nav-upload" onClick={handleShowUpload} type="button">
        ↑ Upload Stock
      </button>
    </nav>
  );

  if (isLoading)
    return (
      <main className="app">
        <div className="empty-state">
          <h3>Loading saved stock...</h3>
          <p>stock data loading </p>
        </div>
      </main>
    );

  if (currentView === "home") {
    const grandMeters = allTotals
      ? allTotals.grand.aGrade + allTotals.grand.bGrade
      : 0;
    const grandRolls = allTotals ? allTotals.grand.rolls : 0;
    return (
      <main className="app home-app">
        <header className="topbar home-topbar">
          <div className="brand">
            <div className="brand-icon">▦</div>
            <div>
              <h1>Stock Manager</h1>
              <p>Inventory Management System</p>
            </div>
            <div className="status">
              {workbook ? "● Excel Loaded" : "● Ready to Start"}
            </div>
          </div>
          <div className="home-header-actions">
            <button
              className="nav-upload"
              onClick={handleShowUpload}
              type="button"
            >
              ↑ Upload Stock
            </button>
          </div>
        </header>

        <section className="home-hero">
          <div className="hero-copy">
            <span className="hero-eyebrow">SMART INVENTORY WORKSPACE</span>
            <h2>
              Manage Your Stock <span>Smartly.</span>
            </h2>
            <p>
              Commercial counting configurable - add any sheet to commercial
              total
            </p>
            {workbook && (
              <div className="hero-date">
                <span>Stock updated through</span>
                <strong>{stockUpdatedTill || "Date not available"}</strong>
              </div>
            )}
          </div>
        </section>

        <section className="workspace-section">
          <div className="workspace-heading">
            <span className="section-kicker">YOUR WORKSPACE</span>
            <h2>Choose Your Workspace</h2>
            <p>Select a page to get started.</p>
          </div>
          <div className="workspace-grid">
            <button
              className="workspace-card card-available"
              onClick={() => setCurrentView("dashboard")}
              type="button"
            >
              <span className="workspace-icon">⌕</span>
              <span className="workspace-arrow">↗</span>
              <span className="workspace-title">Available Stock</span>
              <span className="workspace-subtitle">Search inventory</span>
              <span className="card-shine"></span>
            </button>
            <button
              className="workspace-card card-issued"
              onClick={() => setCurrentView("issued")}
              type="button"
            >
              <span className="workspace-icon">▣−</span>
              <span className="workspace-arrow">↗</span>
              <span className="workspace-title">Issued Stock</span>
              <span className="workspace-subtitle">View issued records</span>
              <span className="card-shine"></span>
            </button>
            <button
              className="workspace-card card-received"
              onClick={() => setCurrentView("received")}
              type="button"
            >
              <span className="workspace-icon">▣＋</span>
              <span className="workspace-arrow">↗</span>
              <span className="workspace-title">Received Stock</span>
              <span className="workspace-subtitle">
                Check receiving records
              </span>
              <span className="card-shine"></span>
            </button>
            <button
              className="workspace-card card-summary"
              onClick={() => setCurrentView("summary")}
              type="button"
            >
              <span className="workspace-icon">▥↗</span>
              <span className="workspace-arrow">↗</span>
              <span className="workspace-title">Stock Summary</span>
              <span className="workspace-subtitle">
                Overview and reports - Commercial ={" "}
                {COMMERCIAL_SHEETS.join(" + ")}
              </span>
              <span className="card-shine"></span>
            </button>
          </div>
        </section>

        <section className="home-metrics" aria-label="Stock overview">
          <div className="metric-item">
            <span className="metric-symbol">▦</span>
            <div>
              <span>Total Stock</span>
              <strong>
                {workbook ? grandMeters.toLocaleString() : "—"}{" "}
                <small>MTR</small>
              </strong>
            </div>
          </div>
          <div className="metric-item">
            <span className="metric-symbol">◉</span>
            <div>
              <span>Total Rolls</span>
              <strong>{workbook ? grandRolls.toLocaleString() : "—"}</strong>
            </div>
          </div>
          <div className="metric-item">
            <span className="metric-symbol">▦</span>
            <div>
              <span>Stock Sheets</span>
              <strong>
                {workbook
                  ? SHEETS.filter((name) => workbook.Sheets[name]).length
                  : "—"}
              </strong>
            </div>
          </div>
        </section>

        <footer className="home-footer">
          <strong>Dada Manager • Smart Inventory. Better Management.</strong>
          <span>Commercial sheets: {COMMERCIAL_SHEETS.join(", ")}</span>
        </footer>
      </main>
    );
  }

  if (currentView === "summary") {
    return (
      <main className="app">
        <header className="topbar">
          <div className="brand">
            <div className="brand-icon">S</div>
            <div>
              <h1>Summary</h1>
              <p>Commercial = {COMMERCIAL_SHEETS.join(" + ")}</p>
            </div>
          </div>
          <NavButtons />
          <div className="status">
            {workbook ? "Excel Loaded" : "Awaiting Excel"}
          </div>
        </header>
        {workbook && (
          <div className="stock-updated-till">
            <strong>Stock Updated Till:</strong>{" "}
            <span>{stockUpdatedTill || "Date not found"}</span>
          </div>
        )}
        {!allTotals ? (
          <div className="empty-state">
            <h3>No data</h3>
            <p>Upload Excel first</p>
          </div>
        ) : (
          <section style={{ margin: "16px 0" }}>
            <div
              className="summary-grid"
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
                gap: "12px",
              }}
            >
              {allTotals.perSheet.map((sheet) => (
                <div
                  key={sheet.name}
                  style={{
                    background: COMMERCIAL_SHEETS.includes(sheet.name)
                      ? "#fffbeb"
                      : "white",
                    border: COMMERCIAL_SHEETS.includes(sheet.name)
                      ? "2px solid #fbbf24"
                      : "1px solid #e5e7eb",
                    borderRadius: "12px",
                    padding: "14px",
                  }}
                >
                  <div
                    style={{
                      fontSize: "11px",
                      fontWeight: 800,
                      color: COMMERCIAL_SHEETS.includes(sheet.name)
                        ? "#d97706"
                        : "#6b7280",
                      textTransform: "uppercase",
                      marginBottom: "6px",
                    }}
                  >
                    {sheet.name}{" "}
                    {COMMERCIAL_SHEETS.includes(sheet.name)
                      ? "⭐ COMMERCIAL"
                      : ""}
                  </div>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      fontSize: "13px",
                      marginBottom: "4px",
                    }}
                  >
                    <span>Items:</span>
                    <strong>{sheet.items}</strong>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      fontSize: "13px",
                      marginBottom: "4px",
                    }}
                  >
                    <span>Rolls:</span>
                    <strong>{sheet.rolls.toLocaleString()}</strong>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      fontSize: "13px",
                    }}
                  >
                    <span>A-Grade:</span>
                    <strong style={{ color: "#2563eb" }}>
                      {sheet.aGrade.toLocaleString()}
                    </strong>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      fontSize: "13px",
                      fontWeight: 700,
                    }}
                  >
                    <span>B-Grade:</span>
                    <strong style={{ color: "#dc2626" }}>
                      {sheet.bGrade.toLocaleString()}
                    </strong>
                  </div>
                  <div
                    style={{
                      marginTop: "8px",
                      borderTop: "1px solid #e5e7eb",
                      paddingTop: "6px",
                      display: "flex",
                      justifyContent: "space-between",
                      fontSize: "13px",
                      fontWeight: 800,
                    }}
                  >
                    <span>Total MTR:</span>
                    <strong>
                      {(sheet.aGrade + sheet.bGrade).toLocaleString()}
                    </strong>
                  </div>
                </div>
              ))}
            </div>
            <div
              className="summary-total-grid"
              style={{
                marginTop: "16px",
                background: "#111827",
                color: "white",
                borderRadius: "12px",
                padding: "20px",
                display: "grid",
                gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
                gap: "12px",
                textAlign: "center",
              }}
            >
              <div style={{ borderRight: "1px solid #374151" }}>
                <div style={{ fontSize: "12px", opacity: 0.7 }}>
                  TOTAL OWN METERS
                </div>
                <div style={{ fontSize: "11px", opacity: 0.6 }}>
                  {allTotals.ownSheets.map((s) => s.name).join(" + ") ||
                    "No Own"}
                </div>
                <div
                  style={{
                    fontSize: "24px",
                    fontWeight: 800,
                    color: "#60a5fa",
                    marginTop: "6px",
                  }}
                >
                  {(
                    allTotals.ownTotal.aGrade + allTotals.ownTotal.bGrade
                  ).toLocaleString()}{" "}
                  MTR
                </div>
                <div style={{ fontSize: "11px", opacity: 0.7 }}>
                  Rolls: {allTotals.ownTotal.rolls.toLocaleString()} | A:{" "}
                  {allTotals.ownTotal.aGrade.toLocaleString()} | B:{" "}
                  {allTotals.ownTotal.bGrade.toLocaleString()}
                </div>
              </div>
              <div style={{ borderRight: "1px solid #374151" }}>
                <div style={{ fontSize: "12px", opacity: 0.7 }}>
                  TOTAL COMMERCIAL METERS
                </div>
                <div style={{ fontSize: "11px", opacity: 0.6 }}>
                  {COMMERCIAL_SHEETS.join(" + ")}
                </div>
                <div
                  style={{
                    fontSize: "24px",
                    fontWeight: 800,
                    color: "#fbbf24",
                    marginTop: "6px",
                  }}
                >
                  {(
                    allTotals.commercialTotal.aGrade +
                    allTotals.commercialTotal.bGrade
                  ).toLocaleString()}{" "}
                  MTR
                </div>
                <div style={{ fontSize: "11px", opacity: 0.7 }}>
                  Rolls: {allTotals.commercialTotal.rolls.toLocaleString()} | A:{" "}
                  {allTotals.commercialTotal.aGrade.toLocaleString()} | B:{" "}
                  {allTotals.commercialTotal.bGrade.toLocaleString()}
                </div>
              </div>
              <div>
                <div style={{ fontSize: "12px", opacity: 0.7 }}>
                  GRAND TOTAL
                </div>
                <div style={{ fontSize: "11px", opacity: 0.6 }}>
                  All 4 Sheets
                </div>
                <div
                  style={{
                    fontSize: "26px",
                    fontWeight: 900,
                    color: "#34d399",
                    marginTop: "6px",
                  }}
                >
                  {(
                    allTotals.grand.aGrade + allTotals.grand.bGrade
                  ).toLocaleString()}{" "}
                  MTR
                </div>
                <div style={{ fontSize: "11px", opacity: 0.7 }}>
                  Rolls: {allTotals.grand.rolls.toLocaleString()} | A:{" "}
                  {allTotals.grand.aGrade.toLocaleString()} | B:{" "}
                  {allTotals.grand.bGrade.toLocaleString()}
                </div>
              </div>
            </div>
          </section>
        )}
      </main>
    );
  }

  if (currentView === "upload") {
    return (
      <main className="app">
        <header className="topbar">
          <div className="brand">
            <div className="brand-icon">U</div>
            <div>
              <h1>Upload Stock</h1>
              <p>Upload new Excel file</p>
            </div>
          </div>
          <NavButtons />
        </header>
        {workbook && (
          <div className="stock-updated-till">
            <strong>Stock Updated Till:</strong>{" "}
            <span>{stockUpdatedTill || "Date not found"}</span> | Current File:{" "}
            {fileName}
          </div>
        )}
        <section
          className="panel"
          ref={uploadRef}
          style={{ marginTop: "16px" }}
        >
          {error && (
            <div className="error-message" style={{ marginBottom: "16px" }}>
              {error}
            </div>
          )}
          <div
            style={{
              background: "#f0f7ff",
              border: "2px dashed #2563eb",
              padding: "24px",
              borderRadius: "12px",
              textAlign: "center",
            }}
          >
            <h3 style={{ marginBottom: "8px" }}>
              Upload Complete Excel Stock File
            </h3>
            <p
              style={{
                fontSize: "13px",
                color: "#6b7280",
                marginBottom: "16px",
              }}
            >
              Ye file Firebase me save ho jayegi aur har refresh par auto-load
              hogi
            </p>
            <input
              id="excel-upload"
              type="file"
              accept=".xlsx,.xls"
              onChange={handleUpload}
              disabled={isSaving}
              style={{ margin: "0 auto", display: "block" }}
            />
            {isSaving && (
              <p
                style={{ marginTop: "12px", fontWeight: 700, color: "#2563eb" }}
              >
                Saving to Firebase... 5-20 sec lag sakta hai
              </p>
            )}
            {fileName && (
              <div
                className="file-info"
                style={{ marginTop: "16px", justifyContent: "center" }}
              >
                <span>Current uploaded file:</span>
                <strong>{fileName}</strong>
              </div>
            )}
          </div>
        </section>
      </main>
    );
  }

  if (currentView === "issued") {
    const ownList = filteredIssued.filter(
      (i) => !COMMERCIAL_SHEETS.includes(i.sheet),
    );
    const commList = filteredIssued.filter((i) =>
      COMMERCIAL_SHEETS.includes(i.sheet),
    );
    const calc = (list) =>
      list.reduce(
        (acc, cur) => ({
          rolls: acc.rolls + cur.rolls,
          meters: acc.meters + cur.totalMeters,
          aGrade: acc.aGrade + cur.aGrade,
          bGrade: acc.bGrade + cur.bGrade,
        }),
        { rolls: 0, meters: 0, aGrade: 0, bGrade: 0 },
      );
    const totalOwn = calc(ownList);
    const calcComm = calc(commList);
    const exactComm =
      exactTotals &&
      exactTotals["Daily Stock Report (Commercial)"] &&
      exactTotals["Daily Stock Report (Commercial)"].issuance
        ? exactTotals["Daily Stock Report (Commercial)"].issuance
        : null;
    let totalComm = calcComm;
    if (exactComm && COMMERCIAL_SHEETS.length === 1) {
      totalComm = {
        rolls: exactComm.rolls || calcComm.rolls,
        meters: exactComm.total || calcComm.meters,
        aGrade: exactComm.aGrade,
        bGrade: exactComm.bGrade,
      };
    }
    const totalIssued = {
      rolls: totalOwn.rolls + totalComm.rolls,
      meters: totalOwn.meters + totalComm.meters,
      aGrade: totalOwn.aGrade + totalComm.aGrade,
      bGrade: totalOwn.bGrade + totalComm.bGrade,
    };
    return (
      <main className="app">
        <header className="topbar">
          <div className="brand">
            <div className="brand-icon">I</div>
            <div>
              <h1>TOTAL ISSUANCE</h1>
            </div>
          </div>
          <NavButtons />
        </header>
        <section className="panel" style={{ marginTop: "16px" }}>
          <div
            style={{
              display: "flex",
              gap: "12px",
              alignItems: "end",
              flexWrap: "wrap",
            }}
          >
            <div className="field">
              <label>Date Filter (DD-MM-YYYY)</label>
              <input
                value={issuedDateFilter}
                onChange={(e) => setIssuedDateFilter(e.target.value)}
                placeholder="DD-MM-YYYY"
              />
            </div>
            <div className="field">
              <label>Quality / Construction</label>
              <input
                value={issuedQualityFilter}
                onChange={(e) => setIssuedQualityFilter(e.target.value)}
                placeholder="Enter quality"
              />
            </div>
            <div className="field">
              <label>Width</label>
              <input
                value={issuedWidthFilter}
                onChange={(e) => setIssuedWidthFilter(e.target.value)}
                placeholder="Enter width"
              />
            </div>
            <button
              onClick={() => {
                setIssuedDateFilter("");
                setIssuedQualityFilter("");
                setIssuedWidthFilter("");
              }}
              style={{
                padding: "10px 14px",
                borderRadius: "8px",
                border: "1px solid #cbd5e1",
                cursor: "pointer",
              }}
            >
              Clear Filters
            </button>
            <span style={{ marginLeft: "auto", fontWeight: 700 }}>
              {filteredIssued.length} records | Total:{" "}
              {totalIssued.meters.toLocaleString()} MTR
            </span>
          </div>
        </section>
        <div
          className="transaction-total-cards"
          style={{
            marginTop: "12px",
            background: "#111827",
            color: "white",
            borderRadius: "12px",
            padding: "18px",
            display: "grid",
            gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
            gap: "12px",
            textAlign: "center",
          }}
        >
          <div style={{ borderRight: "1px solid #374151" }}>
            <div style={{ fontSize: "11px", opacity: 0.7 }}>
              TOTAL OWN ISSUED
            </div>
            <div
              style={{ fontSize: "20px", fontWeight: 800, color: "#60a5fa" }}
            >
              {totalOwn.meters.toLocaleString()} MTR
            </div>
            <div style={{ fontSize: "10px", opacity: 0.6 }}>
              {allTotals?.ownSheets.map((s) => s.name).join(" + ")}
            </div>
            <div style={{ fontSize: "11px", opacity: 0.7, marginTop: "4px" }}>
              Rolls: {totalOwn.rolls.toLocaleString()} | A:{" "}
              {totalOwn.aGrade.toLocaleString()} | B:{" "}
              {totalOwn.bGrade.toLocaleString()}
            </div>
          </div>
          <div style={{ borderRight: "1px solid #374151" }}>
            <div style={{ fontSize: "11px", opacity: 0.7 }}>
              TOTAL COMMERCIAL ISSUED
            </div>
            <div
              style={{ fontSize: "20px", fontWeight: 800, color: "#fbbf24" }}
            >
              {totalComm.meters.toLocaleString()} MTR
            </div>
            <div style={{ fontSize: "10px", opacity: 0.6 }}>
              {COMMERCIAL_SHEETS.join(" + ")}
            </div>
            <div style={{ fontSize: "11px", opacity: 0.7, marginTop: "4px" }}>
              Rolls: {totalComm.rolls.toLocaleString()} | A:{" "}
              {totalComm.aGrade.toLocaleString()} | B:{" "}
              {totalComm.bGrade.toLocaleString()}
            </div>
          </div>
          <div>
            <div style={{ fontSize: "11px", opacity: 0.7 }}>
              GRAND TOTAL ISSUED
            </div>
            <div
              style={{ fontSize: "22px", fontWeight: 900, color: "#f87171" }}
            >
              {totalIssued.meters.toLocaleString()} MTR
            </div>
            <div style={{ fontSize: "10px", opacity: 0.6 }}>All Sheets</div>
            <div style={{ fontSize: "11px", opacity: 0.7, marginTop: "4px" }}>
              Rolls: {totalIssued.rolls.toLocaleString()} | Records:{" "}
              {filteredIssued.length}
            </div>
          </div>
        </div>
        <section className="results-section">
          <div className="stock-list">
            {filteredIssued.map((item, idx) => (
              <article
                className="stock-card"
                key={idx}
                style={{ borderLeft: "4px solid #dc2626" }}
              >
                <div className="stock-date">
                  <span>Date</span>
                  <strong>{item.date || "—"}</strong>
                </div>
                <div className="stock-card-title">
                  <div>
                    <span className="card-eyebrow">
                      {item.sheet} - {item.serial}
                    </span>
                    <h3>{item.quality}</h3>
                  </div>
                  <span className="width-tag">Width: {item.width}</span>
                </div>
                <div className="record-details">
                  <div className="detail-row item-code-box">
                    <span className="detail-label">Item Code</span>
                    <span className="detail-value">{item.itemCode}</span>
                  </div>
                  <div className="detail-row serial-box">
                    <span className="detail-label">Sr-</span>
                    <span className="detail-value">{item.serial}</span>
                  </div>
                  <div className="detail-row weave-box">
                    <span className="detail-label">Weave, Blend, Selvedge</span>
                    <span className="detail-value">{item.weave}</span>
                  </div>
                  <div className="detail-row supplier-box">
                    <span className="detail-label">Supplier / IGP-GRN</span>
                    <span className="detail-value">
                      {item.supplier} / {item.igp}
                    </span>
                  </div>
                  <div className="detail-row customer-box">
                    <span className="detail-label">Customer / Contract</span>
                    <span className="detail-value">
                      {item.customer} / {item.po}
                    </span>
                  </div>
                </div>
                <div className="stock-amounts">
                  <div className="amount-box rolls-box">
                    <span className="amount-label">Type / R/B</span>
                    <strong className="amount-value">
                      {item.type} / {item.rolls}
                    </strong>
                  </div>
                  <div
                    className="amount-box a-grade-box"
                    style={{ background: "#fef2f2", borderColor: "#fecaca" }}
                  >
                    <span className="amount-label">
                      ISSUANCE - A.Grade / B.Grade
                    </span>
                    <strong
                      className="amount-value"
                      style={{ color: "#dc2626" }}
                    >
                      {item.aGrade} / {item.bGrade}
                    </strong>
                    <span className="amount-unit">
                      MTR / MTR - Total: {item.totalMeters}
                    </span>
                  </div>
                </div>
                <div className="remarks-box">
                  <span className="remarks-label">Remarks</span>
                  <p>{item.remarks}</p>
                </div>
              </article>
            ))}
          </div>
        </section>
      </main>
    );
  }

  if (currentView === "received") {
    const ownList = filteredReceived.filter(
      (i) => !COMMERCIAL_SHEETS.includes(i.sheet),
    );
    const commList = filteredReceived.filter((i) =>
      COMMERCIAL_SHEETS.includes(i.sheet),
    );
    const calc = (list) =>
      list.reduce(
        (acc, cur) => ({
          rolls: acc.rolls + cur.rolls,
          meters: acc.meters + cur.totalMeters,
          aGrade: acc.aGrade + cur.aGrade,
          bGrade: acc.bGrade + cur.bGrade,
        }),
        { rolls: 0, meters: 0, aGrade: 0, bGrade: 0 },
      );
    const totalOwn = calc(ownList);
    const totalComm = calc(commList);
    const totalReceived = calc(filteredReceived);
    return (
      <main className="app">
        <header className="topbar">
          <div className="brand">
            <div className="brand-icon">R</div>
            <div>
              <h1>Received - TOTAL RECEIVING</h1>
              <p>Commercial: {COMMERCIAL_SHEETS.join(" + ")}</p>
            </div>
          </div>
          <NavButtons />
        </header>
        <section className="panel" style={{ marginTop: "16px" }}>
          <div
            style={{
              display: "flex",
              gap: "12px",
              alignItems: "end",
              flexWrap: "wrap",
            }}
          >
            <div className="field">
              <label>Date Filter (DD-MM-YYYY)</label>
              <input
                value={receivedDateFilter}
                onChange={(e) => setReceivedDateFilter(e.target.value)}
                placeholder="DD-MM-YYYY"
              />
            </div>
            <div className="field">
              <label>Quality / Construction</label>
              <input
                value={receivedQualityFilter}
                onChange={(e) => setReceivedQualityFilter(e.target.value)}
                placeholder="Enter quality"
              />
            </div>
            <div className="field">
              <label>Width</label>
              <input
                value={receivedWidthFilter}
                onChange={(e) => setReceivedWidthFilter(e.target.value)}
                placeholder="Enter width"
              />
            </div>
            <button
              onClick={() => {
                setReceivedDateFilter("");
                setReceivedQualityFilter("");
                setReceivedWidthFilter("");
              }}
              style={{
                padding: "10px 14px",
                borderRadius: "8px",
                border: "1px solid #cbd5e1",
                cursor: "pointer",
              }}
            >
              Clear Filters
            </button>
            <span style={{ marginLeft: "auto", fontWeight: 700 }}>
              {filteredReceived.length} records | Total:{" "}
              {totalReceived.meters.toLocaleString()} MTR
            </span>
          </div>
        </section>

        <div
          className="transaction-total-cards"
          style={{
            marginTop: "12px",
            background: "#111827",
            color: "white",
            borderRadius: "12px",
            padding: "18px",
            display: "grid",
            gridTemplateColumns: "repeat(3, 1fr)",
            gap: "12px",
            textAlign: "center",
          }}
        >
          <div style={{ borderRight: "1px solid #374151" }}>
            <div style={{ fontSize: "11px", opacity: 0.7 }}>
              TOTAL OWN RECEIVED
            </div>
            <div
              style={{ fontSize: "20px", fontWeight: 800, color: "#60a5fa" }}
            >
              {totalOwn.meters.toLocaleString()} MTR
            </div>
            <div style={{ fontSize: "10px", opacity: 0.6 }}>
              {allTotals?.ownSheets.map((s) => s.name).join(" + ")}
            </div>
            <div style={{ fontSize: "11px", opacity: 0.7, marginTop: "4px" }}>
              Rolls: {totalOwn.rolls.toLocaleString()} | A:{" "}
              {totalOwn.aGrade.toLocaleString()} | B:{" "}
              {totalOwn.bGrade.toLocaleString()}
            </div>
          </div>
          <div style={{ borderRight: "1px solid #374151" }}>
            <div style={{ fontSize: "11px", opacity: 0.7 }}>
              TOTAL COMMERCIAL RECEIVED
            </div>
            <div
              style={{ fontSize: "20px", fontWeight: 800, color: "#fbbf24" }}
            >
              {totalComm.meters.toLocaleString()} MTR
            </div>
            <div style={{ fontSize: "10px", opacity: 0.6 }}>
              {COMMERCIAL_SHEETS.join(" + ")}
            </div>
            <div style={{ fontSize: "11px", opacity: 0.7, marginTop: "4px" }}>
              Rolls: {totalComm.rolls.toLocaleString()} | A:{" "}
              {totalComm.aGrade.toLocaleString()} | B:{" "}
              {totalComm.bGrade.toLocaleString()}
            </div>
          </div>
          <div>
            <div style={{ fontSize: "11px", opacity: 0.7 }}>
              GRAND TOTAL RECEIVED
            </div>
            <div
              style={{ fontSize: "22px", fontWeight: 900, color: "#34d399" }}
            >
              {totalReceived.meters.toLocaleString()} MTR
            </div>
            <div style={{ fontSize: "10px", opacity: 0.6 }}>All Sheets</div>
            <div style={{ fontSize: "11px", opacity: 0.7, marginTop: "4px" }}>
              Rolls: {totalReceived.rolls.toLocaleString()} | Records:{" "}
              {filteredReceived.length}
            </div>
          </div>
        </div>
        <section className="results-section">
          <div className="stock-list issued-stock-list">
            {filteredReceived.map((item, idx) => (
              <article
                className="stock-card issued-stock-card"
                key={idx}
                style={{ borderLeft: "4px solid #059669" }}
              >
                <div className="stock-date">
                  <span>Date</span>
                  <strong>{item.date || "—"}</strong>
                </div>
                <div className="stock-card-title">
                  <div>
                    <span className="card-eyebrow">
                      {item.sheet} - {item.serial}
                    </span>
                    <h3>{item.quality}</h3>
                  </div>
                  <span className="width-tag">Width: {item.width}</span>
                </div>
                <div className="record-details">
                  <div className="detail-row item-code-box">
                    <span className="detail-label">Item Code</span>
                    <span className="detail-value">{item.itemCode}</span>
                  </div>
                  <div className="detail-row serial-box">
                    <span className="detail-label">Sr-</span>
                    <span className="detail-value">{item.serial}</span>
                  </div>
                  <div className="detail-row weave-box">
                    <span className="detail-label">Weave, Blend, Selvedge</span>
                    <span className="detail-value">{item.weave}</span>
                  </div>
                  <div className="detail-row supplier-box">
                    <span className="detail-label">Supplier / IGP-GRN</span>
                    <span className="detail-value">
                      {item.supplier} / {item.igp}
                    </span>
                  </div>
                  <div className="detail-row customer-box">
                    <span className="detail-label">Customer / Contract</span>
                    <span className="detail-value">
                      {item.customer} / {item.po}
                    </span>
                  </div>
                </div>
                <div className="stock-amounts">
                  <div className="amount-box rolls-box">
                    <span className="amount-label">Type / R/B</span>
                    <strong className="amount-value">
                      {item.type} / {item.rolls}
                    </strong>
                  </div>
                  <div
                    className="amount-box a-grade-box"
                    style={{ background: "#ecfdf5", borderColor: "#a7f3d0" }}
                  >
                    <span className="amount-label">
                      RECEIVING - A.Grade / B.Grade
                    </span>
                    <strong
                      className="amount-value"
                      style={{ color: "#059669" }}
                    >
                      {item.aGrade} / {item.bGrade}
                    </strong>
                    <span className="amount-unit">
                      MTR / MTR - Total: {item.totalMeters}
                    </span>
                  </div>
                </div>
                <div className="remarks-box">
                  <span className="remarks-label">Remarks</span>
                  <p>{item.remarks}</p>
                </div>
              </article>
            ))}
          </div>
        </section>
      </main>
    );
  }

  return (
    <main className="app">
      <header className="topbar">
        <div className="brand">
          <div className="brand-icon">S</div>
          <div>
            <h1>Available Stock</h1>
            <p>Search and manage your inventory</p>
          </div>
        </div>
        <NavButtons />
      </header>
      {workbook && (
        <div className="stock-updated-till">
          <strong>Stock Updated Till:</strong>{" "}
          <span>{stockUpdatedTill || "Date not found in Row 3"}</span>
        </div>
      )}
      {error && (
        <div className="error-message" style={{ marginTop: "12px" }}>
          {error}
        </div>
      )}
      <section className="panel" style={{ marginTop: "12px" }}>
        <div className="section-heading">
          <div>
            <h3>Stock Sheet</h3>
            <p>Select the stock report you want to search.</p>
          </div>
        </div>
        <label htmlFor="sheet-select">Select sheet</label>
        <select
          id="sheet-select"
          value={selectedSheet}
          onChange={(event) => setSelectedSheet(event.target.value)}
        >
          {SHEETS.map((sheetName) => (
            <option key={sheetName} value={sheetName}>
              {sheetName}
            </option>
          ))}
        </select>
        <div className="section-heading">
          <div>
            <h3>Search Filters</h3>
            <p>Enter one or more values to narrow down the results.</p>
          </div>
          <button className="clear-button" onClick={clearFilters} type="button">
            Clear filters
          </button>
        </div>
        <div className="filters">
          <div className="field">
            <label htmlFor="quality-filter">Quality / Construction</label>
            <input
              id="quality-filter"
              value={qualityFilter}
              onChange={(event) => setQualityFilter(event.target.value)}
              placeholder="Enter quality"
              list="quality-suggestions"
            />
            <datalist id="quality-suggestions">
              {qualitySuggestions.map((q, i) => (
                <option key={i} value={q} />
              ))}
            </datalist>
            {qualityFilter && qualityWiseTotal && (
              <div
                style={{
                  fontSize: "11px",
                  fontWeight: 700,
                  marginTop: "4px",
                  background: "#f0f7ff",
                  padding: "6px 8px",
                  borderRadius: "6px",
                  border: "1px solid #bfdbfe",
                  lineHeight: "1.4",
                }}
              >
                <span style={{ color: "#2563eb" }}>
                  {qualityWiseTotal.count} items
                </span>{" "}
                | Total:{" "}
                <span style={{ color: "#059669" }}>
                  {qualityWiseTotal.total.toLocaleString()} MTR
                </span>
              </div>
            )}
          </div>
          <div className="field">
            <label htmlFor="width-filter">Width</label>
            <input
              id="width-filter"
              value={widthFilter}
              onChange={(event) => setWidthFilter(event.target.value)}
              placeholder="Enter width"
            />
          </div>
          <div className="field">
            <label htmlFor="igp-filter">IGP / GRN</label>
            <input
              id="igp-filter"
              value={igpFilter}
              onChange={(event) => setIgpFilter(event.target.value)}
              placeholder="Search IGP / GRN"
            />
          </div>
          <div className="field">
            <label htmlFor="po-filter">PO / Contract</label>
            <input
              id="po-filter"
              value={poFilter}
              onChange={(event) => setPoFilter(event.target.value)}
              placeholder="Search PO / Contract"
            />
          </div>
          <div className="field">
            <label htmlFor="supplier-filter">Supplier</label>
            <input
              id="supplier-filter"
              value={supplierFilter}
              onChange={(event) => setSupplierFilter(event.target.value)}
              placeholder="Search supplier"
            />
          </div>
          <div className="field">
            <label htmlFor="customer-filter">Customer</label>
            <input
              id="customer-filter"
              value={customerFilter}
              onChange={(event) => setCustomerFilter(event.target.value)}
              placeholder="Search customer"
            />
          </div>
          <div className="field">
            <label htmlFor="serial-filter">Serial (Sr-)</label>
            <input
              id="serial-filter"
              value={serialFilter}
              onChange={(event) => setSerialFilter(event.target.value)}
              placeholder="Enter serial number"
            />
          </div>
          <div className="field">
            <label htmlFor="last-received-filter">
              Last Received (DD-MM-YYYY)
            </label>
            <input
              id="last-received-filter"
              value={lastReceivedFilter}
              onChange={(event) => setLastReceivedFilter(event.target.value)}
              placeholder="DD-MM-YYYY"
            />
          </div>
        </div>
        <div
          style={{
            marginTop: "12px",
            fontWeight: 700,
            fontSize: "13px",
            textAlign: "right",
            background: "#fff",
            padding: "10px 12px",
            borderRadius: "8px",
            border: "1px solid #e5e7eb",
          }}
        >
          {filteredRows.length} records | Total:{" "}
          {filteredAvailableTotal.total.toLocaleString()} MTR
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: "8px",
            marginTop: "16px",
            fontSize: "13px",
            fontWeight: 700,
          }}
        >
          <input
            id="hide-zero-stock"
            type="checkbox"
            checked={hideZeroStock}
            onChange={(event) => setHideZeroStock(event.target.checked)}
            style={{ width: "16px", height: "16px", cursor: "pointer" }}
          />
          <label htmlFor="hide-zero-stock" style={{ cursor: "pointer" }}>
            Hide Zero Stock
          </label>
        </div>
      </section>
      <section className="results-section">
        <div className="results-heading">
          <div>
            <h3>Stock Results</h3>
            <p>
              {workbook
                ? `Showing records from ${selectedSheet}`
                : "Upload an Excel file to view your stock."}
            </p>
          </div>
          <span className="count">{filteredRows.length} records</span>
        </div>
        {!workbook ? (
          <div className="empty-state">
            <div className="empty-icon">↥</div>
            <h3>Upload your Excel file</h3>
            <p>Go to Upload page</p>
          </div>
        ) : filteredRows.length === 0 ? (
          <div className="empty-state">
            <div className="empty-icon">⌕</div>
            <h3>No matching records</h3>
            <p>Try changing your filters.</p>
          </div>
        ) : (
          <div className="available-stock-list">
            {filteredRows.map((row) => {
              const quality = getColumnCell(row, data.columns.quality),
                width = getColumnCell(row, data.columns.width),
                igpNumber = getColumnCell(row, data.columns.igp),
                contractNumber = getColumnCell(row, data.columns.po),
                lastReceived = normalizeDate(
                  getColumnCell(row, data.columns.lastReceived),
                );
              return (
                <article className="available-stock-card" key={row.id}>
                  <div className="stock-date">
                    <span>Last Received</span>
                    <strong>{lastReceived || "—"}</strong>
                  </div>
                  <div className="stock-card-title">
                    <div>
                      <span className="card-eyebrow">STOCK ITEM</span>
                      <h3>
                        {quality === "—" ? "Quality not available" : quality}
                      </h3>
                    </div>
                    <span className="width-tag">Width: {width}</span>
                  </div>
                  <div className="record-details">
                    <div className="detail-row item-code-box">
                      <span className="detail-label">Item Code</span>
                      <span className="detail-value">
                        {getColumnCell(row, data.columns.itemCode)}
                      </span>
                    </div>
                    <div className="detail-row serial-box">
                      <span className="detail-label">Sr-</span>
                      <span className="detail-value">
                        {getColumnCell(row, data.columns.serial)}
                      </span>
                    </div>
                    <div className="detail-row weave-box">
                      <span className="detail-label">
                        Weave, Blend, Selvedge
                      </span>
                      <span className="detail-value">
                        {getColumnCell(row, data.columns.weaveBlendSelvedge)}
                      </span>
                    </div>
                    <div className="detail-row supplier-box">
                      <span className="detail-label">Supplier / IGP-GRN</span>
                      <span className="detail-value">
                        {getCell(row, [/^greige\s*supplier$/i])} / {igpNumber}
                      </span>
                    </div>
                    <div className="detail-row customer-box">
                      <span className="detail-label">Customer / Contract</span>
                      <span className="detail-value">
                        {getCell(row, [/^processing\s*customer$/i])} /{" "}
                        {contractNumber}
                      </span>
                    </div>
                  </div>
                  <div className="stock-amounts">
                    <div className="amount-box rolls-box">
                      <span className="amount-label">Type / Rolls-Bales</span>
                      <strong className="amount-value">
                        {getColumnCell(row, data.columns.type)} /{" "}
                        {getStockCell(row, data.stockColumns?.rolls)}
                      </strong>
                      <span className="amount-unit">Type / R-B</span>
                    </div>
                    <div className="amount-box a-grade-box">
                      <span className="amount-label">A-Grade / B-Grade</span>
                      <strong className="amount-value">
                        {getStockCell(row, data.stockColumns?.aGrade)} /{" "}
                        {getStockCell(row, data.stockColumns?.bGrade)}
                      </strong>
                      <span className="amount-unit">MTR / MTR</span>
                    </div>
                  </div>
                  <div className="remarks-box">
                    <span className="remarks-label">Remarks</span>
                    <p>{getCell(row, [/^remarks?$/i, /remarks?/i])}</p>
                  </div>
                </article>
              );
            })}
          </div>
        )}
      </section>
      <footer>Stock Search Dashboard</footer>
    </main>
  );
}
export default App;
