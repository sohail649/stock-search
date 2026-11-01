import { useEffect, useState } from "react";
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
import "./Reconciliation.css";

const CHUNK_SIZE = 750000;

const bytesToBase64 = (buffer) => {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
  }
  return btoa(binary);
};

const base64ToBytes = (base64) => {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

const saveReconciliationFile = async (file, type) => {
  const buffer = await file.arrayBuffer();
  const base64 = bytesToBase64(buffer);
  const uploadId = `recon_${type}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const chunks = [];

  for (let i = 0; i < base64.length; i += CHUNK_SIZE) {
    chunks.push(base64.slice(i, i + CHUNK_SIZE));
  }
  if (!chunks.length) chunks.push("");

  for (let i = 0; i < chunks.length; i++) {
    await setDoc(
      doc(db, "reconciliationBackups", uploadId, "chunks", String(i)),
      { index: i, data: chunks[i] }
    );
  }

  const metaRef = doc(db, "reconciliationFiles", type);
  const oldMeta = await getDoc(metaRef);
  const oldId = oldMeta.exists() ? oldMeta.data().uploadId : null;

  await setDoc(metaRef, {
    uploadId,
    fileName: file.name,
    chunkCount: chunks.length,
    updatedAt: Date.now(),
  });

  if (oldId && oldId !== uploadId) {
    try {
      const oldChunks = await getDocs(
        collection(db, "reconciliationBackups", oldId, "chunks")
      );
      for (const cd of oldChunks.docs) await deleteDoc(cd.ref);
    } catch {}
  }
};

const loadReconciliationFile = async (type) => {
  try {
    const metaRef = doc(db, "reconciliationFiles", type);
    const snap = await getDoc(metaRef);
    if (!snap.exists()) return null;

    const meta = snap.data();
    if (!meta.uploadId || !meta.chunkCount) return null;

    const chunksSnap = await getDocs(
      collection(db, "reconciliationBackups", meta.uploadId, "chunks")
    );
    if (chunksSnap.empty) return null;

    const sorted = chunksSnap.docs
      .map((d) => d.data())
      .sort((a, b) => a.index - b.index);

    const base64 = sorted.map((c) => c.data).join("");
    const bytes = base64ToBytes(base64);

    return new File([bytes], meta.fileName, {
      type: "application/octet-stream",
    });
  } catch {
    return null;
  }
};

const REPORT_CHUNK_SIZE = 750000;

const saveReconciliationReport = async (payload) => {
  const jsonStr = JSON.stringify(payload);
  const base64 = btoa(unescape(encodeURIComponent(jsonStr)));
  const uploadId = `report_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  const chunks = [];

  for (let i = 0; i < base64.length; i += REPORT_CHUNK_SIZE) {
    chunks.push(base64.slice(i, i + REPORT_CHUNK_SIZE));
  }
  if (!chunks.length) chunks.push("");

  for (let i = 0; i < chunks.length; i++) {
    await setDoc(
      doc(db, "reconciliationReportBackups", uploadId, "chunks", String(i)),
      { index: i, data: chunks[i] }
    );
  }

  const metaRef = doc(db, "reconciliationReports", "latest");
  const oldMeta = await getDoc(metaRef);
  const oldId = oldMeta.exists() ? oldMeta.data().uploadId : null;

  await setDoc(metaRef, {
    uploadId,
    chunkCount: chunks.length,
    issuanceFileName: payload.issuanceFileName || "",
    erpFileName: payload.erpFileName || "",
    savedAt: Date.now(),
    totalIssue: payload.totalIssue,
    totalErp: payload.totalErp,
    resultsCount: payload.results?.length || 0,
  });

  if (oldId && oldId !== uploadId) {
    try {
      const oldChunks = await getDocs(
        collection(db, "reconciliationReportBackups", oldId, "chunks")
      );
      for (const cd of oldChunks.docs) await deleteDoc(cd.ref);
    } catch {}
  }
};

const loadReconciliationReport = async () => {
  try {
    const metaRef = doc(db, "reconciliationReports", "latest");
    const snap = await getDoc(metaRef);
    if (!snap.exists()) return null;

    const meta = snap.data();
    if (!meta.uploadId || !meta.chunkCount) return null;

    const chunksSnap = await getDocs(
      collection(db, "reconciliationReportBackups", meta.uploadId, "chunks")
    );
    if (chunksSnap.empty) return null;

    const sorted = chunksSnap.docs
      .map((d) => d.data())
      .sort((a, b) => a.index - b.index);

    const base64 = sorted.map((c) => c.data).join("");
    const jsonStr = decodeURIComponent(escape(atob(base64)));

    return JSON.parse(jsonStr);
  } catch {
    return null;
  }
};

// Helpers
const clean = (v) => String(v ?? "").trim();

const normalizeHeader = (v) =>
  clean(v).toLowerCase().replace(/[^a-z0-9]/g, "");

const normalizeLot = (v) => clean(v).toUpperCase();

const normalizeItemCode = (v) => clean(v).toUpperCase();

const isNonEmpty = (v) => clean(v) !== "";

const isSumRow = (lot, item) => {
  const txt = `${lot} ${item}`.toLowerCase();

  return (
    txt.includes("grand total") ||
    txt === "total" ||
    txt.startsWith("total ") ||
    txt.includes("subtotal") ||
    txt.includes("closing") ||
    txt.includes("opening") ||
    txt.includes("balance")
  );
};

const parseQuantity = (value) => {
  if (value === null || value === undefined || value === "") return null;

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  const c = clean(value).replace(/,/g, "");
  if (c === "") return null;

  const n = Number(c);
  return Number.isFinite(n) ? n : null;
};

const getColumnIndexes = (headers) => {
  const norm = headers.map(normalizeHeader);

  const find = (aliases) => {
    for (const alias of aliases) {
      const na = normalizeHeader(alias);
      const idx = norm.findIndex((h) => h === na);
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const batching = find(["Batching Lot", "ERP Lot #", "ERP Lot"]);
  const lot =
    batching !== -1
      ? batching
      : find(["Lot Number", "Lot No", "Lot"]);

  const item = find(["ITEM CODE", "Item Code", "Item"]);
  const qty = find(["Issue Qty", "QTY", "Qty"]);

  return { lot, item, qty, batching };
};

// READ EXCEL WITH ROW-WISE DIAGNOSTICS.
const readExcel = async (file, type) => {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
  });

  console.group(`RECONCILIATION DIAGNOSTICS: ${type}`);

  console.log("File:", file.name);
  console.log("All Excel Sheets:", workbook.SheetNames);

  let sheetOrder = [...workbook.SheetNames];

  if (type === "issuance") {
    const idx = sheetOrder.findIndex((n) =>
      n.toLowerCase().includes("oct")
    );

    if (idx > 0) {
      const [s] = sheetOrder.splice(idx, 1);
      sheetOrder.unshift(s);
    }
  } else {
    const idx = sheetOrder.findIndex((n) => {
      const name = n.toLowerCase();
      return (
        name.includes("sub inventory") ||
        name.includes("sub inventry")
      );
    });

    if (idx > 0) {
      const [s] = sheetOrder.splice(idx, 1);
      sheetOrder.unshift(s);
    }
  }

  let targetSheetName = sheetOrder[0];

  for (const name of sheetOrder) {
    if (
      name.toLowerCase().includes("sheet2") ||
      name.toLowerCase().includes("205 code") ||
      name.toLowerCase().includes("sheet1") ||
      name.toLowerCase().includes("sheet3")
    ) {
      continue;
    }

    targetSheetName = name;
    break;
  }

  console.log("SELECTED SHEET:", targetSheetName);

  const sheet = workbook.Sheets[targetSheetName];

  if (!sheet) {
    console.groupEnd();
    throw new Error("Selected Excel sheet not found");
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: true,
  });

  for (
    let headerRow = 0;
    headerRow < Math.min(rows.length, 40);
    headerRow++
  ) {
    const headers = rows[headerRow].map(clean);

    if (headers.filter((h) => h).length < 2) continue;

    const indexes = getColumnIndexes(headers);

    if (indexes.qty === -1 || indexes.lot === -1) continue;

    console.log("HEADER ROW:", headerRow + 1);
    console.log("DETECTED HEADERS:", headers);
    console.log("COLUMN INDEXES:", indexes);
    console.log("LOT COLUMN:", headers[indexes.lot]);
    console.log(
      "BATCHING COLUMN:",
      indexes.batching !== -1
        ? headers[indexes.batching]
        : "Not found"
    );
    console.log("QUANTITY COLUMN:", headers[indexes.qty]);

    const records = [];
    const lotDiagnostics = new Map();

    let totalQuantity = 0;
    let validRows = 0;
    let exclBlank = 0;
    let exclSum = 0;
    let exclQty = 0;

    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i];

      if (!row || row.every((c) => clean(c) === "")) continue;

    const rawLot =
  type === "erp"
    ? row[indexes.batching]
    : indexes.batching !== -1 &&
      isNonEmpty(row[indexes.batching])
    ? row[indexes.batching]
    : row[indexes.lot];

      const rawItem =
        indexes.item !== -1 ? row[indexes.item] : "";

      const rawQty = row[indexes.qty];

      const lotDisp = clean(rawLot);
      const itemDisp = clean(rawItem);

      if (!isNonEmpty(lotDisp)) {
        exclBlank++;
        continue;
      }

      if (isSumRow(lotDisp, itemDisp)) {
        exclSum++;
        continue;
      }

      const qty = parseQuantity(rawQty);

      if (qty === null) {
        exclQty++;
        continue;
      }

      const normalizedLot = normalizeLot(lotDisp);
      const excelRow = i + 1;

      totalQuantity += qty;
      validRows++;

      records.push({
        lot: normalizedLot,
        lotDisplay: lotDisp,
        itemCode: itemDisp || "",
        normalizedCode: normalizeItemCode(itemDisp || ""),
        quantity: qty,
      });

      if (!lotDiagnostics.has(normalizedLot)) {
        lotDiagnostics.set(normalizedLot, {
          lot: normalizedLot,
          rowCount: 0,
          total: 0,
          rows: [],
        });
      }

      const diagnostic = lotDiagnostics.get(normalizedLot);

      diagnostic.rowCount++;
      diagnostic.total += qty;

      diagnostic.rows.push({
        ExcelRow: excelRow,
        Quantity: qty,
        Lot: lotDisp,
        RawQuantity: rawQty,
      });
    }

    if (validRows > 0) {
      console.log("TOTAL QUANTITY:", totalQuantity);
      console.log("VALID ROWS:", validRows);
      console.log("EXCLUDED BLANK LOT:", exclBlank);
      console.log("EXCLUDED SUM ROWS:", exclSum);
      console.log("EXCLUDED INVALID QTY:", exclQty);

      console.log("ALL LOT-WISE TOTALS:");

      console.table(
        [...lotDiagnostics.values()].map((item) => ({
          Lot: item.lot,
          Rows: item.rowCount,
          Total: item.total,
          ExcelRows: item.rows
            .map((r) => `${r.ExcelRow}: ${r.Quantity}`)
            .join(" | "),
        }))
      );

      if (type === "issuance") {
        const examples = [
          "PO-2026-49996",
          "PO-2026-49997",
          "PO-2026-49998",
          "PO-2026-49999",
        ];

        examples.forEach((lot) => {
          const detail = lotDiagnostics.get(lot);

          if (detail) {
            console.log(`DETAIL FOR ${lot}:`);
            console.table(detail.rows);
            console.log("LOT TOTAL:", detail.total);
            console.log("NUMBER OF ROWS:", detail.rowCount);
          } else {
            console.warn(`${lot} NOT FOUND IN SELECTED SHEET`);
          }
        });
      }

      console.groupEnd();

      return {
        records,
        sheetName: targetSheetName,
        totalQuantity,
        validRows,
        exclBlank,
        exclSum,
        exclQty,
      };
    }
  }

  console.groupEnd();

  throw new Error(
    "Lot/Qty headers or valid rows not found in selected sheet"
  );
};

// SUM ALL ROWS AGAINST THE SAME LOT.
const aggregateByLot = (records) => {
  const map = new Map();

  records.forEach((r) => {
    if (!map.has(r.lot)) {
      map.set(r.lot, {
        lot: r.lot,
        lotDisplay: r.lotDisplay,
        total: 0,
      });
    }

    const lot = map.get(r.lot);
    lot.total += r.quantity;
  });

  return map;
};

// COMPARE ONLY LOT + TOTAL QUANTITY.
// COMPARE ONLY ISSUANCE LOTS + ERP BATCHING LOT QUANTITY.
const compareData = (issuanceRecords, erpRecords) => {
  const issuance = aggregateByLot(issuanceRecords);
  const erp = aggregateByLot(erpRecords);

  // Sirf Issuance ke lots results mein show honge.
  // ERP ke extra lots ignore honge.
  return [...issuance.keys()]
    .map((lot) => {
      const issue = issuance.get(lot);
      const inv = erp.get(lot);

      const issueTotal = issue?.total ?? 0;
      const erpTotal = inv?.total ?? 0;
      const difference = issueTotal - erpTotal;

      const missingInErp = !inv;
      const quantityMismatch = Math.abs(difference) > 0.01;

      return {
        lot,
        lotDisplay: issue?.lotDisplay || lot,
        issueTotal,
        erpTotal,
        difference,
        issueCodes: [],
        erpCodes: [],
        codeDifferences: [],
        codeMismatch: false,
        quantityMismatch,
        missingInIssuance: false,
        missingInErp,
        matched: !quantityMismatch && !missingInErp,
      };
    })
    .sort((a, b) => a.lot.localeCompare(b.lot));
};
const formatNumber = (v) =>
  Number(v.toFixed(2)).toLocaleString("en-US");

export default function Reconciliation({ onHome }) {
  const [issuanceFile, setIssuanceFile] = useState(null);
  const [erpFile, setErpFile] = useState(null);
  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [showOnlyDifferences, setShowOnlyDifferences] = useState(false);
  const [issuanceSheet, setIssuanceSheet] = useState("");
  const [erpSheet, setErpSheet] = useState("");
  const [totalIssue, setTotalIssue] = useState(0);
  const [totalErp, setTotalErp] = useState(0);
  const [hasReconciled, setHasReconciled] = useState(false);
  const [restoringFiles, setRestoringFiles] = useState(true);
  const [savingFiles, setSavingFiles] = useState({
    issuance: false,
    erp: false,
    report: false,
  });
  const [storageStatus, setStorageStatus] = useState({
    issuance: "",
    erp: "",
    report: "",
  });

  useEffect(() => {
    let cancelled = false;

    const restoreAll = async () => {
      try {
        const [savedIssuance, savedErp, savedReport] =
          await Promise.all([
            loadReconciliationFile("issuance"),
            loadReconciliationFile("erp"),
            loadReconciliationReport(),
          ]);

        if (cancelled) return;

        if (savedIssuance) {
          setIssuanceFile(savedIssuance);
          setStorageStatus((p) => ({
            ...p,
            issuance: `Loaded: ${savedIssuance.name}`,
          }));
        }

        if (savedErp) {
          setErpFile(savedErp);
          setStorageStatus((p) => ({
            ...p,
            erp: `Loaded: ${savedErp.name}`,
          }));
        }

        if (savedReport) {
          setResults(savedReport.results || []);
          setTotalIssue(savedReport.totalIssue || 0);
          setTotalErp(savedReport.totalErp || 0);
          setIssuanceSheet(savedReport.issuanceSheet || "");
          setErpSheet(savedReport.erpSheet || "");
          setHasReconciled(true);

          setStorageStatus((p) => ({
            ...p,
            report: `Restored ${savedReport.results?.length || 0} lots`,
          }));

          setSuccess("Last saved report restored");

          setTimeout(() => setSuccess(""), 4000);
        }
      } catch (e) {
        if (!cancelled) {
          setError("Restore failed: " + e.message);
        }
      } finally {
        if (!cancelled) setRestoringFiles(false);
      }
    };

    restoreAll();

    return () => {
      cancelled = true;
    };
  }, []);

  const handleFileSelect = async (file, type) => {
    if (!file) return;

    if (type === "issuance") setIssuanceFile(file);
    else setErpFile(file);

    setError("");
    setSuccess("");
    setResults([]);
    setHasReconciled(false);
    setIssuanceSheet("");
    setErpSheet("");
    setTotalIssue(0);
    setTotalErp(0);

    setSavingFiles((p) => ({ ...p, [type]: true }));
    setStorageStatus((p) => ({ ...p, [type]: "Saving..." }));

    try {
      await saveReconciliationFile(file, type);

      setStorageStatus((p) => ({
        ...p,
        [type]: `Saved: ${file.name}`,
      }));

      setSuccess(`${type} file saved - Firebase`);
    } catch (err) {
      setStorageStatus((p) => ({
        ...p,
        [type]: `Save FAILED: ${err.message}`,
      }));

      setError(
        `Save failed: ${err.message}. Check Firestore Rules for reconciliationFiles & reconciliationBackups`
      );
    } finally {
      setSavingFiles((p) => ({ ...p, [type]: false }));
    }
  };

  const handleReconcile = async () => {
    if (!issuanceFile || !erpFile) {
      setError("Dono files select karein");
      return;
    }

    setLoading(true);
    setError("");
    setSuccess("");
    setResults([]);
    setHasReconciled(false);

    try {
      const issData = await readExcel(issuanceFile, "issuance");
      const erpData = await readExcel(erpFile, "erp");

      setIssuanceSheet(issData.sheetName);
      setErpSheet(erpData.sheetName);

      setTotalIssue(issData.totalQuantity);
      setTotalErp(erpData.totalQuantity);

      const comp = compareData(
        issData.records,
        erpData.records
      );

      setResults(comp);
      setHasReconciled(true);

      setSavingFiles((p) => ({ ...p, report: true }));

      try {
        await saveReconciliationReport({
          issuanceFileName: issuanceFile.name,
          erpFileName: erpFile.name,
          issuanceSheet: issData.sheetName,
          erpSheet: erpData.sheetName,
          totalIssue: issData.totalQuantity,
          totalErp: erpData.totalQuantity,
          results: comp,
          savedAt: Date.now(),
        });

        setStorageStatus((p) => ({
          ...p,
          report: `Saved ${comp.length} lots`,
        }));

        setSuccess(
          `Reconciliation complete: Issuance ${issData.validRows} rows, Total ${formatNumber(issData.totalQuantity)} | ERP ${erpData.validRows} rows, Total ${formatNumber(erpData.totalQuantity)}`
        );
      } catch (e) {
        setError("Report save fail: " + e.message);
      } finally {
        setSavingFiles((p) => ({ ...p, report: false }));
      }
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const differences = results.filter((r) => !r.matched);
  const displayed = showOnlyDifferences ? differences : results;
  const matchedCount = results.length - differences.length;

  return (
    <div className="recon-page">
      
      <div className="recon-header">
  <div>
    <h1>Reconciliation - Lot Wise</h1>
    <p>
      Independent: Only Lot Number + Issue Quantity | All rows summed once |
      Firebase + Responsive
    </p>
  </div>

  <button
    type="button"
    className="nav-home"
    onClick={() => {
      if (onHome) {
        onHome();
      } else {
        window.location.reload();
      }
    }}
  >
    ⌂ Home
  </button>
</div>

      {(restoringFiles ||
        savingFiles.issuance ||
        savingFiles.erp ||
        savingFiles.report ||
        loading) && (
        <div className="recon-status-bar">
          {restoringFiles && (
            <span className="status loading">
              🔄 Restoring files & report...
            </span>
          )}
          {savingFiles.issuance && (
            <span className="status saving">
              💾 Saving Issuance to Firebase...
            </span>
          )}
          {savingFiles.erp && (
            <span className="status saving">
              💾 Saving ERP to Firebase...
            </span>
          )}
          {savingFiles.report && (
            <span className="status saving">
              💾 Saving report to Firebase...
            </span>
          )}
          {loading && (
            <span className="status loading">
              ⏳ Reconciling - Summing all Lot rows...
            </span>
          )}
        </div>
      )}

      {error && <div className="recon-error">❌ {error}</div>}
      {success && <div className="recon-success">✅ {success}</div>}

      <div className="recon-upload-grid">
        <div className="recon-upload-card">
          <label>Issuance File (Daily_Greige_Issuance_Report_OCT)</label>
          <small>
            Uses: ERP Lot # + Issue Qty | All rows of same Lot are summed
          </small>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={(e) =>
              handleFileSelect(e.target.files?.[0], "issuance")
            }
          />
          <span className="recon-file-name">
            {issuanceFile ? issuanceFile.name : "Select Excel"}
          </span>
          {storageStatus.issuance && (
            <small className="recon-storage-status">
              {storageStatus.issuance}
            </small>
          )}
        </div>

        <div className="recon-upload-card">
          <label>ERP File (SUBINVENTRY_31.10.26)</label>
          <small>
            Uses: Batching Lot + QTY | All rows of same Lot are summed
          </small>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={(e) =>
              handleFileSelect(e.target.files?.[0], "erp")
            }
          />
          <span className="recon-file-name">
            {erpFile ? erpFile.name : "Select Excel"}
          </span>
          {storageStatus.erp && (
            <small className="recon-storage-status">
              {storageStatus.erp}
            </small>
          )}
        </div>
      </div>

      <button
        className="recon-run-button"
        onClick={handleReconcile}
        disabled={
          loading ||
          restoringFiles ||
          savingFiles.issuance ||
          savingFiles.erp
        }
      >
        {restoringFiles
          ? "Loading saved files..."
          : loading
          ? "Checking..."
          : "Run Reconciliation"}
      </button>

      {/* WORK IN PROGRESS NOTICE */}
      <div className="recon-development-notice">
        <div className="recon-development-title">
          ⚠ WORK IN PROGRESS
        </div>
        <div className="recon-development-heading">
          THIS PAGE IS UNDER DEVELOPMENT
        </div>
        <p>
          This page is not finalized yet. Some results may not be accurate.
          Please use it for testing purposes only.
        </p>
      </div>

      {hasReconciled && (
        <div className="recon-total-box">
          <div>
            <span>Total Issuance (All Rows Sum)</span>
            <strong>{formatNumber(totalIssue)} Mtr</strong>
          </div>
          <div>
            <span>Total ERP (All Rows Sum)</span>
            <strong>{formatNumber(totalErp)} Mtr</strong>
          </div>
          <div>
            <span>Difference</span>
            <strong
              className={
                Math.abs(totalIssue - totalErp) > 0.01
                  ? "recon-cell-error"
                  : "recon-cell-success"
              }
            >
              {formatNumber(totalIssue - totalErp)} Mtr
            </strong>
          </div>
        </div>
      )}

      {results.length > 0 && (
        <>
          <div className="recon-summary">
            <div className="recon-summary-card">
              <span>Total Lots</span>
              <strong>{results.length}</strong>
            </div>
            <div className="recon-summary-card">
              <span>Matched</span>
              <strong className="recon-green">{matchedCount}</strong>
            </div>
            <div className="recon-summary-card">
              <span>Differences</span>
              <strong className="recon-red">{differences.length}</strong>
            </div>
          </div>

          <div className="recon-results-header">
            <h2>
              Lot Wise - Single Row Per Lot{" "}
              {storageStatus.report && (
                <small>• {storageStatus.report}</small>
              )}
            </h2>

            <label className="recon-filter">
              <input
                type="checkbox"
                checked={showOnlyDifferences}
                onChange={(e) =>
                  setShowOnlyDifferences(e.target.checked)
                }
              />{" "}
              Differences only
            </label>
          </div>

          <div className="recon-table-wrapper desktop-only">
            <table className="recon-table">
              <thead>
                <tr>
                  <th>Lot Number (PO)</th>
                  <th>Issuance Qty (All Rows Sum)</th>
                  <th>ERP Qty (All Rows Sum)</th>
                  <th>Difference</th>
                  <th>Status</th>
                </tr>
              </thead>

              <tbody>
                {displayed.map((row) => (
                  <tr key={row.lot}>
                    <td className="cell-wrap" title={row.lotDisplay}>
                      {row.lotDisplay}
                    </td>
                    <td>{formatNumber(row.issueTotal)}</td>
                    <td>{formatNumber(row.erpTotal)}</td>
                    <td
                      className={
                        row.quantityMismatch
                          ? "recon-cell-error"
                          : "recon-cell-success"
                      }
                    >
                      {formatNumber(row.difference)}
                    </td>
                    <td>
                      {row.matched ? (
                        <span className="recon-badge recon-badge-green">
                          Matched
                        </span>
                      ) : (
                        <span className="recon-badge recon-badge-red">
                          Diff
                        </span>
                      )}

                      {row.missingInIssuance && (
                        <div className="recon-status-detail">
                          Missing in Issuance
                        </div>
                      )}

                      {row.missingInErp && (
                        <div className="recon-status-detail">
                          Missing in ERP
                        </div>
                      )}
                    </td>
                  </tr>
                ))}

                {displayed.length === 0 && (
                  <tr>
                    <td colSpan="5" className="recon-empty">
                      No differences - All lots matched
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="recon-mobile-cards mobile-only">
            {displayed.length === 0 ? (
              <div className="recon-mobile-empty">
                No differences - All matched
              </div>
            ) : (
              displayed.map((row) => (
                <div key={row.lot} className="recon-mobile-card">
                  <div className="recon-mobile-card-header">
                    <div>
                      <span className="recon-mobile-label">
                        LOT NUMBER
                      </span>
                      <strong className="recon-mobile-lot cell-wrap">
                        {row.lotDisplay}
                      </strong>
                    </div>

                    <span
                      className={`recon-badge ${
                        row.matched
                          ? "recon-badge-green"
                          : "recon-badge-red"
                      }`}
                    >
                      {row.matched ? "Matched" : "Diff"}
                    </span>
                  </div>

                  <div className="recon-mobile-quantities">
                    <div>
                      <span>Issuance (All Rows Sum)</span>
                      <strong>
                        {formatNumber(row.issueTotal)} Mtr
                      </strong>
                    </div>

                    <div>
                      <span>ERP (All Rows Sum)</span>
                      <strong>
                        {formatNumber(row.erpTotal)} Mtr
                      </strong>
                    </div>
                  </div>

                  <div className="recon-mobile-difference">
                    <span>Difference</span>
                    <strong
                      className={
                        row.quantityMismatch
                          ? "recon-cell-error"
                          : "recon-cell-success"
                      }
                    >
                      {formatNumber(row.difference)} Mtr
                    </strong>
                  </div>

                  {(row.missingInIssuance || row.missingInErp) && (
                    <div className="recon-mobile-status-details">
                      {row.missingInIssuance && (
                        <span>Missing in Issuance</span>
                      )}
                      {row.missingInErp && (
                        <span>Missing in ERP</span>
                      )}
                    </div>
                  )}
                </div>
              ))
            )}
          </div>

          <div className="recon-sheet-info">
            Sheets: Issuance={issuanceSheet} | ERP={erpSheet} |
            Logic: Lot + Qty only, all rows summed once | Firebase:{" "}
            {storageStatus.issuance} | {storageStatus.erp} |{" "}
            {storageStatus.report}
          </div>
        </>
      )}
    </div>
  );
}