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

const normalizeItemCode = (v) =>
  clean(v)
    .toUpperCase()
    .replace(/\./g, "")
    .replace(/0+$/, "");

const normalizeContract = (v) => clean(v).toUpperCase().replace(/\s+/g, "");

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

const getIssueDateInfo = (value) => {
  if (value === null || value === undefined || value === "") {
    return { key: "", display: "" };
  }

  let date = null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    date = value;
  } else if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const y = String(parsed.y).padStart(4, "0");
      const m = String(parsed.m).padStart(2, "0");
      const d = String(parsed.d).padStart(2, "0");
      return { key: `${y}-${m}-${d}`, display: `${d}-${m}-${y}` };
    }
  } else {
    const text = clean(value);
    const dmy = text.match(/^(\d{1,2})[-/ ]([A-Za-z]{3,}|\d{1,2})[-/ ](\d{4})$/);

    if (dmy) {
      const monthNames = {
        jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
        jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
      };
      const monthToken = dmy[2].toLowerCase();
      const month = /^\d+$/.test(monthToken)
        ? Number(monthToken)
        : monthNames[monthToken.slice(0, 3)];
      const day = Number(dmy[1]);
      const year = Number(dmy[3]);

      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        const y = String(year).padStart(4, "0");
        const m = String(month).padStart(2, "0");
        const d = String(day).padStart(2, "0");
        return { key: `${y}-${m}-${d}`, display: `${d}-${m}-${y}` };
      }
    }

    const parsedDate = new Date(text);
    if (!Number.isNaN(parsedDate.getTime())) date = parsedDate;
  }

  if (!date) {
    return { key: "", display: clean(value) };
  }

  const y = String(date.getFullYear()).padStart(4, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return { key: `${y}-${m}-${d}`, display: `${d}-${m}-${y}` };
};

const getColumnIndexes = (headers, type) => {
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

  const item = find(["ITEM CODE", "Item Code", "Item", "ItemCode"]);

  const contract = find([
    "ERP Contract #",
    "Contract #",
    "Sale Order #",
    "Sale Order",
    "Contract",
    "ERP Contract",
    "NEW ERP Contract #",
    "OLD ERP Contract #",
  ]);

  const qty =
    type === "erp"
      ? find(["QTY", "Qty", "Issue Qty"])
      : find(["Issue Qty", "QTY", "Qty"]);

  const issueDetails = type === "issuance"
    ? {
        greigeSupplier: find(["Greige Supplier"]),
        customer: find(["Customer"]),
        quality: find(["Quality"]),
        greigeWidth: find(["Greige Width"]),
        weaveBlend: find(["Weave Blend"]),
        date: find(["Date"]),
      }
    : {};

  return { lot, item, qty, batching, contract, ...issueDetails };
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

    const indexes = getColumnIndexes(headers, type);

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
    console.log("ITEM CODE COLUMN:", indexes.item !== -1 ? headers[indexes.item] : "Not found");
    console.log("CONTRACT COLUMN:", indexes.contract !== -1 ? headers[indexes.contract] : "Not found");
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
          ? indexes.batching !== -1 && isNonEmpty(row[indexes.batching])
            ? row[indexes.batching]
            : row[indexes.lot]
          : indexes.batching !== -1 &&
            isNonEmpty(row[indexes.batching])
          ? row[indexes.batching]
          : row[indexes.lot];

      const rawItem =
        indexes.item !== -1 ? row[indexes.item] : "";

      const rawContract = indexes.contract !== -1 ? row[indexes.contract] : "";

      const rawQty = row[indexes.qty];

      const lotDisp = clean(rawLot);
      const itemDisp = clean(rawItem);
      const contractDisp = clean(rawContract);

      const issueQuality = type === "issuance" && indexes.quality !== -1
        ? clean(row[indexes.quality])
        : "";
      const missingLotNumber = type === "issuance" && !isNonEmpty(lotDisp) && isNonEmpty(issueQuality);

      // Issuance rows without a Lot Number are included only when Quality is filled.
      if (!isNonEmpty(lotDisp) && !missingLotNumber) {
        exclBlank++;
        continue;
      }

      if (isNonEmpty(lotDisp) && isSumRow(lotDisp, itemDisp)) {
        exclSum++;
        continue;
      }

      const qty = parseQuantity(rawQty);

      if (qty === null) {
        exclQty++;
        continue;
      }

      const excelRow = i + 1;
      const normalizedLot = missingLotNumber
        ? `__LOT_MISSING_ROW_${excelRow}`
        : normalizeLot(lotDisp);

      const issueField = (index) =>
        type === "issuance" && index !== -1 ? clean(row[index]) : "";
      const issueDate = getIssueDateInfo(
        type === "issuance" && indexes.date !== -1 ? row[indexes.date] : ""
      );

      totalQuantity += qty;
      validRows++;

      records.push({
        lot: normalizedLot,
        lotDisplay: missingLotNumber ? "Lot Number Miss" : lotDisp,
        missingLotNumber,
        itemCode: itemDisp || "",
        normalizedCode: normalizeItemCode(itemDisp || ""),
        contract: contractDisp || "",
        normalizedContract: normalizeContract(contractDisp || ""),
        contractDisplay: contractDisp || "",
        greigeSupplier: issueField(indexes.greigeSupplier),
        customer: issueField(indexes.customer),
        quality: issueField(indexes.quality),
        greigeWidth: issueField(indexes.greigeWidth),
        weaveBlend: issueField(indexes.weaveBlend),
        issueDate: issueDate.display,
        issueDateKey: issueDate.key,
        quantity: qty,
      });

      if (!lotDiagnostics.has(normalizedLot)) {
        lotDiagnostics.set(normalizedLot, {
          lot: normalizedLot,
          rowCount: 0,
          total: 0,
          rows: [],
          contracts: new Set(),
          itemCodes: new Set(),
        });
      }

      const diagnostic = lotDiagnostics.get(normalizedLot);

      diagnostic.rowCount++;
      diagnostic.total += qty;
      if (contractDisp) diagnostic.contracts.add(contractDisp);
      if (itemDisp) diagnostic.itemCodes.add(itemDisp);

      diagnostic.rows.push({
        ExcelRow: excelRow,
        Quantity: qty,
        Lot: lotDisp,
        Contract: contractDisp,
        ItemCode: itemDisp,
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
          Contracts: [...item.contracts].join(", "),
          ItemCodes: [...item.itemCodes].join(", "),
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
        contracts: new Set(),
        contractDisplays: new Set(),
        normalizedContracts: new Set(),
        itemCodes: new Set(),
        itemCodeDisplays: new Set(),
        normalizedCodes: new Set(),
        greigeSuppliers: new Set(),
        customers: new Set(),
        qualities: new Set(),
        greigeWidths: new Set(),
        weaveBlends: new Set(),
        issueDates: new Set(),
        issueDateKeys: new Set(),
        rows: [],
      });
    }

    const lot = map.get(r.lot);
    lot.total += r.quantity;
    if (r.contract) {
      lot.contracts.add(r.contract);
      lot.contractDisplays.add(r.contractDisplay || r.contract);
      lot.normalizedContracts.add(r.normalizedContract);
    }
    if (r.itemCode) {
      lot.itemCodes.add(r.itemCode);
      lot.itemCodeDisplays.add(r.itemCode);
      lot.normalizedCodes.add(r.normalizedCode);
    }
    if (r.greigeSupplier) lot.greigeSuppliers.add(r.greigeSupplier);
    if (r.customer) lot.customers.add(r.customer);
    if (r.quality) lot.qualities.add(r.quality);
    if (r.greigeWidth) lot.greigeWidths.add(r.greigeWidth);
    if (r.weaveBlend) lot.weaveBlends.add(r.weaveBlend);
    if (r.issueDate) lot.issueDates.add(r.issueDate);
    if (r.issueDateKey) lot.issueDateKeys.add(r.issueDateKey);
    lot.rows.push(r);
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

      // Missing-lot issuance rows are shown as their own Difference cards.
      if (issue?.missingLotNumber) {
        const issueTotal = issue.total || 0;
        return {
          lot,
          lotDisplay: "Lot Number Miss",
          issueTotal,
          erpTotal: 0,
          difference: issueTotal,
          issueContracts: [...(issue.contracts || [])],
          erpContracts: [],
          issueCodes: [...(issue.itemCodes || [])],
          erpCodes: [],
          issueContractsDisplay: [...(issue.contracts || [])],
          erpContractsDisplay: [],
          issueCodesDisplay: [...(issue.itemCodes || [])],
          erpCodesDisplay: [],
          greigeSupplier: [...(issue.greigeSuppliers || [])].join(", "),
          customer: [...(issue.customers || [])].join(", "),
          quality: [...(issue.qualities || [])].join(", "),
          greigeWidth: [...(issue.greigeWidths || [])].join(", "),
          weaveBlend: [...(issue.weaveBlends || [])].join(", "),
          issueDate: [...(issue.issueDates || [])].join(", "),
          issueDateKeys: [...(issue.issueDateKeys || [])],
          codeDifferences: ["Lot Number Miss"],
          contractMismatch: false,
          codeMismatch: false,
          quantityMismatch: false,
          missingLotNumber: true,
          missingInIssuance: false,
          missingInErp: false,
          matched: false,
        };
      }

      const issueTotal = issue?.total ?? 0;
      const erpTotal = inv?.total ?? 0;
      const difference = issueTotal - erpTotal;

      const missingInErp = !inv;
      const quantityMismatch = Math.abs(difference) > 0.01;

      // Contract & Item Code logic
      const issueContracts = [...(issue?.contracts || [])];
      const erpContracts = [...(inv?.contracts || [])];
      const issueCodes = [...(issue?.itemCodes || [])];
      const erpCodes = [...(inv?.itemCodes || [])];

      const issueNormContracts = [...(issue?.normalizedContracts || [])];
      const erpNormContracts = [...(inv?.normalizedContracts || [])];
      const issueNormCodes = [...(issue?.normalizedCodes || [])];
      const erpNormCodes = [...(inv?.normalizedCodes || [])];

      // Check contract mismatch: if any contract in issuance not in ERP or vice versa
      let contractMismatch = false;
      if (!missingInErp) {
        const setIssue = new Set(issueNormContracts);
        const setErp = new Set(erpNormContracts);
        if (setIssue.size !== setErp.size || [...setIssue].some(v => !setErp.has(v))) {
          contractMismatch = true;
        }
      }

      let codeMismatch = false;
      if (!missingInErp) {
        const setIssue = new Set(issueNormCodes);
        const setErp = new Set(erpNormCodes);
        if (setIssue.size !== setErp.size || [...setIssue].some(v => !setErp.has(v))) {
          codeMismatch = true;
        }
      }

    console.log("ITEM CODE MATCH CHECK:", {
  lot,
  issueRaw: issueCodes,
  erpRaw: erpCodes,
  issueNormalized: issueCodes.map(normalizeItemCode),
  erpNormalized: erpCodes.map(normalizeItemCode),
});

const codeDifferences = [];
if (contractMismatch) {
  codeDifferences.push(`Contract: Issue[${issueContracts.join(', ')}] vs ERP[${erpContracts.join(', ')}]`);
}
if (codeMismatch) {
  codeDifferences.push(`ItemCode: Issue[${issueCodes.join(', ')}] vs ERP[${erpCodes.join(', ')}]`);
}

const matched = !quantityMismatch && !missingInErp && !contractMismatch && !codeMismatch;
      return {
        lot,
        lotDisplay: issue?.lotDisplay || lot,
        issueTotal,
        erpTotal,
        difference,
        issueContracts,
        erpContracts,
        issueCodes,
        erpCodes,
        issueContractsDisplay: issueContracts,
        erpContractsDisplay: erpContracts,
        issueCodesDisplay: issueCodes,
        erpCodesDisplay: erpCodes,
        greigeSupplier: [...(issue?.greigeSuppliers || [])].join(", "),
        customer: [...(issue?.customers || [])].join(", "),
        quality: [...(issue?.qualities || [])].join(", "),
        greigeWidth: [...(issue?.greigeWidths || [])].join(", "),
        weaveBlend: [...(issue?.weaveBlends || [])].join(", "),
        issueDate: [...(issue?.issueDates || [])].join(", "),
        issueDateKeys: [...(issue?.issueDateKeys || [])],
        codeDifferences,
        contractMismatch,
        codeMismatch,
        quantityMismatch,
        missingInIssuance: false,
        missingInErp,
        matched,
      };
    })
    .sort((a, b) => a.lot.localeCompare(b.lot));
};
const formatNumber = (v) =>
  Number(((Number(v) || 0).toFixed(2))).toLocaleString("en-US");

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
  // Search input values (search runs only when Search button is pressed)
  const [lotSearch, setLotSearch] = useState("");
  const [contractSearch, setContractSearch] = useState("");
  const [itemCodeSearch, setItemCodeSearch] = useState("");

  // Values currently applied to the results
  const [appliedLotSearch, setAppliedLotSearch] = useState("");
  const [appliedContractSearch, setAppliedContractSearch] = useState("");
  const [appliedItemCodeSearch, setAppliedItemCodeSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [appliedDateFrom, setAppliedDateFrom] = useState("");
  const [appliedDateTo, setAppliedDateTo] = useState("");

  const normalizeSearchText = (value) =>
    clean(value).toLowerCase().replace(/[^a-z0-9]/g, "");

  const handleSearch = () => {
    setAppliedLotSearch(normalizeSearchText(lotSearch));
    setAppliedContractSearch(normalizeSearchText(contractSearch));
    setAppliedItemCodeSearch(normalizeSearchText(itemCodeSearch));
    setAppliedDateFrom(dateFrom);
    setAppliedDateTo(dateTo);
  };

  const handleClearSearch = () => {
    setLotSearch("");
    setContractSearch("");
    setItemCodeSearch("");
    setAppliedLotSearch("");
    setAppliedContractSearch("");
    setAppliedItemCodeSearch("");
    setDateFrom("");
    setDateTo("");
    setAppliedDateFrom("");
    setAppliedDateTo("");
  };

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
          const restoredResults = (savedReport.results || [])
            .filter((row) => row && typeof row === "object")
            .map((row) => ({
              ...row,
              issueContracts: Array.isArray(row.issueContracts) ? row.issueContracts : [],
              erpContracts: Array.isArray(row.erpContracts) ? row.erpContracts : [],
              issueCodes: Array.isArray(row.issueCodes) ? row.issueCodes : [],
              erpCodes: Array.isArray(row.erpCodes) ? row.erpCodes : [],
              issueDateKeys: Array.isArray(row.issueDateKeys) ? row.issueDateKeys : [],
              issueTotal: Number(row.issueTotal) || 0,
              erpTotal: Number(row.erpTotal) || 0,
              difference: Number(row.difference) || 0,
              quantityMismatch: Boolean(row.quantityMismatch),
              contractMismatch: Boolean(row.contractMismatch),
              codeMismatch: Boolean(row.codeMismatch),
              matched: Boolean(row.matched),
            }));
          setResults(restoredResults);
          setTotalIssue(savedReport.totalIssue || 0);
          setTotalErp(savedReport.totalErp || 0);
          setIssuanceSheet(savedReport.issuanceSheet || "");
          setErpSheet(savedReport.erpSheet || "");
          setHasReconciled(true);

          setStorageStatus((p) => ({
            ...p,
            report: `Restored ${restoredResults.length} lots`,
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
  const contractDiffs = results.filter((r) => r.contractMismatch);
  const itemCodeDiffs = results.filter((r) => r.codeMismatch);
  
  // Apply filters only after the Search button is pressed.
  // Punctuation is ignored, so 51130 also matches PO-2026-51130.
  const filteredBySearch = results.filter((r) => {
    const lotText = normalizeSearchText(`${r.lot || ""} ${r.lotDisplay || ""}`);
    const contractText = normalizeSearchText(
      [...(r.issueContracts || []), ...(r.erpContracts || [])].join(" ")
    );
    const itemText = normalizeSearchText(
      [...(r.issueCodes || []), ...(r.erpCodes || [])].join(" ")
    );

    const lotMatch = !appliedLotSearch || lotText.includes(appliedLotSearch);
    const contractMatch =
      !appliedContractSearch || contractText.includes(appliedContractSearch);
    const itemMatch =
      !appliedItemCodeSearch || itemText.includes(appliedItemCodeSearch);

    const dateMatch =
      (!appliedDateFrom && !appliedDateTo) ||
      (r.issueDateKeys || []).some((dateKey) =>
        dateKey &&
        (!appliedDateFrom || dateKey >= appliedDateFrom) &&
        (!appliedDateTo || dateKey <= appliedDateTo)
      );

    return lotMatch && contractMatch && itemMatch && dateMatch;
  });

  const displayedBase = showOnlyDifferences ? differences : filteredBySearch;
  // If showOnlyDifferences is on, still apply search on differences
  const displayed = showOnlyDifferences 
    ? filteredBySearch.filter(r => !r.matched)
    : filteredBySearch;

  // Totals follow the currently applied Search filters (including the date range).
  const filteredTotalQty = filteredBySearch.reduce(
    (sum, row) => sum + (Number(row.issueTotal) || 0),
    0
  );
  const filteredDiffQty = filteredBySearch.reduce(
    (sum, row) => sum + (row.quantityMismatch ? Math.abs(Number(row.difference) || 0) : 0),
    0
  );

  const matchedCount = results.length - differences.length;

  return (
    <div className="recon-page">
      
      <div className="recon-header">
  <div>
    <h1>Reconciliation - Lot Wise</h1>
    <p>
      Lot + Contract # + Item Code Matching | All rows summed once |
      Firebase + Responsive + Lot Search Filter
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
          <label>Issuance File (Daily-Greige-Issuance-Report)</label>
        
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
          <label>ERP File (SUBINVENTRY)</label>
        
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
          THIS PAGE IS UNDER DEVELOPMENT - UPDATED WITH CONTRACT + ITEM CODE MATCH
        </div>
        <p>
          Updated: Now matching Lot + Contract # + Item Code. Different lots shown in Diff section. Lot Search Filter added.
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
              <span>Matched (Lot+Contract+Code+Qty)</span>
              <strong className="recon-green">{matchedCount}</strong>
            </div>
            <div className="recon-summary-card">
              <span>Differences (Any)</span>
              <strong className="recon-red">{differences.length}</strong>
            </div>
            <div className="recon-summary-card">
              <span>Contract Mismatch</span>
              <strong className="recon-red">{contractDiffs.length}</strong>
            </div>
            <div className="recon-summary-card">
              <span>Item Code Mismatch</span>
              <strong className="recon-red">{itemCodeDiffs.length}</strong>
            </div>
          </div>

          {/* Search Filters */}
          <div
            className="recon-search-grid"
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
              gap: "12px",
              margin: "16px 0",
            }}
          >
            <div className="recon-upload-card">
              <label>🔍 Lot Search Filter</label>
              <input
                type="text"
                placeholder="Enter full PO or last digits e.g. 51130"
                value={lotSearch}
                onChange={(e) => setLotSearch(e.target.value)}
                style={{ width: "100%", padding: "8px", borderRadius: "6px", border: "1px solid #ccc" }}
              />
            </div>
            <div className="recon-upload-card">
              <label>🔍 Contract Search</label>
              <input
                type="text"
                placeholder="Search Contract #"
                value={contractSearch}
                onChange={(e) => setContractSearch(e.target.value)}
                style={{ width: "100%", padding: "8px", borderRadius: "6px", border: "1px solid #ccc" }}
              />
            </div>
            <div className="recon-upload-card">
              <label>🔍 Item Code Search</label>
              <input
                type="text"
                placeholder="Search Item Code"
                value={itemCodeSearch}
                onChange={(e) => setItemCodeSearch(e.target.value)}
                style={{ width: "100%", padding: "8px", borderRadius: "6px", border: "1px solid #ccc" }}
              />
            </div>
            <div className="recon-upload-card">
              <label>📅 Date From</label>
              <input
                type="date"
                value={dateFrom}
                max={dateTo || undefined}
                onChange={(e) => setDateFrom(e.target.value)}
                style={{ width: "100%", padding: "8px", borderRadius: "6px", border: "1px solid #ccc" }}
              />
            </div>
            <div className="recon-upload-card">
              <label>📅 Date To</label>
              <input
                type="date"
                value={dateTo}
                min={dateFrom || undefined}
                onChange={(e) => setDateTo(e.target.value)}
                style={{ width: "100%", padding: "8px", borderRadius: "6px", border: "1px solid #ccc" }}
              />
            </div>
          </div>

          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", margin: "0 0 16px" }}>
            <button
              type="button"
              onClick={handleSearch}
              style={{
                padding: "10px 24px",
                border: "none",
                borderRadius: "6px",
                background: "#1769aa",
                color: "#fff",
                fontWeight: "700",
                cursor: "pointer",
              }}
            >
              Search
            </button>
            <button
              type="button"
              onClick={handleClearSearch}
              style={{
                padding: "10px 20px",
                border: "1px solid #9ca3af",
                borderRadius: "6px",
                background: "#fff",
                color: "#1f2937",
                fontWeight: "700",
                cursor: "pointer",
              }}
            >
              Clear Filters
            </button>
            <div
              style={{
                display: "flex",
                gap: "10px",
                flexWrap: "wrap",
                alignItems: "center",
                marginLeft: "auto",
              }}
            >
              <div
                style={{
                  padding: "9px 14px",
                  border: "1px solid #bfdbfe",
                  borderRadius: "7px",
                  background: "#eff6ff",
                  color: "#1e3a8a",
                  fontWeight: "700",
                }}
              >
                Total Qty: {formatNumber(filteredTotalQty)} Mtr
              </div>
              <div
                style={{
                  padding: "9px 14px",
                  border: "1px solid #fecaca",
                  borderRadius: "7px",
                  background: "#fef2f2",
                  color: "#991b1b",
                  fontWeight: "700",
                }}
              >
                Diff Qty: {formatNumber(filteredDiffQty)} Mtr
              </div>
            </div>
          </div>

          <div className="recon-results-header">
            <h2>
              Lot Wise - With Contract & Item Code Matching{" "}
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
              Differences only (Lot/Contract/Code/Qty)
            </label>
          </div>

          <div className="recon-table-wrapper desktop-only">
            <table className="recon-table">
              <thead>
                <tr>
                  <th>Lot Number (PO)</th>
                  <th>Issuance Contract #</th>
                  <th>ERP Contract #</th>
                  <th>Issuance Item Code</th>
                  <th>ERP Item Code</th>
                  <th>Issuance Qty (All Rows Sum)</th>
                  <th>ERP Qty (All Rows Sum)</th>
                  <th>Difference</th>
                  <th>Status</th>
                </tr>
              </thead>

              <tbody>
                {displayed.map((row) => (
                  <tr key={row.lot} className={row.matched ? '' : 'recon-row-diff'}>
                    <td className="cell-wrap" title={row.lotDisplay}>
                      {row.lotDisplay}
                    </td>
                    <td className={row.contractMismatch ? "recon-cell-error" : ""}>
                      {row.issueContracts.join(', ') || '-'}
                    </td>
                    <td className={row.contractMismatch ? "recon-cell-error" : ""}>
                      {row.erpContracts.join(', ') || '-'}
                    </td>
                    <td className={row.codeMismatch ? "recon-cell-warning" : ""} style={{maxWidth:'200px', overflow:'hidden', textOverflow:'ellipsis'}}>
                      {row.issueCodes.join(', ') || '-'}
                    </td>
                    <td className={row.codeMismatch ? "recon-cell-warning" : ""} style={{maxWidth:'200px', overflow:'hidden', textOverflow:'ellipsis'}}>
                      {row.erpCodes.join(', ') || '-'}
                    </td>
                    <td>{formatNumber(row.issueTotal)}</td>
                    <td>{formatNumber(row.erpTotal)}</td>
                    <td
                      className={
                        row.quantityMismatch || row.missingLotNumber
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
                        <div style={{display:'flex', flexDirection:'column', gap:'4px'}}>
                          <span className="recon-badge recon-badge-red">
                            Diff
                          </span>
                          {row.contractMismatch && <small style={{color:'#d32f2f'}}>Contract Diff</small>}
                          {row.codeMismatch && <small style={{color:'#ed6c02'}}>Code Diff</small>}
                          {row.quantityMismatch && <small style={{color:'#d32f2f'}}>Qty Diff</small>}
                          {row.missingLotNumber && <small style={{color:'#d32f2f', fontWeight:700}}>Lot Number Miss</small>}
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
                    <td colSpan="9" className="recon-empty">
                      No records found - Try clearing search filters
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="recon-mobile-cards mobile-only">
            {displayed.length === 0 ? (
              <div className="recon-mobile-empty">
                No records - Clear search filters
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

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
                      gap: "5px 8px",
                      margin: "7px 0",
                      padding: "7px",
                      borderRadius: "6px",
                      background: "rgba(100, 116, 139, 0.08)",
                      fontSize: "11px",
                      lineHeight: "1.25",
                    }}
                  >
                    {[
                      ["Greige Supplier", row.greigeSupplier],
                      ["Customer", row.customer],
                      ["Quality", row.quality],
                      ["Greige Width", row.greigeWidth],
                      ["Weave Blend", row.weaveBlend],
                      ["Date", row.issueDate],
                    ].map(([label, value]) => (
                      <div key={label} style={{ minWidth: 0, overflowWrap: "anywhere" }}>
                        <span style={{ display: "block", color: "#64748b", fontSize: "10px" }}>
                          {label}
                        </span>
                        <strong style={{ display: "block", fontSize: "11px", fontWeight: 700 }}>
                          {value || "-"}
                        </strong>
                      </div>
                    ))}
                  </div>

                  <div style={{display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px', margin:'8px 0', fontSize:'12px'}}>
                    <div>
                      <span>Issue Contract</span><br/>
                      <strong className={row.contractMismatch ? "recon-cell-error" : ""}>{row.issueContracts.join(', ') || '-'}</strong>
                    </div>
                    <div>
                      <span>ERP Contract</span><br/>
                      <strong className={row.contractMismatch ? "recon-cell-error" : ""}>{row.erpContracts.join(', ') || '-'}</strong>
                    </div>
                    <div style={{gridColumn:'1 / span 2'}}>
                      <span>Issue Code</span><br/>
                      <strong>{row.issueCodes.join(', ') || '-'}</strong>
                    </div>
                    <div style={{gridColumn:'1 / span 2'}}>
                      <span>ERP Code</span><br/>
                      <strong>{row.erpCodes.join(', ') || '-'}</strong>
                    </div>
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
                        row.quantityMismatch || row.missingLotNumber
                          ? "recon-cell-error"
                          : "recon-cell-success"
                      }
                    >
                      {formatNumber(row.difference)} Mtr
                    </strong>
                  </div>

                  {(row.contractMismatch || row.codeMismatch) && (
                    <div className="recon-mobile-status-details">
                      {row.contractMismatch && <span>Contract Mismatch</span>}
                      {row.codeMismatch && <span>Item Code Mismatch</span>}
                      {row.missingInErp && <span>Missing in ERP</span>}
                    </div>
                  )}
                </div>
              ))
            )}
          </div>

          <div className="recon-sheet-info">
            Sheets: Issuance={issuanceSheet} | ERP={erpSheet} |
            Logic: Lot + Contract # + Item Code + Qty | Firebase:{" "}
            {storageStatus.issuance} | {storageStatus.erp} |{" "}
            {storageStatus.report} | Filtered: {displayed.length}/{results.length}
          </div>
        </>
      )}
    </div>
  );
}
