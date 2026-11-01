
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

const CHUNK_SIZE = 800000;

const bytesToBase64 = (buffer) => {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const step = 0x8000;

  for (let i = 0; i < bytes.length; i += step) {
    binary += String.fromCharCode(...bytes.subarray(i, i + step));
  }

  return btoa(binary);
};

const base64ToBytes = (base64) => {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  return bytes;
};

// Firebase mein Excel file save karega
const saveReconciliationFile = async (file, type) => {
  const buffer = await file.arrayBuffer();
  const base64 = bytesToBase64(buffer);

  const uploadId = `recon_${type}_${Date.now()}_${Math.random()
    .toString(36)
    .slice(2, 8)}`;

  const chunks = [];

  for (let i = 0; i < base64.length; i += CHUNK_SIZE) {
    chunks.push(base64.slice(i, i + CHUNK_SIZE));
  }

  if (!chunks.length) chunks.push("");

  for (let i = 0; i < chunks.length; i++) {
    await setDoc(
      doc(
        db,
        "reconciliationBackups",
        uploadId,
        "chunks",
        String(i)
      ),
      {
        index: i,
        data: chunks[i],
      }
    );
  }

  const metadataRef = doc(db, "reconciliationFiles", type);
  const oldMetadata = await getDoc(metadataRef);

  const oldUploadId = oldMetadata.exists()
    ? oldMetadata.data().uploadId
    : null;

  await setDoc(metadataRef, {
    uploadId,
    fileName: file.name,
    chunkCount: chunks.length,
    updatedAt: Date.now(),
  });

  if (oldUploadId && oldUploadId !== uploadId) {
    const oldChunksRef = collection(
      db,
      "reconciliationBackups",
      oldUploadId,
      "chunks"
    );

    const oldChunks = await getDocs(oldChunksRef);

    for (const chunkDoc of oldChunks.docs) {
      await deleteDoc(chunkDoc.ref);
    }
  }
};

// Firebase se Excel file load karega
const loadReconciliationFile = async (type) => {
  const metadataRef = doc(db, "reconciliationFiles", type);
  const metadataSnap = await getDoc(metadataRef);

  if (!metadataSnap.exists()) return null;

  const metadata = metadataSnap.data();

  if (!metadata.uploadId || !metadata.chunkCount) {
    return null;
  }

  const chunksRef = collection(
    db,
    "reconciliationBackups",
    metadata.uploadId,
    "chunks"
  );

  const chunksSnap = await getDocs(chunksRef);

  const sortedChunks = chunksSnap.docs
    .map((item) => item.data())
    .sort((a, b) => a.index - b.index);

  if (sortedChunks.length !== metadata.chunkCount) {
    throw new Error(
      `${metadata.fileName} ke Firebase chunks incomplete hain.`
    );
  }

  const base64 = sortedChunks
    .map((chunk) => chunk.data)
    .join("");

  const bytes = base64ToBytes(base64);

  return new File(
    [bytes],
    metadata.fileName,
    { type: "application/octet-stream" }
  );
};

const clean = (value) => String(value ?? "").trim();

const normalizeHeader = (value) =>
  clean(value).toLowerCase().replace(/[^a-z0-9]/g, "");

const normalizeLot = (value) => clean(value).toUpperCase();

const normalizeItemCode = (value) =>
  clean(value).toUpperCase().replace(/(\.0000)+$/g, "");

const parseQuantity = (value) => {
  if (typeof value === "number") return value;

  const number = Number(clean(value).replace(/,/g, ""));

  return Number.isFinite(number) ? number : 0;
};

const findColumn = (headers, aliases) => {
  const normalized = headers.map(normalizeHeader);

  return normalized.findIndex((header) =>
    aliases.some((alias) => header === normalizeHeader(alias))
  );
};

const getColumnIndexes = (headers) => {
  const aliases = {
    lot: ["ERP Lot #", "Batching Lot", "Lot Number", "Lot No"],
    item: ["ITEM CODE", "Item Code", "Item Number"],
    qty: ["Issue Qty", "QTY", "Quantity", "Qty"],
  };

  return {
    lot: findColumn(headers, aliases.lot),
    item: findColumn(headers, aliases.item),
    qty: findColumn(headers, aliases.qty),
  };
};

// Excel read + quantity total
const readExcel = async (file, type) => {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
  });

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

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
      const indexes = getColumnIndexes(headers);

      // Quantity column aur kam az kam ek identifier column zaroori hai
      if (
        indexes.qty === -1 ||
        (indexes.lot === -1 && indexes.item === -1)
      ) {
        continue;
      }

      const records = [];
      let totalQuantity = 0;
      let validQuantityRows = 0;

      for (let i = headerRow + 1; i < rows.length; i++) {
        const row = rows[i];

        const rawQuantity = row[indexes.qty];

        const lot = normalizeLot(
          indexes.lot !== -1 ? row[indexes.lot] : ""
        );

        const itemCode = clean(
          indexes.item !== -1 ? row[indexes.item] : ""
        );

        // Dono khaali hon to row ignore hogi
        if (!lot && !itemCode) continue;

        // Quantity khaali ho to row ignore hogi
        if (
          rawQuantity === "" ||
          rawQuantity === null ||
          rawQuantity === undefined
        ) {
          continue;
        }

        const quantity = parseQuantity(rawQuantity);

        // OR condition:
        // Lot ho YA Item Code ho, quantity total mein add hogi
        totalQuantity += quantity;
        validQuantityRows++;

        // Comparison ke liye dono identifiers zaroori hain
        if (lot && itemCode) {
          records.push({
            lot,
            itemCode,
            normalizedCode: normalizeItemCode(itemCode),
            quantity,
          });
        }
      }

      // Totals ke liye complete lot/item matching zaroori nahi
      if (validQuantityRows > 0) {
        return {
          records,
          sheetName,
          totalQuantity,
        };
      }
    }
  }

  throw new Error(
    type === "issuance"
      ? "Issuance file mein Lot Number ya Item Code aur Quantity headings/valid rows nahi milin."
      : "ERP file mein Lot Number ya Item Code aur Quantity headings/valid rows nahi milin."
  );
};

const aggregateByLot = (records) => {
  const result = new Map();

  records.forEach((record) => {
    if (!result.has(record.lot)) {
      result.set(record.lot, {
        lot: record.lot,
        total: 0,
        codes: new Map(),
      });
    }

    const lot = result.get(record.lot);
    lot.total += record.quantity;

    const codeKey = record.normalizedCode;

    if (!lot.codes.has(codeKey)) {
      lot.codes.set(codeKey, {
        code: codeKey,
        originalCodes: new Set(),
        quantity: 0,
      });
    }

    const code = lot.codes.get(codeKey);
    code.quantity += record.quantity;
    code.originalCodes.add(record.itemCode);
  });

  return result;
};

const compareData = (issuanceRecords, erpRecords) => {
  const issuance = aggregateByLot(issuanceRecords);
  const erp = aggregateByLot(erpRecords);

  const allLots = new Set([
    ...issuance.keys(),
    ...erp.keys(),
  ]);

  return [...allLots]
    .map((lotNumber) => {
      const issue = issuance.get(lotNumber);
      const inventory = erp.get(lotNumber);

      const issueTotal = issue?.total ?? 0;
      const erpTotal = inventory?.total ?? 0;
      const difference = issueTotal - erpTotal;

      const issueCodes = new Set([
        ...(issue?.codes.keys() ?? []),
      ]);

      const erpCodes = new Set([
        ...(inventory?.codes.keys() ?? []),
      ]);

      const allCodes = new Set([
        ...issueCodes,
        ...erpCodes,
      ]);

      const codeDifferences = [...allCodes]
        .filter(
          (code) =>
            !issueCodes.has(code) || !erpCodes.has(code)
        )
        .map((code) => ({
          code,
          issuance: issue?.codes.get(code)?.quantity ?? 0,
          erp: inventory?.codes.get(code)?.quantity ?? 0,
        }));

      const codeMismatch = codeDifferences.length > 0;
      const quantityMismatch = Math.abs(difference) > 0.01;
      const missingInIssuance = !issue;
      const missingInErp = !inventory;

      return {
        lot: lotNumber,
        issueTotal,
        erpTotal,
        difference,
        issueCodes: [...issueCodes],
        erpCodes: [...erpCodes],
        codeDifferences,
        codeMismatch,
        quantityMismatch,
        missingInIssuance,
        missingInErp,
        matched:
          !quantityMismatch &&
          !codeMismatch &&
          !missingInIssuance &&
          !missingInErp,
      };
    })
    .sort((a, b) => a.lot.localeCompare(b.lot));
};

const formatNumber = (value) =>
  Number(value.toFixed(2)).toLocaleString("en-US");

export default function Reconciliation() {
  const [issuanceFile, setIssuanceFile] = useState(null);
  const [erpFile, setErpFile] = useState(null);

  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const [showOnlyDifferences, setShowOnlyDifferences] = useState(true);

  const [issuanceSheet, setIssuanceSheet] = useState("");
  const [erpSheet, setErpSheet] = useState("");

  const [totalIssue, setTotalIssue] = useState(0);
  const [totalErp, setTotalErp] = useState(0);

  // Totals show karne ke liye
  const [hasReconciled, setHasReconciled] = useState(false);

  const [restoringFiles, setRestoringFiles] = useState(true);

  const [savingFiles, setSavingFiles] = useState({
    issuance: false,
    erp: false,
  });

  const [storageStatus, setStorageStatus] = useState({
    issuance: "",
    erp: "",
  });

  // Page open hone par Firebase se files load hongi
  useEffect(() => {
    let cancelled = false;

    const restoreFiles = async () => {
      try {
        const [savedIssuance, savedErp] = await Promise.all([
          loadReconciliationFile("issuance"),
          loadReconciliationFile("erp"),
        ]);

        if (cancelled) return;

        if (savedIssuance) {
          setIssuanceFile(savedIssuance);

          setStorageStatus((prev) => ({
            ...prev,
            issuance: "Firebase se loaded",
          }));
        }

        if (savedErp) {
          setErpFile(savedErp);

          setStorageStatus((prev) => ({
            ...prev,
            erp: "Firebase se loaded",
          }));
        }
      } catch (err) {
        if (!cancelled) {
          setError(
            "Firebase se saved files load nahi ho sakin: " +
              err.message
          );
        }
      } finally {
        if (!cancelled) {
          setRestoringFiles(false);
        }
      }
    };

    restoreFiles();

    return () => {
      cancelled = true;
    };
  }, []);

  // File select hote hi Firebase mein save hogi
  const handleFileSelect = async (file, type) => {
    if (!file) return;

    if (type === "issuance") {
      setIssuanceFile(file);
    } else {
      setErpFile(file);
    }

    setResults([]);
    setHasReconciled(false);
    setError("");

    setSavingFiles((prev) => ({
      ...prev,
      [type]: true,
    }));

    setStorageStatus((prev) => ({
      ...prev,
      [type]: "Saving...",
    }));

    try {
      await saveReconciliationFile(file, type);

      setStorageStatus((prev) => ({
        ...prev,
        [type]: "Firebase mein saved",
      }));
    } catch (err) {
      setStorageStatus((prev) => ({
        ...prev,
        [type]: "Save failed",
      }));

      setError(
        "File Firebase mein save nahi ho saki: " +
          err.message
      );
    } finally {
      setSavingFiles((prev) => ({
        ...prev,
        [type]: false,
      }));
    }
  };

  const handleReconcile = async () => {
    if (!issuanceFile || !erpFile) {
      setError("Pehle dono Excel files select karein.");
      return;
    }

    if (
      savingFiles.issuance ||
      savingFiles.erp ||
      restoringFiles
    ) {
      setError(
        "Please wait, files Firebase ke sath sync ho rahi hain."
      );
      return;
    }

    setLoading(true);
    setError("");
    setResults([]);
    setHasReconciled(false);

    try {
      const issuanceData = await readExcel(
        issuanceFile,
        "issuance"
      );

      const erpData = await readExcel(erpFile, "erp");

      setIssuanceSheet(issuanceData.sheetName);
      setErpSheet(erpData.sheetName);

      setTotalIssue(issuanceData.totalQuantity);
      setTotalErp(erpData.totalQuantity);

      const comparison = compareData(
        issuanceData.records,
        erpData.records
      );

      setResults(comparison);
      setHasReconciled(true);
    } catch (err) {
      setError(
        err.message || "Excel files read karte waqt error aaya."
      );
    } finally {
      setLoading(false);
    }
  };

  const differences = results.filter((row) => !row.matched);

  const displayedResults = showOnlyDifferences
    ? differences
    : results;

  const matchedCount = results.length - differences.length;

  return (
    <div className="recon-page">
      <div className="recon-header">
        <div>
          <h1>Reconciliation</h1>
          <p>
            Issuance Sheet aur ERP Sub Inventory ka lot-wise comparison
          </p>
        </div>
      </div>

      <div className="recon-upload-grid">
        <div className="recon-upload-card">
          <label htmlFor="issuance-file">Issuance Sheet</label>

          <input
            id="issuance-file"
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) =>
              handleFileSelect(
                e.target.files?.[0],
                "issuance"
              )
            }
          />

          <span className="recon-file-name">
            {issuanceFile
              ? issuanceFile.name
              : "Excel file select karein"}
          </span>

          {storageStatus.issuance && (
            <small>{storageStatus.issuance}</small>
          )}
        </div>

        <div className="recon-upload-card">
          <label htmlFor="erp-file">ERP Sub Inventory</label>

          <input
            id="erp-file"
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) =>
              handleFileSelect(
                e.target.files?.[0],
                "erp"
              )
            }
          />

          <span className="recon-file-name">
            {erpFile
              ? erpFile.name
              : "Excel file select karein"}
          </span>

          {storageStatus.erp && (
            <small>{storageStatus.erp}</small>
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
          ? "Loading Saved Files..."
          : savingFiles.issuance || savingFiles.erp
          ? "Saving Files..."
          : loading
          ? "Checking..."
          : "Run Reconciliation"}
      </button>

      {error && <div className="recon-error">{error}</div>}

      {/* Totals: results hon ya na hon, totals show honge */}
      {hasReconciled && (
        <div className="recon-total-box">
          <div>
            <span>Total Issuance</span>
            <strong>{formatNumber(totalIssue)} Mtr</strong>
          </div>

          <div>
            <span>Total ERP / SAP</span>
            <strong>{formatNumber(totalErp)} Mtr</strong>
          </div>

          <div>
            <span>Net Difference</span>
            <strong>
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
              <strong className="recon-green">
                {matchedCount}
              </strong>
            </div>

            <div className="recon-summary-card">
              <span>Differences</span>
              <strong className="recon-red">
                {differences.length}
              </strong>
            </div>
          </div>

          <div className="recon-results-header">
            <h2>Comparison Report</h2>

            <label className="recon-filter">
              <input
                type="checkbox"
                checked={showOnlyDifferences}
                onChange={(e) =>
                  setShowOnlyDifferences(e.target.checked)
                }
              />
              Show differences only
            </label>
          </div>

          {/* Desktop Table */}
          <div className="recon-table-wrapper">
            <table className="recon-table">
              <thead>
                <tr>
                  <th>Lot Number</th>
                  <th>Issuance (Mtr)</th>
                  <th>ERP (Mtr)</th>
                  <th>Difference</th>
                  <th>Item Code Status</th>
                  <th>Status</th>
                </tr>
              </thead>

              <tbody>
                {displayedResults.map((row) => (
                  <tr key={row.lot}>
                    <td>{row.lot}</td>
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
                      {row.codeMismatch ? (
                        <div className="recon-code-detail">
                          <span className="recon-badge recon-badge-red">
                            Mismatch
                          </span>

                          {row.codeDifferences.map((code) => (
                            <div
                              className="recon-code-row"
                              key={code.code}
                            >
                              <strong>{code.code}</strong>
                              <span>
                                Issuance: {formatNumber(code.issuance)}
                              </span>
                              <span>
                                ERP: {formatNumber(code.erp)}
                              </span>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <span className="recon-badge recon-badge-green">
                          Matched
                        </span>
                      )}
                    </td>

                    <td>
                      {row.matched ? (
                        <span className="recon-badge recon-badge-green">
                          Matched
                        </span>
                      ) : (
                        <span className="recon-badge recon-badge-red">
                          Difference
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

                      {row.quantityMismatch && (
                        <div className="recon-status-detail">
                          Quantity mismatch
                        </div>
                      )}
                    </td>
                  </tr>
                ))}

                {displayedResults.length === 0 && (
                  <tr>
                    <td colSpan="6" className="recon-empty">
                      Koi difference nahi mila.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Mobile Cards */}
          <div className="recon-mobile-cards">
            {displayedResults.length === 0 ? (
              <div className="recon-mobile-empty">
                Koi difference nahi mila.
              </div>
            ) : (
              displayedResults.map((row) => (
                <div
                  className="recon-mobile-card"
                  key={row.lot}
                >
                  <div className="recon-mobile-card-header">
                    <div>
                      <span className="recon-mobile-label">
                        LOT NUMBER
                      </span>

                      <strong className="recon-mobile-lot">
                        {row.lot}
                      </strong>
                    </div>

                    <span
                      className={`recon-badge ${
                        row.matched
                          ? "recon-badge-green"
                          : "recon-badge-red"
                      }`}
                    >
                      {row.matched ? "Matched" : "Difference"}
                    </span>
                  </div>

                  <div className="recon-mobile-quantities">
                    <div>
                      <span>Issuance</span>
                      <strong>
                        {formatNumber(row.issueTotal)} Mtr
                      </strong>
                    </div>

                    <div>
                      <span>ERP / SAP</span>
                      <strong>
                        {formatNumber(row.erpTotal)} Mtr
                      </strong>
                    </div>
                  </div>

                  <div className="recon-mobile-difference">
                    <span>Quantity Difference</span>

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

                  <div className="recon-mobile-code">
                    <span className="recon-mobile-label">
                      ITEM CODE STATUS
                    </span>

                    {row.codeMismatch ? (
                      <>
                        <span className="recon-badge recon-badge-red">
                          Mismatch
                        </span>

                        {row.codeDifferences.map((code) => (
                          <div
                            className="recon-mobile-code-row"
                            key={code.code}
                          >
                            <strong>{code.code}</strong>
                            <span>
                              Issuance: {formatNumber(code.issuance)}
                            </span>
                            <span>
                              ERP: {formatNumber(code.erp)}
                            </span>
                          </div>
                        ))}
                      </>
                    ) : (
                      <span className="recon-badge recon-badge-green">
                        Matched
                      </span>
                    )}
                  </div>

                  {(row.missingInIssuance ||
                    row.missingInErp ||
                    row.quantityMismatch) && (
                    <div className="recon-mobile-status-details">
                      {row.missingInIssuance && (
                        <span>Missing in Issuance</span>
                      )}

                      {row.missingInErp && (
                        <span>Missing in ERP</span>
                      )}

                      {row.quantityMismatch && (
                        <span>Quantity mismatch</span>
                      )}
                    </div>
                  )}
                </div>
              ))
            )}
          </div>

          <div className="recon-sheet-info">
            Issuance Sheet: {issuanceSheet} | ERP Sheet: {erpSheet}
          </div>
        </>
      )}

      {hasReconciled && results.length === 0 && (
        <div className="recon-sheet-info">
          Totals calculate ho gaye hain, lekin lot-wise comparison
          ke liye complete Lot Number aur Item Code wali rows nahi milin.
          <br />
          Issuance Sheet: {issuanceSheet} | ERP Sheet: {erpSheet}
        </div>
      )}
    </div>
  );
}
