/**
 * worker/src/sheetsHelper.js
 * Google Sheets helper functions — mirrors src/sheetsHelper.js exactly,
 * but uses SheetsClient instead of googleapis.
 */

// Quote sheet name to handle spaces / special chars (same as original)
const q = (name) => `'${name.replace(/'/g, "''")}'`;

// ── rowsToObjects ─────────────────────────────────────────────

export function rowsToObjects(rows = []) {
  if (rows.length < 1) return [];
  const [headers, ...dataRows] = rows;
  return dataRows.map((row) =>
    headers.reduce((obj, key, i) => {
      if (key) obj[key] = row[i] ?? '';
      return obj;
    }, {})
  );
}

// ── getSheetsMeta ─────────────────────────────────────────────

export async function getSheetsMeta(client, spreadsheetId) {
  try {
    const data = await client.spreadsheetGet(spreadsheetId);
    return data.sheets.map((s) => ({
      sheetId: s.properties.sheetId,
      title:   s.properties.title,
    }));
  } catch (err) {
    throwMappedError(err, spreadsheetId);
  }
}

// ── findSheetId ───────────────────────────────────────────────

export async function findSheetId(client, spreadsheetId, sheetName) {
  const meta = await getSheetsMeta(client, spreadsheetId);
  return meta.find((s) => s.title === sheetName)?.sheetId ?? null;
}

// ── createSheet ───────────────────────────────────────────────

export async function createSheet(client, spreadsheetId, sheetName) {
  const data = await client.batchUpdate(spreadsheetId, {
    requests: [{ addSheet: { properties: { title: sheetName } } }],
  });
  return data.replies[0].addSheet.properties.sheetId;
}

// ── ensureSheet ───────────────────────────────────────────────

export async function ensureSheet(client, spreadsheetId, sheetName) {
  const existing = await findSheetId(client, spreadsheetId, sheetName);
  if (existing !== null) return { sheetId: existing, created: false };
  const sheetId = await createSheet(client, spreadsheetId, sheetName);
  return { sheetId, created: true };
}

// ── getHeaders ────────────────────────────────────────────────

export async function getHeaders(client, spreadsheetId, sheetName) {
  const data = await client.valuesGet(spreadsheetId, `${q(sheetName)}!1:1`);
  return data.values?.[0] ?? [];
}

// ── ensureHeaders ─────────────────────────────────────────────

export async function ensureHeaders(client, spreadsheetId, sheetName, keys) {
  const existing = await getHeaders(client, spreadsheetId, sheetName);
  if (existing.length > 0) return false;

  await client.valuesUpdate(spreadsheetId, `${q(sheetName)}!A1`, {
    values: [keys],
  });
  return true;
}

// ── readSheet ─────────────────────────────────────────────────

export async function readSheet(client, spreadsheetId, sheetName) {
  const data = await client.valuesGet(spreadsheetId, q(sheetName));
  const rows = data.values ?? [];
  return {
    headers: rows[0] ?? [],
    objects: rowsToObjects(rows),
  };
}

// ── Helper: Check if value exists in first column ──────────

async function isDuplicateInFirstColumn(client, spreadsheetId, sheetName, value) {
  try {
    const data = await client.valuesGet(spreadsheetId, `${q(sheetName)}!A:A`);
    const firstColumnValues = (data.values || []).flat();
    return firstColumnValues.includes(String(value));
  } catch (err) {
    // If sheet doesn't exist or error, assume no duplicates
    return false;
  }
}

// ── appendRow ─────────────────────────────────────────────────

export async function appendRow(client, spreadsheetId, sheetName, record) {
  await ensureSheet(client, spreadsheetId, sheetName);
  const keys = Object.keys(record);
  await ensureHeaders(client, spreadsheetId, sheetName, keys);

  const headers = await getHeaders(client, spreadsheetId, sheetName);
  
  // Check for duplicate in first column
  const firstColumnKey = headers[0];
  if (firstColumnKey && record[firstColumnKey]) {
    const isDuplicate = await isDuplicateInFirstColumn(
      client, spreadsheetId, sheetName, record[firstColumnKey]
    );
    if (isDuplicate) {
      const err = new Error(`Duplicate value "${record[firstColumnKey]}" found in first column "${firstColumnKey}".`);
      err.statusCode = 409; // Conflict
      throw err;
    }
  }

  const rowV = headers.map((h) => record[h] ?? '');

  const data = await client.valuesAppend(spreadsheetId, `${q(sheetName)}!A1`, {
    values: [rowV],
  });
  return data.updates;
}

// ── appendRowsBatch ───────────────────────────────────────────

export async function appendRowsBatch(client, spreadsheetId, sheetName, records) {
  if (!Array.isArray(records) || records.length === 0) return null;

  await ensureSheet(client, spreadsheetId, sheetName);
  const keys = Object.keys(records[0]);
  await ensureHeaders(client, spreadsheetId, sheetName, keys);

  const headers = await getHeaders(client, spreadsheetId, sheetName);
  const firstColumnKey = headers[0];

  // Get existing first column values for duplicate check
  let existingFirstColumnValues = [];
  if (firstColumnKey) {
    try {
      const data = await client.valuesGet(spreadsheetId, `${q(sheetName)}!A:A`);
      existingFirstColumnValues = (data.values || []).flat().map(String);
    } catch (err) {
      // If error, proceed without duplicate check
    }
  }

  // Filter out duplicates and build rows
  const validRecords = [];
  const skippedDuplicates = [];

  for (const record of records) {
    const firstColumnValue = String(record[firstColumnKey] || '');
    
    if (firstColumnKey && firstColumnValue && existingFirstColumnValues.includes(firstColumnValue)) {
      skippedDuplicates.push({
        value: firstColumnValue,
        record: record
      });
    } else {
      validRecords.push(record);
      // Add to existing values to prevent duplicates within the same batch
      if (firstColumnValue) existingFirstColumnValues.push(firstColumnValue);
    }
  }

  // If no valid records after filtering, return info about skipped items
  if (validRecords.length === 0) {
    return {
      updates: null,
      addedCount: 0,
      skippedCount: skippedDuplicates.length,
      skippedDuplicates: skippedDuplicates.map(d => d.value)
    };
  }

  const rows = validRecords.map((rec) => headers.map((h) => rec[h] ?? ''));

  const data = await client.valuesAppend(spreadsheetId, `${q(sheetName)}!A1`, {
    values: rows,
  });

  // Return detailed result
  return {
    updates: data.updates,
    addedCount: validRecords.length,
    skippedCount: skippedDuplicates.length,
    skippedDuplicates: skippedDuplicates.map(d => d.value)
  };
}

// ── updateRowsByObject ────────────────────────────────────────

// Helper function to convert column index to Excel column letter (A, B, ..., Z, AA, AB, ...)
function indexToColumnLetter(index) {
  let result = '';
  while (index >= 0) {
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
}

export async function updateRowsByObject(client, spreadsheetId, sheetName, updates) {
  try {
    const headers = await getHeaders(client, spreadsheetId, sheetName);
    
    // Build individual cell updates instead of full row updates
    const payload = [];
    
    for (const update of updates) {
      const { row, data } = update;
      
      // Only update cells for fields that are provided in data
      for (const [field, value] of Object.entries(data)) {
        const colIndex = headers.indexOf(field);
        if (colIndex === -1) continue; // Skip if field not found in headers
        
        // Convert column index to letter (A=0, B=1, ..., Z=25, AA=26, etc.)
        const colLetter = indexToColumnLetter(colIndex);
        
        payload.push({
          range: `${q(sheetName)}!${colLetter}${row}`,
          values: [[value ?? '']],
        });
      }
    }

    if (payload.length === 0) {
      return { totalUpdatedCells: 0 };
    }

    const data = await client.valuesBatchUpdate(spreadsheetId, {
      valueInputOption: 'USER_ENTERED',
      data: payload,
    });
    return data;
  } catch (err) {
    // If sheet doesn't exist, throw a clearer error
    if (err.statusCode === 404) {
      const e = new Error(`Sheet "${sheetName}" not found. Create it first or use append operations.`);
      e.statusCode = 404;
      throw e;
    }
    throw err;
  }
}

// ── deleteRows ────────────────────────────────────────────────

export async function deleteRows(client, spreadsheetId, sheetName, rowIndices) {
  const sheetId = await findSheetId(client, spreadsheetId, sheetName);
  if (sheetId === null) {
    const err = new Error(`Tab "${sheetName}" not found.`);
    err.statusCode = 404;
    throw err;
  }

  // Delete from bottom up so row indices don't shift
  const sorted   = [...new Set(rowIndices)].sort((a, b) => b - a);
  const requests = sorted.map((idx) => ({
    deleteDimension: {
      range: { sheetId, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 },
    },
  }));

  return client.batchUpdate(spreadsheetId, { requests });
}

// ── deleteRowsByKeys ─────────────────────────────────────────

export async function deleteRowsByKeys(client, spreadsheetId, sheetName, keys) {
  const sheetId = await findSheetId(client, spreadsheetId, sheetName);
  if (sheetId === null) {
    const err = new Error(`Tab "${sheetName}" not found.`);
    err.statusCode = 404;
    throw err;
  }

  // Get all data to find row indices for the given keys
  const data = await client.valuesGet(spreadsheetId, q(sheetName));
  const rows = data.values ?? [];
  
  if (rows.length <= 1) {
    // No data rows (only headers or empty sheet)
    return { deletedCount: 0, notFound: keys };
  }

  const [headers, ...dataRows] = rows;
  const firstColumnName = headers[0] || 'first column';
  
  if (!headers[0]) {
    const err = new Error('No headers found in sheet.');
    err.statusCode = 400;
    throw err;
  }

  // Find row indices for the given keys (convert to string for comparison)
  const keysToFind = keys.map(String);
  const rowIndicesToDelete = [];
  const foundKeys = [];
  const notFoundKeys = [];

  dataRows.forEach((row, index) => {
    const cellValue = String(row[0] || ''); // Always first column (index 0)
    if (keysToFind.includes(cellValue)) {
      // Row index in sheet = data row index + 2 (1 for 0-based, 1 for header)
      rowIndicesToDelete.push(index + 1); // +1 because we skip header row
      foundKeys.push(cellValue);
    }
  });

  // Check which keys were not found
  keysToFind.forEach(key => {
    if (!foundKeys.includes(key)) {
      notFoundKeys.push(key);
    }
  });

  if (rowIndicesToDelete.length === 0) {
    return { 
      deletedCount: 0, 
      notFound: notFoundKeys,
      searchColumn: firstColumnName
    };
  }

  // Delete rows (from bottom up so indices don't shift)
  const sorted = [...new Set(rowIndicesToDelete)].sort((a, b) => b - a);
  const requests = sorted.map((idx) => ({
    deleteDimension: {
      range: { sheetId, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 },
    },
  }));

  await client.batchUpdate(spreadsheetId, { requests });

  return {
    deletedCount: rowIndicesToDelete.length,
    deletedKeys: foundKeys,
    notFound: notFoundKeys,
    searchColumn: firstColumnName
  };
}

// ── throwMappedError ──────────────────────────────────────────

export function throwMappedError(err, spreadsheetId) {
  // SheetsClient sets err.statusCode from the HTTP response status
  const code = err?.statusCode ?? err?.code;
  if (code === 404) {
    const e = new Error(`Spreadsheet not found: ${spreadsheetId}`);
    e.statusCode = 404;
    throw e;
  }
  if (code === 403) {
    const e = new Error(
      `Permission denied. Check if the Service Account is shared as Editor on sheet ${spreadsheetId}`
    );
    e.statusCode = 403;
    throw e;
  }
  throw err;
}
