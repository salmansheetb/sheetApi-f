/**
 * worker/src/routes.js
 * API route handlers — mirrors src/routes.js exactly.
 */

import { getSheetsClient } from './sheetsClient.js';
import {
  getSheetsMeta,
  readSheet,
  appendRow,
  appendRowsBatch,
  updateRowsByObject,
  deleteRows,
  deleteRowsByKeys,
} from './sheetsHelper.js';
import { Router, jsonResponse, errorResponse } from './router.js';

// ── helpers ───────────────────────────────────────────────────

/** Parse JSON body safely — works for GET/POST/PUT/DELETE */
async function parseBody(request) {
  const text = await request.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    const err = new Error('Invalid JSON body.');
    err.statusCode = 400;
    throw err;
  }
}

function handleError(reply, err) {
  const statusCode = err.statusCode || 500;
  return jsonResponse(
    { success: false, error: err.message || 'Internal Server Error' },
    statusCode
  );
}

// ── router ────────────────────────────────────────────────────

export function buildRouter() {
  const router = new Router();

  // 1. Read All Tabs
  router.get('/read/:spreadsheetId', async (request, { params, env }) => {
    try {
      const { spreadsheetId } = params;
      const client = getSheetsClient(env);
      const meta   = await getSheetsMeta(client, spreadsheetId);

      const tabs = await Promise.all(
        meta.map(async ({ title }) => {
          const { objects, headers } = await readSheet(client, spreadsheetId, title);
          return { tab: title, headers, rows: objects };
        })
      );

      return jsonResponse({ success: true, spreadsheetId, tabs });
    } catch (err) {
      return handleError(null, err);
    }
  });

  // 2. Append (Single Row)
  router.post('/append', async (request, { env }) => {
    try {
      const { spreadsheetId, sheetName, data } = await parseBody(request);
      if (!spreadsheetId || !sheetName || !data)
        throw Object.assign(new Error('Missing spreadsheetId, sheetName or data in body.'), { statusCode: 400 });

      const result = await appendRow(getSheetsClient(env), spreadsheetId, sheetName, data);
      return jsonResponse({ success: true, updates: result }, 201);
    } catch (err) {
      return handleError(null, err);
    }
  });

  // 3. Append Batch (Multiple Rows)
  router.post('/append-batch', async (request, { env }) => {
    try {
      const { spreadsheetId, sheetName, data } = await parseBody(request);
      if (!Array.isArray(data))
        throw Object.assign(new Error('data must be an array of objects.'), { statusCode: 400 });

      const result = await appendRowsBatch(getSheetsClient(env), spreadsheetId, sheetName, data);
      
      // Handle the new detailed response format
      if (result && typeof result === 'object' && 'addedCount' in result) {
        return jsonResponse({
          success: true,
          updates: result.updates,
          addedCount: result.addedCount,
          skippedCount: result.skippedCount,
          skippedDuplicates: result.skippedDuplicates
        }, 201);
      } else {
        // Fallback for old format (null or simple updates object)
        return jsonResponse({ success: true, updates: result }, 201);
      }
    } catch (err) {
      return handleError(null, err);
    }
  });

  // 4. Update (Batch - Row based)
  router.put('/update-batch', async (request, { env }) => {
    try {
      const { spreadsheetId, sheetName, updates } = await parseBody(request);
      if (!sheetName)
        throw Object.assign(new Error('sheetName is required for updates.'), { statusCode: 400 });
      if (!Array.isArray(updates))
        throw Object.assign(new Error('updates must be an array.'), { statusCode: 400 });

      const result = await updateRowsByObject(getSheetsClient(env), spreadsheetId, sheetName, updates);
      return jsonResponse({ success: true, totalUpdatedCells: result.totalUpdatedCells || 0 });
    } catch (err) {
      return handleError(null, err);
    }
  });

  // 5. Delete (Batch - Index based)
  router.delete('/delete-batch', async (request, { env }) => {
    try {
      const body = await parseBody(request);
      const { spreadsheetId, sheetName } = body;
      const rowIndices = body.rowIndices || body.rows;
      
      if (!spreadsheetId || !sheetName)
        throw Object.assign(new Error('Missing spreadsheetId or sheetName.'), { statusCode: 400 });
      if (!rowIndices || !Array.isArray(rowIndices))
        throw Object.assign(new Error('rowIndices (or rows) must be an array.'), { statusCode: 400 });

      await deleteRows(getSheetsClient(env), spreadsheetId, sheetName, rowIndices);
      return jsonResponse({ success: true, deletedRows: rowIndices.length });
    } catch (err) {
      return handleError(null, err);
    }
  });

  // 6. Delete by Keys (First Column Values)
  router.delete('/delete-by-keys', async (request, { env }) => {
    try {
      const { spreadsheetId, sheetName, keys } = await parseBody(request);
      
      if (!spreadsheetId || !sheetName)
        throw Object.assign(new Error('Missing spreadsheetId or sheetName.'), { statusCode: 400 });
      if (!keys || !Array.isArray(keys))
        throw Object.assign(new Error('keys must be an array.'), { statusCode: 400 });
      if (keys.length === 0)
        throw Object.assign(new Error('keys array cannot be empty.'), { statusCode: 400 });

      const result = await deleteRowsByKeys(getSheetsClient(env), spreadsheetId, sheetName, keys);
      return jsonResponse({
        success: true,
        deletedCount: result.deletedCount,
        deletedKeys: result.deletedKeys,
        notFound: result.notFound,
        searchColumn: result.searchColumn
      });
    } catch (err) {
      return handleError(null, err);
    }
  });

  // Health check
  router.get('/health', () =>
    jsonResponse({ status: 'online', time: new Date().toISOString() })
  );

  return router;
}
