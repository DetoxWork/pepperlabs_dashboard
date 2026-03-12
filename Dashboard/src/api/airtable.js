/**
 * Airtable API client with paginated fetch helpers and retry/backoff support.
 */

const TOKEN = import.meta.env.VITE_AIRTABLE_TOKEN;
const BASE_ID = import.meta.env.VITE_AIRTABLE_BASE_ID;
const DEFAULT_TABLE_NAME = import.meta.env.VITE_AIRTABLE_TABLE_NAME;

const STATUS_FIELD_NAME = 'Registration Status';
const MIN_TABLE_FETCH_CONCURRENCY = 1;
const MAX_TABLE_FETCH_CONCURRENCY = 8;
const DEFAULT_TABLE_FETCH_CONCURRENCY = 6;
const MAX_RETRY_ATTEMPTS = 4;
const BASE_RETRY_DELAY_MS = 400;
const MAX_RETRY_DELAY_MS = 5000;
const RETRYABLE_STATUS_CODES = new Set([408, 425, 429, 500, 502, 503, 504]);

const TABLE_FETCH_CONCURRENCY = parseConcurrency(
    import.meta.env.VITE_AIRTABLE_CONCURRENCY,
    DEFAULT_TABLE_FETCH_CONCURRENCY
);

function parseConcurrency(rawValue, fallback) {
    const parsed = Number.parseInt(String(rawValue ?? ''), 10);
    if (Number.isNaN(parsed)) return fallback;
    return Math.min(MAX_TABLE_FETCH_CONCURRENCY, Math.max(MIN_TABLE_FETCH_CONCURRENCY, parsed));
}

function sleep(ms) {
    return new Promise((resolve) => {
        setTimeout(resolve, ms);
    });
}

function getRetryDelayMs(attempt, response) {
    const retryAfterHeader = response?.headers?.get?.('retry-after');
    if (retryAfterHeader) {
        const retryAfterSeconds = Number.parseInt(retryAfterHeader, 10);
        if (!Number.isNaN(retryAfterSeconds) && retryAfterSeconds >= 0) {
            return retryAfterSeconds * 1000;
        }
    }

    const exponentialDelay = BASE_RETRY_DELAY_MS * (2 ** attempt);
    const jitter = Math.floor(Math.random() * 220);
    return Math.min(MAX_RETRY_DELAY_MS, exponentialDelay + jitter);
}

function resolveBaseId(baseIdOverride) {
    return baseIdOverride || BASE_ID;
}

function ensureBaseConfig(baseIdOverride) {
    const baseId = resolveBaseId(baseIdOverride);
    if (!TOKEN || !baseId) {
        throw new Error(
            'Missing Airtable configuration. Please check your .env file for VITE_AIRTABLE_TOKEN and VITE_AIRTABLE_BASE_ID.'
        );
    }

    return baseId;
}

function ensureTableConfig(tableNameOrId, baseIdOverride) {
    ensureBaseConfig(baseIdOverride);
    if (!tableNameOrId) {
        throw new Error('Missing Airtable table name or ID.');
    }
}

function createAuthHeaders() {
    return {
        Authorization: `Bearer ${TOKEN}`,
    };
}

async function fetchWithRetry(url, contextLabel) {
    class NonRetryableRequestError extends Error {}

    let attempt = 0;
    let lastError = null;

    while (attempt <= MAX_RETRY_ATTEMPTS) {
        try {
            const response = await fetch(url, { headers: createAuthHeaders() });
            if (response.ok) {
                return response;
            }

            if (!RETRYABLE_STATUS_CODES.has(response.status) || attempt === MAX_RETRY_ATTEMPTS) {
                const errorBody = await response.text();
                throw new NonRetryableRequestError(`${contextLabel} (${response.status}): ${errorBody}`);
            }

            const retryDelay = getRetryDelayMs(attempt, response);
            await sleep(retryDelay);
            attempt++;
            continue;
        } catch (error) {
            if (error instanceof NonRetryableRequestError) {
                throw error;
            }

            lastError = error;
            if (attempt === MAX_RETRY_ATTEMPTS) {
                break;
            }

            const retryDelay = getRetryDelayMs(attempt);
            await sleep(retryDelay);
            attempt++;
        }
    }

    throw lastError || new Error(`${contextLabel} failed after retries.`);
}

function appendRequestedFields(url, fields) {
    if (!Array.isArray(fields) || fields.length === 0) return;

    fields.forEach((fieldName) => {
        if (!fieldName) return;
        url.searchParams.append('fields[]', fieldName);
    });
}

/**
 * Fetch ALL records from a specific Airtable table (handles pagination).
 * @param {string} [tableNameOrId] - The name or ID of the table to fetch from. Defaults to VITE_AIRTABLE_TABLE_NAME.
 * @param {{ fields?: string[], baseId?: string }} [options] - Optional query options.
 * @returns {Promise<Array>} - Array of { id, fields } objects.
 */
export async function fetchRecordsFromTable(tableNameOrId = DEFAULT_TABLE_NAME, options = {}) {
    const { fields = [], baseId: baseIdOverride } = options;
    const baseId = ensureBaseConfig(baseIdOverride);
    ensureTableConfig(tableNameOrId, baseId);

    const apiUrl = `https://api.airtable.com/v0/${baseId}/${encodeURIComponent(tableNameOrId)}`;
    const allRecords = [];
    let offset = null;

    do {
        const url = new URL(apiUrl);
        if (offset) url.searchParams.set('offset', offset);
        appendRequestedFields(url, fields);

        const response = await fetchWithRetry(url.toString(), `Airtable table request failed for "${tableNameOrId}"`);
        const data = await response.json();
        allRecords.push(...(data.records || []));
        offset = data.offset || null;
    } while (offset);

    return allRecords;
}

/**
 * Legacy wrapper for fetchRecordsFromTable using the default table name.
 */
export async function fetchAllRecords() {
    return fetchRecordsFromTable();
}

/**
 * Fetch metadata for all tables in the base.
 * @param {{ baseId?: string }} [options] - Optional base override.
 * @returns {Promise<Array>} - Array of table metadata objects.
 */
export async function fetchAllTables(options = {}) {
    const baseId = ensureBaseConfig(options.baseId);

    const url = `https://api.airtable.com/v0/meta/bases/${baseId}/tables`;
    const response = await fetchWithRetry(url, 'Airtable metadata request failed');
    const data = await response.json();
    return data.tables || [];
}

async function runWithConcurrencyLimit(items, limit, worker) {
    if (!Array.isArray(items) || items.length === 0) return;

    const workerCount = Math.max(1, Math.min(limit, items.length));
    let nextIndex = 0;

    async function runWorker() {
        while (nextIndex < items.length) {
            const currentIndex = nextIndex++;
            await worker(items[currentIndex], currentIndex);
        }
    }

    await Promise.all(Array.from({ length: workerCount }, () => runWorker()));
}

async function fetchStatusSummaryForTable(tableId, baseId) {
    const apiUrl = `https://api.airtable.com/v0/${baseId}/${encodeURIComponent(tableId)}`;
    const statusCounts = {};
    let offset = null;
    let total = 0;

    do {
        const url = new URL(apiUrl);
        if (offset) url.searchParams.set('offset', offset);
        url.searchParams.append('fields[]', STATUS_FIELD_NAME);

        const response = await fetchWithRetry(url.toString(), `Airtable status summary request failed for "${tableId}"`);
        const data = await response.json();
        const records = data.records || [];

        records.forEach((record) => {
            total++;
            const rawStatus = record?.fields?.[STATUS_FIELD_NAME];
            const statusKey = rawStatus === null || rawStatus === undefined ? '' : String(rawStatus).trim();
            statusCounts[statusKey] = (statusCounts[statusKey] || 0) + 1;
        });

        offset = data.offset || null;
    } while (offset);

    return { total, statusCounts };
}

/**
 * Fetch compact status summaries for every table in the base.
 * Only the Registration Status field is requested for dashboard loading speed.
 * @param {{ baseId?: string }} [options] - Optional base override.
 * @returns {Promise<Array>} - Array of { tableId, name, total, statusCounts }.
 */
export async function fetchStatusSummariesAcrossBase(options = {}) {
    const baseId = ensureBaseConfig(options.baseId);
    const tables = await fetchAllTables({ baseId });
    const summaries = new Array(tables.length);

    await runWithConcurrencyLimit(tables, TABLE_FETCH_CONCURRENCY, async (table, index) => {
        try {
            const compact = await fetchStatusSummaryForTable(table.id, baseId);
            summaries[index] = {
                tableId: table.id,
                name: table.name,
                total: compact.total,
                statusCounts: compact.statusCounts,
            };
        } catch (error) {
            console.error(`Failed to fetch status summary for table "${table.name}":`, error);
            summaries[index] = {
                tableId: table.id,
                name: table.name,
                total: 0,
                statusCounts: {},
            };
        }
    });

    return summaries.filter(Boolean);
}

/**
 * Fetch all records from every table in the base.
 * @param {{ baseId?: string }} [options] - Optional base override.
 * @returns {Promise<Object>} - A map of { tableName: records[] }.
 */
export async function fetchAllDataAcrossBase(options = {}) {
    const baseId = ensureBaseConfig(options.baseId);
    const tables = await fetchAllTables({ baseId });
    const results = {};

    await runWithConcurrencyLimit(tables, TABLE_FETCH_CONCURRENCY, async (table) => {
        try {
            const records = await fetchRecordsFromTable(table.id, { baseId });
            results[table.name] = records;
        } catch (error) {
            console.error(`Failed to fetch records for table "${table.name}":`, error);
            results[table.name] = [];
        }
    });

    return results;
}
