import './style.css';
import { fetchStatusSummariesAcrossBase, fetchRecordsFromTable } from './api/airtable.js';
import { fetchAvpnSummary } from './api/thinkific.js';
import { Chart, DoughnutController, ArcElement, Tooltip, Legend } from 'chart.js';

const doughnutCenterTextPlugin = {
    id: 'doughnutCenterText',
    afterDraw(chart) {
        if (chart.config.type !== 'doughnut') return;

        const dataset = chart.data.datasets?.[0];
        if (!dataset || !Array.isArray(dataset.data)) return;

        const total = dataset.data.reduce((acc, n) => acc + (Number(n) || 0), 0);
        if (total <= 0) return;

        const { ctx, chartArea } = chart;
        if (!chartArea) return;

        const x = (chartArea.left + chartArea.right) / 2;
        const y = (chartArea.top + chartArea.bottom) / 2;

        ctx.save();
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        ctx.fillStyle = '#8d7a66';
        ctx.font = "500 11px 'Space Grotesk', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
        ctx.fillText('TOTAL', x, y - 13);

        ctx.fillStyle = '#4a4137';
        ctx.font = "500 21px 'Space Grotesk', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
        ctx.fillText(total.toLocaleString(), x, y + 11);

        ctx.restore();
    }
};

Chart.register(DoughnutController, ArcElement, Tooltip, Legend, doughnutCenterTextPlugin);
Chart.defaults.color = '#7b7165';

/* ===== Constants ===== */
const TRAINING_STATUSES = [
    'Approved',
    'Accepted',
    'Incomplete',
    'Signed',
    'Review',
    'Pending',
    'Rejected',
    'Withdraw',
    'No Application Found',
    'Candidate Does not Exist',
    'No Data',
    'Deleted',
];

const TABLE_STATUS_COLUMNS = [
    'Accepted',
    'Incomplete',
    'Signed',
    'Review',
    'Pending',
    'Rejected',
];

const ACTIONABLE_STATUSES = ['Accepted', 'Incomplete', 'Review', 'Pending', 'Rejected'];
const WORKFLOW_STATUSES = ['Accepted', 'Incomplete', 'Review', 'Pending', 'Rejected', 'Signed', 'Withdraw', 'Approved'];
const DATA_QUALITY_STATUSES = ['No Application Found', 'No Data', 'Candidate Does not Exist', 'Deleted'];
const STATUS_GROUPS = [
    { key: 'workflow', label: 'Application Workflow', statuses: WORKFLOW_STATUSES },
    { key: 'data-quality', label: 'Data Quality', statuses: DATA_QUALITY_STATUSES },
];
const WORKFLOW_STATUS_SET = new Set(WORKFLOW_STATUSES);
const STATUS_ORDER = new Map(
    [...WORKFLOW_STATUSES, ...DATA_QUALITY_STATUSES].map((status, index) => [status, index])
);
const CHART_MAX_SEGMENTS = 6;
const CHART_OTHER_LABEL = 'Other';
const CHART_OTHER_COLOR = '#b4ad9f';

const STATUS_COLORS = {
    Approved: '#8f7350',
    Accepted: '#e97824',
    Incomplete: '#b88a2f',
    Signed: '#bf7b52',
    Review: '#d95a3a',
    Pending: '#c86f1e',
    Rejected: '#c5513a',
    Withdraw: '#97733f',
    'No Application Found': '#9a9488',
    'Candidate Does not Exist': '#8b95a5',
    'No Data': '#a3abae',
    Deleted: '#7d7468',
};

const NAME_COLLATOR = new Intl.Collator(undefined, { sensitivity: 'base', numeric: true });
const SEARCH_DEBOUNCE_MS = 120;
const APPLICANT_QUERY_PARAM = 'applicant';
const CONTEXT_QUERY_PARAM = 'context';

const DASHBOARD_CONTEXTS = [
    {
        key: 'yayasan-pepper-labs',
        label: 'Yayasan PenerajuX Pepper Labs',
        baseId: import.meta.env.VITE_AIRTABLE_BASE_ID_YAYASAN_PEPPER_LABS || import.meta.env.VITE_AIRTABLE_BASE_ID,
    },
    {
        key: 'avpn',
        label: 'AVPN',
        baseId: import.meta.env.VITE_AIRTABLE_BASE_ID_AVPN || import.meta.env.VITE_AIRTABLE_BASE_ID,
    },
];

const DASHBOARD_CONTEXTS_BY_KEY = new Map(
    DASHBOARD_CONTEXTS.map((context) => [context.key, context])
);
const DEFAULT_DASHBOARD_CONTEXT_KEY = DASHBOARD_CONTEXTS[0]?.key || 'yayasan-pepper-labs';
const AVPN_CONTEXT_KEY = 'avpn';
const AVPN_CATEGORY_ORDER = ['Students', 'Educators', 'SMEs'];

const AVPN_TARGETS = {
    Students: parseNumericEnv(import.meta.env.VITE_AVPN_TARGET_STUDENTS, 25000),
    Educators: parseNumericEnv(import.meta.env.VITE_AVPN_TARGET_EDUCATORS, 12500),
    SMEs: parseNumericEnv(import.meta.env.VITE_AVPN_TARGET_SMES, 12500),
};

const AVPN_BASELINES = {
    Students: parseNumericEnv(import.meta.env.VITE_AVPN_BASELINE_STUDENTS, 0),
    Educators: parseNumericEnv(import.meta.env.VITE_AVPN_BASELINE_EDUCATORS, 0),
    SMEs: parseNumericEnv(import.meta.env.VITE_AVPN_BASELINE_SMES, 0),
};

function parseNumericEnv(rawValue, fallbackValue) {
    const parsed = Number(rawValue);
    return Number.isFinite(parsed) ? parsed : fallbackValue;
}

/* ===== DOM Elements ===== */
const $tableBody = document.getElementById('table-body');
const $tableFooter = document.getElementById('table-footer');
const $loadingOverlay = document.getElementById('loading-overlay');
const $errorToast = document.getElementById('error-toast');
const $errorMessage = document.getElementById('error-message');
const $tableTotalCount = document.getElementById('table-total-count');
const $lastUpdated = document.getElementById('last-updated');
const $btnRefresh = document.getElementById('btn-refresh');
const $tableSearch = document.getElementById('table-search');
const $statusFilter = document.getElementById('status-filter');
const $tableHead = document.querySelector('#summary-table thead');
const $btnPrev = document.getElementById('btn-prev');
const $btnNext = document.getElementById('btn-next');
const $currentPageLabel = document.getElementById('current-page');
const $totalPagesLabel = document.getElementById('total-pages');
const $summarySection = document.getElementById('summary-section');
const $tableSection = document.getElementById('table-section');
const $acceptedTableSection = document.getElementById('accepted-table-section');
const $acceptedTableCount = document.getElementById('accepted-table-count');
const $acceptedCount = document.getElementById('accepted-count');
const $acceptedInvoiceList = document.getElementById('accepted-invoice-list');
const $applicantPage = document.getElementById('applicant-page');
const $kpiStrip = document.getElementById('kpi-strip');
const $kpiRegistered = document.getElementById('kpi-registered');
const $kpiProgressBand = document.getElementById('kpi-progress-band');
const $kpiCompleted = document.getElementById('kpi-completed');
const $kpiPreSurvey = document.getElementById('kpi-presurvey');
const $kpiLabel1 = document.getElementById('kpi-label-1');
const $kpiLabel2 = document.getElementById('kpi-label-2');
const $kpiLabel3 = document.getElementById('kpi-label-3');
const $kpiLabel4 = document.getElementById('kpi-label-4');
const $kpiMeta1 = document.getElementById('kpi-meta-1');
const $kpiMeta2 = document.getElementById('kpi-meta-2');
const $kpiMeta3 = document.getElementById('kpi-meta-3');
const $kpiMeta4 = document.getElementById('kpi-meta-4');
const $avpnSection = document.getElementById('avpn-section');
const $avpnSummaryBody = document.getElementById('avpn-summary-body');
const $avpnSummaryFooter = document.getElementById('avpn-summary-footer');
const $avpnTargetBody = document.getElementById('avpn-target-body');
const $avpnGeneratedAt = document.getElementById('avpn-generated-at');
const $dashboardContextLabel = document.getElementById('dashboard-context-label');
const $dashboardContextLinks = Array.from(document.querySelectorAll('.sidebar-link[data-dashboard-key]'));
const $analyticsKpis = document.getElementById('analytics-kpis');
const $chartActionableOnly = document.getElementById('chart-actionable-only');
const $chartResetFilters = document.getElementById('chart-reset-filters');
const $statusLegendGroups = document.getElementById('status-legend-groups');

/* ===== State ===== */
let tableSummaries = [];
let acceptedApplicantSummaries = [];
const applicantRecordsCache = new Map();
let chartInstance = null;
let currentPage = 1;
const rowsPerPage = 10;
let searchQuery = '';
let statusFilterValue = 'all';
let sortState = { key: 'name', dir: 'asc' };
let activeApplicant = '';
let isLoadingData = false;
let searchDebounceId = null;
let tableViewCache = { key: '', rows: [] };
let tableDataVersion = 0;
let errorHideTimerId = null;
let activeDashboardContextKey = DEFAULT_DASHBOARD_CONTEXT_KEY;
let avpnSummaryRows = [];
let latestChartTotals = null;
let chartActionableOnly = false;
let chartIsolatedStatus = '';
let legendClickTimeoutId = null;
const chartStatusVisibility = new Map(TRAINING_STATUSES.map((status) => [status, true]));

/* ===== Initialise ===== */
async function init() {
    if (
        !$tableBody ||
        !$tableFooter ||
        !$loadingOverlay ||
        !$errorToast ||
        !$errorMessage ||
        !$tableTotalCount ||
        !$lastUpdated ||
        !$btnRefresh ||
        !$tableSearch ||
        !$statusFilter ||
        !$tableHead ||
        !$btnPrev ||
        !$btnNext ||
        !$currentPageLabel ||
        !$totalPagesLabel ||
        !$summarySection ||
        !$tableSection ||
        !$acceptedTableSection ||
        !$acceptedTableCount ||
        !$acceptedCount ||
        !$acceptedInvoiceList ||
        !$applicantPage ||
        !$kpiStrip ||
        !$kpiRegistered ||
        !$kpiProgressBand ||
        !$kpiCompleted ||
        !$kpiPreSurvey ||
        !$kpiLabel1 ||
        !$kpiLabel2 ||
        !$kpiLabel3 ||
        !$kpiLabel4 ||
        !$kpiMeta1 ||
        !$kpiMeta2 ||
        !$kpiMeta3 ||
        !$kpiMeta4 ||
        !$avpnSection ||
        !$avpnSummaryBody ||
        !$avpnSummaryFooter ||
        !$avpnTargetBody ||
        !$avpnGeneratedAt ||
        !$dashboardContextLabel ||
        !$analyticsKpis ||
        !$chartActionableOnly ||
        !$chartResetFilters ||
        !$statusLegendGroups ||
        $dashboardContextLinks.length === 0
    ) {
        console.error('Dashboard initialization failed: missing required DOM elements.');
        return;
    }

    activeDashboardContextKey = getDashboardContextFromRoute();
    syncDashboardContextUI();
    const isAvpnContext = getActiveDashboardContext()?.key === AVPN_CONTEXT_KEY;
    if (isAvpnContext && getApplicantFromRoute()) {
        clearApplicantRoute();
    }
    activeApplicant = isAvpnContext ? '' : getApplicantFromRoute();
    togglePageSections(isAvpnContext ? 'avpn' : activeApplicant ? 'applicant' : 'dashboard');

    populateStatusFilterOptions();
    bindSortingHandlers();
    updateSortIndicators();

    try {
        await loadData();
    } catch (err) {
        showError(err.message);
        hideLoading();
    }

    $btnRefresh.addEventListener('click', handleRefresh);
    $btnPrev.addEventListener('click', () => changePage(-1));
    $btnNext.addEventListener('click', () => changePage(1));
    $tableSearch.addEventListener('input', handleSearchInput);
    $tableBody.addEventListener('click', handleApplicantLinkClick);
    $applicantPage.addEventListener('click', handleApplicantBackClick);
    $dashboardContextLinks.forEach((link) => {
        link.addEventListener('click', handleDashboardContextClick);
    });
    $chartActionableOnly.addEventListener('change', handleChartActionableToggle);
    $chartResetFilters.addEventListener('click', handleChartResetFilters);
    $statusLegendGroups.addEventListener('click', handleStatusLegendClick);
    $statusLegendGroups.addEventListener('dblclick', handleStatusLegendDoubleClick);
    window.addEventListener('popstate', handlePopState);
    $statusFilter.addEventListener('change', () => {
        statusFilterValue = $statusFilter.value;
        currentPage = 1;
        invalidateTableViewCache();
        renderTable();
    });
}

function getActiveDashboardContext() {
    const activeContext = DASHBOARD_CONTEXTS_BY_KEY.get(activeDashboardContextKey);
    if (activeContext) return activeContext;

    return DASHBOARD_CONTEXTS_BY_KEY.get(DEFAULT_DASHBOARD_CONTEXT_KEY) || DASHBOARD_CONTEXTS[0];
}

function getDashboardContextFromRoute() {
    const params = new URLSearchParams(window.location.search);
    const contextKey = params.get(CONTEXT_QUERY_PARAM);
    if (contextKey && DASHBOARD_CONTEXTS_BY_KEY.has(contextKey)) {
        return contextKey;
    }

    return DEFAULT_DASHBOARD_CONTEXT_KEY;
}

function syncDashboardContextUI() {
    const activeContext = getActiveDashboardContext();
    if (!activeContext) return;

    $dashboardContextLabel.textContent = activeContext.label;
    $dashboardContextLinks.forEach((link) => {
        const isActive = link.dataset.dashboardKey === activeContext.key;
        link.classList.toggle('is-active', isActive);
        link.setAttribute('aria-pressed', String(isActive));
    });
}

async function handleDashboardContextClick(event) {
    const contextKey = event.currentTarget?.dataset?.dashboardKey;
    if (!contextKey || !DASHBOARD_CONTEXTS_BY_KEY.has(contextKey) || isLoadingData) {
        return;
    }

    if (contextKey === activeDashboardContextKey) {
        if (!activeApplicant) return;

        activeApplicant = '';
        window.history.pushState({}, '', getDashboardHref());
        togglePageSections('dashboard');
        renderTable();
        chartInstance?.resize();
        return;
    }

    activeDashboardContextKey = contextKey;
    activeApplicant = '';
    syncDashboardContextUI();
    window.history.pushState({}, '', getDashboardHref());
    togglePageSections(contextKey === AVPN_CONTEXT_KEY ? 'avpn' : 'dashboard');
    await loadData();
}

function formatKpiValue(value) {
    const numericValue = Number(value);
    if (!Number.isFinite(numericValue)) return '--';
    return numericValue.toLocaleString();
}

function formatLegendPercent(value, total) {
    if (!Number.isFinite(total) || total <= 0) return '0%';
    const pct = (value / total) * 100;
    if (pct > 0 && pct < 1) return '<1%';
    return `${pct.toFixed(1)}%`;
}

function updateKpiCardsForStandardDashboard(totals) {
    $kpiLabel1.textContent = 'Total Applicants';
    $kpiLabel2.textContent = 'Incomplete';
    $kpiLabel3.textContent = 'Accepted';
    $kpiLabel4.textContent = 'Pending';

    $kpiMeta1.textContent = 'Across all applicant tables';
    $kpiMeta2.textContent = 'Needs follow-up';
    $kpiMeta3.textContent = 'Confirmed enrollments';
    $kpiMeta4.textContent = 'Awaiting resolution';

    $kpiRegistered.textContent = formatKpiValue(totals?.totalRecords ?? 0);
    $kpiProgressBand.textContent = formatKpiValue(totals?.Incomplete ?? 0);
    $kpiCompleted.textContent = formatKpiValue(totals?.Accepted ?? 0);
    $kpiPreSurvey.textContent = formatKpiValue(totals?.Pending ?? 0);
}

function updateKpiCardsForAvpn(rows) {
    const totals = (rows || []).reduce((acc, row) => ({
        registeredInLms: acc.registeredInLms + (Number(row.registeredInLms) || 0),
        completed50: acc.completed50 + (Number(row.completed50) || 0),
        completed100: acc.completed100 + (Number(row.completed100) || 0),
        completedPreSurvey: acc.completedPreSurvey + (Number(row.completedPreSurvey) || 0),
    }), { registeredInLms: 0, completed50: 0, completed100: 0, completedPreSurvey: 0 });

    $kpiLabel1.textContent = 'Total Registered';
    $kpiLabel2.textContent = 'Completed 40%-59%';
    $kpiLabel3.textContent = 'Completed 100%';
    $kpiLabel4.textContent = 'Completed Pre-Survey';

    $kpiMeta1.textContent = 'Across 6 AVPN programs';
    $kpiMeta2.textContent = 'Mid-funnel learners';
    $kpiMeta3.textContent = 'Graduated learners';
    $kpiMeta4.textContent = 'Survey completion';

    $kpiRegistered.textContent = formatKpiValue(totals.registeredInLms);
    $kpiProgressBand.textContent = formatKpiValue(totals.completed50);
    $kpiCompleted.textContent = formatKpiValue(totals.completed100);
    $kpiPreSurvey.textContent = formatKpiValue(totals.completedPreSurvey);
}

function renderAnalyticsKpis(totals) {
    const chips = ACTIONABLE_STATUSES.map((status) => {
        const count = Number(totals?.[status]) || 0;
        const color = STATUS_COLORS[status] || '#9a7743';
        return `
            <article class="analytics-kpi-chip" style="--chip-color:${color}">
                <p class="analytics-kpi-chip-label">${escapeHtml(status)}</p>
                <p class="analytics-kpi-chip-value">${count.toLocaleString()}</p>
            </article>
        `;
    }).join('');

    $analyticsKpis.innerHTML = chips;
}

function resetChartVisibility() {
    TRAINING_STATUSES.forEach((status) => {
        chartStatusVisibility.set(status, true);
    });
}

function getChartEligibleStatuses() {
    const baseStatuses = chartActionableOnly ? ACTIONABLE_STATUSES : TRAINING_STATUSES;
    if (chartIsolatedStatus && baseStatuses.includes(chartIsolatedStatus)) {
        return [chartIsolatedStatus];
    }
    return baseStatuses;
}

function ensureChartHasVisibleStatus() {
    const eligibleStatuses = getChartEligibleStatuses();
    const hasVisibleStatus = eligibleStatuses.some((status) => chartStatusVisibility.get(status) !== false);
    if (!hasVisibleStatus && eligibleStatuses.length > 0) {
        chartStatusVisibility.set(eligibleStatuses[0], true);
    }
}

function rerenderStatusAnalytics() {
    if (!latestChartTotals) return;
    renderChart(latestChartTotals);
}

function handleChartActionableToggle() {
    chartActionableOnly = Boolean($chartActionableOnly.checked);
    chartIsolatedStatus = '';
    ensureChartHasVisibleStatus();
    rerenderStatusAnalytics();
}

function handleChartResetFilters() {
    chartActionableOnly = false;
    chartIsolatedStatus = '';
    $chartActionableOnly.checked = false;
    resetChartVisibility();
    rerenderStatusAnalytics();
}

function getLegendItemStatusFromEvent(event) {
    const item = event.target.closest('[data-status]');
    if (!item) return '';
    return String(item.dataset.status || '').trim();
}

function handleStatusLegendClick(event) {
    const status = getLegendItemStatusFromEvent(event);
    if (!status) return;

    if (legendClickTimeoutId !== null) {
        window.clearTimeout(legendClickTimeoutId);
    }

    legendClickTimeoutId = window.setTimeout(() => {
        legendClickTimeoutId = null;
        if (!TRAINING_STATUSES.includes(status)) return;

        chartIsolatedStatus = '';
        const currentlyVisible = chartStatusVisibility.get(status) !== false;
        chartStatusVisibility.set(status, !currentlyVisible);
        ensureChartHasVisibleStatus();
        rerenderStatusAnalytics();
    }, 180);
}

function handleStatusLegendDoubleClick(event) {
    const status = getLegendItemStatusFromEvent(event);
    if (!status || !TRAINING_STATUSES.includes(status)) return;

    event.preventDefault();
    if (legendClickTimeoutId !== null) {
        window.clearTimeout(legendClickTimeoutId);
        legendClickTimeoutId = null;
    }

    const eligibleStatuses = chartActionableOnly ? ACTIONABLE_STATUSES : TRAINING_STATUSES;
    if (!eligibleStatuses.includes(status)) return;

    if (chartIsolatedStatus === status) {
        chartIsolatedStatus = '';
        resetChartVisibility();
    } else {
        chartIsolatedStatus = status;
        TRAINING_STATUSES.forEach((entry) => {
            chartStatusVisibility.set(entry, entry === status);
        });
    }

    rerenderStatusAnalytics();
}

/* ===== Data Loading ===== */
async function loadData(options = {}) {
    if (isLoadingData) return;
    const { forceAvpnRefresh = false } = options;

    isLoadingData = true;
    showLoading();
    $btnRefresh.disabled = true;

    try {
        const activeContext = getActiveDashboardContext();
        if (activeContext?.key === AVPN_CONTEXT_KEY) {
            const avpnSummary = await fetchAvpnSummary({ forceRefresh: forceAvpnRefresh });
            avpnSummaryRows = normalizeAvpnSummaryRows(avpnSummary.rows);
            tableSummaries = [];
            acceptedApplicantSummaries = [];
            latestChartTotals = null;
            $analyticsKpis.innerHTML = '';
            $statusLegendGroups.innerHTML = '';
            tableDataVersion++;
            invalidateTableViewCache();
            updateKpiCardsForAvpn(avpnSummaryRows);
            renderAvpnSummaryTables(avpnSummaryRows);
            updateAvpnGeneratedAt(avpnSummary.generatedAt);
            togglePageSections('avpn');
            updateLastUpdated();
            return;
        }

        const statusSummaries = await fetchStatusSummariesAcrossBase({ baseId: activeContext?.baseId });
        const processedSummaries = processStatusSummaries(statusSummaries);
        acceptedApplicantSummaries = processedSummaries.filter((row) => isAcceptedAggregateApplicant(row.name));
        tableSummaries = processedSummaries.filter((row) => !isAcceptedAggregateApplicant(row.name));
        applicantRecordsCache.clear();
        avpnSummaryRows = [];
        tableDataVersion++;
        currentPage = 1;
        invalidateTableViewCache();
        const totals = calculateGrandTotals(tableSummaries);
        updateKpiCardsForStandardDashboard(totals);
        await renderAcceptedApplicantsCard();

        if (activeApplicant) {
            const rendered = await renderApplicantPage(activeApplicant);
            if (!rendered) {
                showError(`Applicant "${activeApplicant}" was not found. Showing dashboard list.`);
                activeApplicant = '';
                clearApplicantRoute();
                togglePageSections('dashboard');
                renderChart(totals);
                renderTable();
            }
        } else {
            togglePageSections('dashboard');
            renderChart(totals);
            renderTable();
        }

        updateLastUpdated();
    } catch (err) {
        showError(err.message);
    } finally {
        hideLoading();
        $btnRefresh.disabled = false;
        isLoadingData = false;
    }
}

async function handleRefresh() {
    if (isLoadingData) return;

    $btnRefresh.classList.add('spinning');
    try {
        await loadData({ forceAvpnRefresh: true });
    } catch (err) {
        showError(err.message);
    } finally {
        $btnRefresh.classList.remove('spinning');
    }
}

/* ===== Data Processing ===== */
function normalizeStatus(rawStatus) {
    if (!rawStatus) return 'No Data';

    const s = rawStatus.trim();
    const lower = s.toLowerCase();

    if (lower.includes('approved')) return 'Approved';
    if (lower.includes('accepted')) return 'Accepted';
    if (lower.includes('incomplete')) return 'Incomplete';
    if (lower.includes('signed')) return 'Signed';
    if (lower.includes('review')) return 'Review';
    if (lower.includes('deleted')) return 'Deleted';

    if (lower.includes('pending')) return 'Pending';
    if (lower.includes('rejected')) return 'Rejected';
    if (lower.includes('withdraw')) return 'Withdraw';

    if (lower.includes('no application found') || lower.includes('no account') || lower.includes('not registered yet') || lower === 'no application') {
        return 'No Application Found';
    }

    if (lower.includes('candidate does not exist') || lower.includes('candidates does not exist') || lower.includes('candidates not exist') || lower.includes('candidate not exist') || lower.includes('candidates do not exist')) {
        return 'Candidate Does not Exist';
    }

    if (lower === '[empty]' || lower === '') return 'No Data';

    return 'No Data';
}

function isAcceptedAggregateApplicant(name) {
    const normalized = String(name || '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();

    return normalized === 'accepted applicants' || normalized.startsWith('accepted applicants ');
}

function processStatusSummaries(summaries) {
    return summaries.map((summary) => {
        const counts = {};
        TRAINING_STATUSES.forEach((status) => {
            counts[status] = 0;
        });

        Object.entries(summary.statusCounts || {}).forEach(([rawStatus, count]) => {
            const status = normalizeStatus(rawStatus);
            if (counts[status] !== undefined) {
                counts[status] += Number(count) || 0;
            }
        });

        return {
            name: summary.name,
            tableId: summary.tableId,
            counts,
            total: Number(summary.total) || 0
        };
    }).sort((a, b) => NAME_COLLATOR.compare(a.name, b.name));
}

function calculateGrandTotals(summaries) {
    const totals = { totalRecords: 0 };
    TRAINING_STATUSES.forEach((status) => {
        totals[status] = 0;
    });

    summaries.forEach((summary) => {
        totals.totalRecords += summary.total;
        TRAINING_STATUSES.forEach((status) => {
            totals[status] += summary.counts[status];
        });
    });

    return totals;
}

function populateStatusFilterOptions() {
    $statusFilter.innerHTML = [
        '<option value="all">All statuses</option>',
        ...TABLE_STATUS_COLUMNS.map((status) => `<option value="${status}">${status}</option>`)
    ].join('');
}

function bindSortingHandlers() {
    $tableHead.querySelectorAll('th[data-sort-key]').forEach((header) => {
        header.addEventListener('click', () => {
            const key = header.dataset.sortKey;
            if (!key) return;
            handleSort(key);
        });
    });
}

function handleSort(key) {
    if (sortState.key === key) {
        sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
    } else {
        sortState.key = key;
        sortState.dir = key === 'name' ? 'asc' : 'desc';
    }

    currentPage = 1;
    invalidateTableViewCache();
    updateSortIndicators();
    renderTable();
}

function updateSortIndicators() {
    $tableHead.querySelectorAll('th[data-sort-key]').forEach((header) => {
        const key = header.dataset.sortKey;
        const icon = header.querySelector('.sort-icon');
        const isActive = key === sortState.key;
        const direction = isActive ? sortState.dir : 'none';

        header.classList.toggle('sorted', isActive);
        header.setAttribute('aria-sort', direction === 'asc' ? 'ascending' : direction === 'desc' ? 'descending' : 'none');

        if (icon) {
            icon.textContent = direction === 'asc' ? '^' : direction === 'desc' ? 'v' : '<>';
        }
    });
}

function handleSearchInput() {
    const nextQuery = $tableSearch.value.trim().toLowerCase();

    if (searchDebounceId !== null) {
        window.clearTimeout(searchDebounceId);
    }

    searchDebounceId = window.setTimeout(() => {
        searchDebounceId = null;
        if (nextQuery === searchQuery) return;

        searchQuery = nextQuery;
        currentPage = 1;
        invalidateTableViewCache();
        renderTable();
    }, SEARCH_DEBOUNCE_MS);
}

function invalidateTableViewCache() {
    tableViewCache.key = '';
}

function getTableViewState() {
    const cacheKey = [
        tableDataVersion,
        searchQuery,
        statusFilterValue,
        sortState.key,
        sortState.dir,
    ].join('::');

    if (tableViewCache.key === cacheKey) {
        return tableViewCache;
    }

    const rows = getFilteredAndSortedRows();

    tableViewCache = {
        key: cacheKey,
        rows,
    };

    return tableViewCache;
}

function getFilteredAndSortedRows() {
    const filteredRows = tableSummaries.filter((row) => {
        const hasAnyData = row.total > 0;
        if (!hasAnyData) return false;

        const matchesSearch = row.name.toLowerCase().includes(searchQuery);
        const matchesStatus = statusFilterValue === 'all' ? true : row.counts[statusFilterValue] > 0;
        return matchesSearch && matchesStatus;
    });

    filteredRows.sort((a, b) => {
        if (sortState.key === 'name') {
            const cmp = NAME_COLLATOR.compare(a.name, b.name);
            return sortState.dir === 'asc' ? cmp : -cmp;
        }

        const aValue = sortState.key === 'total' ? a.total : (a.counts[sortState.key] || 0);
        const bValue = sortState.key === 'total' ? b.total : (b.counts[sortState.key] || 0);
        if (aValue === bValue) {
            return NAME_COLLATOR.compare(a.name, b.name);
        }
        const numericCmp = aValue - bValue;
        return sortState.dir === 'asc' ? numericCmp : -numericCmp;
    });

    return filteredRows;
}

/* ===== Routing Helpers ===== */
function getApplicantFromRoute() {
    const params = new URLSearchParams(window.location.search);
    const applicant = params.get(APPLICANT_QUERY_PARAM);
    return applicant ? applicant.trim() : '';
}

function buildDashboardHref({ applicantName = '', contextKey = activeDashboardContextKey } = {}) {
    const url = new URL(window.location.href);
    if (applicantName) {
        url.searchParams.set(APPLICANT_QUERY_PARAM, applicantName);
    } else {
        url.searchParams.delete(APPLICANT_QUERY_PARAM);
    }

    if (contextKey && contextKey !== DEFAULT_DASHBOARD_CONTEXT_KEY) {
        url.searchParams.set(CONTEXT_QUERY_PARAM, contextKey);
    } else {
        url.searchParams.delete(CONTEXT_QUERY_PARAM);
    }

    return `${url.pathname}${url.search}${url.hash}`;
}

function buildApplicantHref(name) {
    return buildDashboardHref({ applicantName: name });
}

function getDashboardHref() {
    return buildDashboardHref();
}

function clearApplicantRoute() {
    const nextHref = getDashboardHref();
    window.history.replaceState({}, '', nextHref);
}

function togglePageSections(mode) {
    const isApplicantMode = mode === 'applicant';
    const isAvpnMode = mode === 'avpn';
    const hideDashboardSections = isApplicantMode || isAvpnMode;

    $kpiStrip.classList.toggle('hidden-view', isApplicantMode);
    $summarySection.classList.toggle('hidden-view', hideDashboardSections);
    $tableSection.classList.toggle('hidden-view', hideDashboardSections);
    $acceptedTableSection.classList.toggle('hidden-view', hideDashboardSections || acceptedApplicantSummaries.length === 0);
    $applicantPage.classList.toggle('hidden-view', !isApplicantMode);
    $avpnSection.classList.toggle('hidden-view', !isAvpnMode);
}

function findApplicantSummary(applicantName) {
    const normalizedApplicant = String(applicantName || '').trim().toLowerCase();
    if (!normalizedApplicant) return null;
    return (
        tableSummaries.find((item) => item.name.toLowerCase() === normalizedApplicant) ||
        acceptedApplicantSummaries.find((item) => item.name.toLowerCase() === normalizedApplicant) ||
        null
    );
}

async function getApplicantRecords(summary) {
    if (!summary?.tableId) return [];

    const activeContext = getActiveDashboardContext();
    const baseId = activeContext?.baseId || '';
    const cacheKey = `${baseId}::${summary.tableId}`;

    if (applicantRecordsCache.has(cacheKey)) {
        return applicantRecordsCache.get(cacheKey);
    }

    const records = await fetchRecordsFromTable(summary.tableId, { baseId });
    applicantRecordsCache.set(cacheKey, records);
    return records;
}

function isModifiedClick(event) {
    return event.defaultPrevented || event.button !== 0 || event.metaKey || event.ctrlKey || event.altKey || event.shiftKey;
}

async function openApplicantFromRoute(applicantName, pushHistory = true) {
    const normalizedName = String(applicantName || '').trim();
    if (!normalizedName) return;

    activeApplicant = normalizedName;
    togglePageSections('applicant');
    showLoading();

    try {
        const rendered = await renderApplicantPage(normalizedName);
        if (!rendered) {
            showError(`Applicant "${normalizedName}" was not found. Showing dashboard list.`);
            activeApplicant = '';
            clearApplicantRoute();
            togglePageSections('dashboard');
            renderTable();
            chartInstance?.resize();
            return;
        }

        if (pushHistory) {
            window.history.pushState({}, '', buildApplicantHref(normalizedName));
        }
    } catch (err) {
        showError(err.message);
        activeApplicant = '';
        clearApplicantRoute();
        togglePageSections('dashboard');
        renderTable();
        chartInstance?.resize();
    } finally {
        hideLoading();
    }
}

async function handleApplicantLinkClick(event) {
    const link = event.target.closest('a.applicant-link');
    if (!link || isModifiedClick(event)) return;

    event.preventDefault();
    if (isLoadingData) return;

    const applicantName = link.dataset.applicantName || link.textContent;
    await openApplicantFromRoute(applicantName, true);
}

function handleApplicantBackClick(event) {
    const backLink = event.target.closest('a.back-link');
    if (!backLink || isModifiedClick(event)) return;

    event.preventDefault();
    activeApplicant = '';
    window.history.pushState({}, '', getDashboardHref());
    togglePageSections('dashboard');
    renderTable();
    chartInstance?.resize();
}

async function handlePopState() {
    if (isLoadingData) return;

    const contextFromRoute = getDashboardContextFromRoute();
    const contextChanged = contextFromRoute !== activeDashboardContextKey;
    if (contextChanged) {
        activeDashboardContextKey = contextFromRoute;
        activeApplicant = '';
        syncDashboardContextUI();
        togglePageSections(contextFromRoute === AVPN_CONTEXT_KEY ? 'avpn' : 'dashboard');
        await loadData();
    }

    if (getActiveDashboardContext()?.key === AVPN_CONTEXT_KEY) {
        activeApplicant = '';
        togglePageSections('avpn');
        return;
    }

    if (tableSummaries.length === 0) return;

    const applicantFromRoute = getApplicantFromRoute();
    if (!applicantFromRoute) {
        activeApplicant = '';
        togglePageSections('dashboard');
        renderTable();
        chartInstance?.resize();
        return;
    }

    await openApplicantFromRoute(applicantFromRoute, false);
}

/* ===== Formatting Helpers ===== */
function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function toCssToken(value) {
    return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

function formatFieldValue(value) {
    if (value === null || value === undefined || value === '') {
        return '-';
    }

    if (Array.isArray(value)) {
        const compact = value.map((item) => formatFieldValue(item)).join(', ');
        return compact || '-';
    }

    if (typeof value === 'object') {
        if (typeof value.name === 'string') return value.name;
        return JSON.stringify(value);
    }

    const formatted = String(value);
    return formatted.length > 140 ? `${formatted.slice(0, 137)}...` : formatted;
}

function normalizeFieldName(fieldName) {
    return String(fieldName || '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
}

function selectBestField(fieldNames, scorer, minScore = 1) {
    let bestField = null;
    let bestScore = minScore - 1;

    fieldNames.forEach((fieldName) => {
        const score = scorer(fieldName);
        if (score > bestScore) {
            bestScore = score;
            bestField = fieldName;
            return;
        }

        if (score === bestScore && score >= minScore && bestField && NAME_COLLATOR.compare(fieldName, bestField) < 0) {
            bestField = fieldName;
        }
    });

    return bestScore >= minScore ? bestField : null;
}

function scoreInvoiceStatusField(fieldName) {
    const normalized = normalizeFieldName(fieldName);
    if (!normalized.includes('invoice')) return 0;

    let score = 40;
    if (normalized === 'invoice status') score += 90;
    if (normalized.includes('status')) score += 30;
    if (normalized.includes('payment')) score += 12;

    return score;
}

function scoreTrainingDateField(fieldName) {
    const normalized = normalizeFieldName(fieldName);
    const hasTraining = normalized.includes('training');
    const hasDate = normalized.includes('date');
    const hasSession = normalized.includes('session') || normalized.includes('cohort');

    if (!hasTraining && !(hasDate && hasSession)) {
        return 0;
    }

    let score = 0;
    if (hasTraining) score += 60;
    if (hasDate) score += 40;
    if (hasSession) score += 20;
    if (normalized.includes('start')) score += 8;
    if (normalized.includes('end')) score += 8;
    if (normalized.includes('schedule')) score += 8;

    return score;
}

function resolveApplicantInsightFields(records) {
    const fieldNameSet = new Set();

    records.forEach((record) => {
        Object.keys(record.fields || {}).forEach((fieldName) => {
            fieldNameSet.add(fieldName);
        });
    });

    const fieldNames = Array.from(fieldNameSet);
    const invoiceField = selectBestField(fieldNames, scoreInvoiceStatusField, 20);

    const trainingFields = fieldNames
        .map((fieldName) => ({ fieldName, score: scoreTrainingDateField(fieldName) }))
        .filter((entry) => entry.score > 0)
        .sort((a, b) => b.score - a.score || NAME_COLLATOR.compare(a.fieldName, b.fieldName))
        .slice(0, 3)
        .map((entry) => entry.fieldName);

    return { invoiceField, trainingFields };
}

async function renderApplicantPage(applicantName) {
    const summary = findApplicantSummary(applicantName);
    if (!summary) {
        return false;
    }

    togglePageSections('applicant');
    const records = await getApplicantRecords(summary);
    const { invoiceField, trainingFields } = resolveApplicantInsightFields(records);
    const missingFieldNotes = [];
    if (!invoiceField) missingFieldNotes.push('Invoice status field was not found in this applicant data');
    if (trainingFields.length === 0) missingFieldNotes.push('Training date field was not found in this applicant data');

    // ── Status breakdown (already pre-computed in summary.counts) ──────────
    const nonZeroStatusEntries = TRAINING_STATUSES
        .map((status) => ({
            status,
            count: summary.counts[status] || 0,
            color: STATUS_COLORS[status] || '#8b6a3d',
        }))
        .filter((entry) => entry.count > 0)
        .sort((a, b) => b.count - a.count || (STATUS_ORDER.get(a.status) ?? 999) - (STATUS_ORDER.get(b.status) ?? 999));

    const dominantStatus = nonZeroStatusEntries[0] || null;
    const statusBreakdownRows = nonZeroStatusEntries.length === 0
        ? '<p class="summary-empty-note">No status data available.</p>'
        : nonZeroStatusEntries.map((entry) => {
            const { status, count, color } = entry;
            const pct = summary.total > 0 ? Math.round((count / summary.total) * 100) : 0;
            return `
                <div class="applicant-status-row">
                    <div class="applicant-status-main">
                        <span class="applicant-status-dot" style="--status-color:${color}"></span>
                        <span class="applicant-status-label">${escapeHtml(status)}</span>
                    </div>
                    <span class="applicant-status-count">${count.toLocaleString()}</span>
                    <div class="applicant-status-bar">
                        <div class="applicant-status-bar-fill" style="width:${pct}%;background:${color};"></div>
                    </div>
                    <span class="applicant-status-pct">${pct}%</span>
                </div>`;
        }).join('');

    // ── Invoice status breakdown (tallied from records) ────────────────────
    let invoiceSection = '';
    if (invoiceField && records.length > 0) {
        const invoiceCounts = {};
        records.forEach((record) => {
            const raw = formatFieldValue(record.fields?.[invoiceField]);
            const val = raw === '-' ? 'No Invoice' : raw;
            invoiceCounts[val] = (invoiceCounts[val] || 0) + 1;
        });
        const invoiceRows = Object.entries(invoiceCounts)
            .sort((a, b) => b[1] - a[1] || NAME_COLLATOR.compare(a[0], b[0]));
        const invoiceTotal = invoiceRows.reduce((sum, [, count]) => sum + count, 0);
        const invoiceRowsMarkup = invoiceRows.map(([val, cnt]) => {
            const color = getInvoiceStatusColor(val);
            const pctLabel = formatLegendPercent(cnt, invoiceTotal);
            const pct = invoiceTotal > 0 ? (cnt / invoiceTotal) * 100 : 0;
            const barWidth = pct <= 0 ? 0 : Math.max(3, Math.min(100, pct));
            return `
                <div class="applicant-invoice-row">
                    <div class="applicant-invoice-main">
                        <span class="applicant-invoice-dot" style="--invoice-color:${color}"></span>
                        <span class="applicant-invoice-status">${escapeHtml(val)}</span>
                    </div>
                    <span class="applicant-invoice-count">${cnt.toLocaleString()}</span>
                    <span class="applicant-invoice-share">${pctLabel}</span>
                </div>
                <div class="applicant-invoice-bar" aria-hidden="true">
                    <span style="width:${barWidth}%;--invoice-color:${color}"></span>
                </div>`;
        }).join('');
        invoiceSection = `
            <article class="applicant-summary-card applicant-summary-card--detail applicant-summary-card--invoice">
                <div class="summary-card-head">
                    <div>
                        <h3 class="summary-card-title">Invoice Status Breakdown</h3>
                        <p class="applicant-panel-copy">Operational billing progress for this applicant.</p>
                    </div>
                </div>
                <div class="applicant-invoice-list">
                    <div class="applicant-invoice-head" role="presentation">
                        <span>Status</span>
                        <span>Count</span>
                        <span>Share</span>
                    </div>
                    ${invoiceRowsMarkup}
                </div>
            </article>`;
    }

    // ── Training dates (unique values across all records) ──────────────────
    let trainingSection = '';
    if (trainingFields.length > 0 && records.length > 0) {
        const trainingRows = trainingFields.map((fieldName) => {
            const values = [...new Set(
                records
                    .map((r) => formatFieldValue(r.fields?.[fieldName]))
                    .filter((v) => v !== '-')
            )];
            return values.length === 0 ? '' : `
                <div class="applicant-training-item">
                    <span class="training-date-field">${escapeHtml(fieldName)}</span>
                    <span class="summary-value">${values.map(escapeHtml).join(', ')}</span>
                </div>`;
        }).join('');

        if (trainingRows.trim()) {
            trainingSection = `
                <article class="applicant-summary-card applicant-summary-card--detail applicant-summary-card--training">
                    <div class="summary-card-head">
                        <div>
                            <h3 class="summary-card-title">Training Dates</h3>
                            <p class="applicant-panel-copy">Latest schedule-related data captured on this applicant.</p>
                        </div>
                    </div>
                    <div class="applicant-training-list">${trainingRows}</div>
                </article>`;
        }
    }

    const emptyContent = records.length === 0
        ? `<div class="empty-state">No records found for this applicant.</div>` : '';

    $applicantPage.innerHTML = `
        <div class="applicant-page-shell">
            <div class="applicant-hero">
                <div class="applicant-hero-copy">
                    <a class="back-link" href="${getDashboardHref()}">Back to Applicant List</a>
                    <p class="applicant-eyebrow">Applicant Overview</p>
                    <div class="applicant-page-title-wrap">
                        <h2 class="section-title applicant-page-title">${escapeHtml(summary.name)}</h2>
                        <p class="applicant-meta">${summary.total.toLocaleString()} total record${summary.total === 1 ? '' : 's'} across this applicant dataset</p>
                    </div>
                </div>
                <div class="applicant-hero-stats">
                    <article class="applicant-hero-stat">
                        <p class="applicant-hero-stat-label">Total Records</p>
                        <p class="applicant-hero-stat-value">${summary.total.toLocaleString()}</p>
                        <p class="applicant-hero-stat-meta">Current applicant volume</p>
                    </article>
                    <article class="applicant-hero-stat">
                        <p class="applicant-hero-stat-label">Lead Status</p>
                        <p class="applicant-hero-stat-value">${dominantStatus ? dominantStatus.count.toLocaleString() : '-'}</p>
                        <p class="applicant-hero-stat-meta">${dominantStatus ? escapeHtml(dominantStatus.status) : 'No status data'}</p>
                    </article>
                    <article class="applicant-hero-stat">
                        <p class="applicant-hero-stat-label">Active Statuses</p>
                        <p class="applicant-hero-stat-value">${nonZeroStatusEntries.length.toLocaleString()}</p>
                        <p class="applicant-hero-stat-meta">Non-zero status buckets</p>
                    </article>
                </div>
            </div>
            ${missingFieldNotes.length > 0 ? `<p class="applicant-warning">${escapeHtml(missingFieldNotes.join('. '))}.</p>` : ''}
            ${emptyContent}
            ${records.length > 0 ? `
            <div class="applicant-summary-grid applicant-summary-grid--refined">
                <article class="applicant-summary-card applicant-summary-card--wide applicant-summary-card--status">
                    <div class="summary-card-head">
                        <div>
                            <h3 class="summary-card-title">Registration Status</h3>
                            <p class="applicant-panel-copy">Current registration distribution across this applicant's records.</p>
                        </div>
                        <span class="agg-total-badge">${summary.total.toLocaleString()} total</span>
                    </div>
                    <div class="applicant-status-list">${statusBreakdownRows}</div>
                </article>
                ${invoiceSection}
                ${trainingSection}
            </div>` : ''}
        </div>
    `;

    return true;
}

function normalizeAvpnSummaryRows(rows) {
    const byCategory = new Map(
        AVPN_CATEGORY_ORDER.map((category) => [category, {
            category,
            registeredInLms: 0,
            completed50: 0,
            completed100: 0,
            completedPreSurvey: 0,
        }])
    );

    (rows || []).forEach((row) => {
        const rowCategory = String(row?.category || '').trim().toLowerCase();
        const normalizedCategory = AVPN_CATEGORY_ORDER.find((category) => category.toLowerCase() === rowCategory);
        if (!normalizedCategory) return;

        byCategory.set(normalizedCategory, {
            category: normalizedCategory,
            registeredInLms: Number(row.registeredInLms ?? row.registered_in_lms ?? 0) || 0,
            completed50: Number(row.completed50 ?? row.completed_50 ?? 0) || 0,
            completed100: Number(row.completed100 ?? row.completed_100 ?? 0) || 0,
            completedPreSurvey: Number(row.completedPreSurvey ?? row.completed_pre_survey ?? 0) || 0,
        });
    });

    return AVPN_CATEGORY_ORDER.map((category) => byCategory.get(category));
}

function calculateWeeklyIncrement(current, baseline) {
    if (!Number.isFinite(baseline) || baseline <= 0) return 0;
    return ((current - baseline) / baseline) * 100;
}

function formatPercentInteger(value) {
    if (!Number.isFinite(value)) return '-';
    return `${Math.round(value)}`;
}

function renderAvpnSummaryTables(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];

    $avpnSummaryBody.innerHTML = safeRows.map((row) => `
        <tr>
            <td>${escapeHtml(row.category)}</td>
            <td>${row.registeredInLms.toLocaleString()}</td>
            <td>${row.completed50.toLocaleString()}</td>
            <td>${row.completed100.toLocaleString()}</td>
            <td>${row.completedPreSurvey.toLocaleString()}</td>
        </tr>
    `).join('');

    const totals = safeRows.reduce((acc, row) => ({
        registeredInLms: acc.registeredInLms + row.registeredInLms,
        completed50: acc.completed50 + row.completed50,
        completed100: acc.completed100 + row.completed100,
        completedPreSurvey: acc.completedPreSurvey + row.completedPreSurvey,
    }), { registeredInLms: 0, completed50: 0, completed100: 0, completedPreSurvey: 0 });

    $avpnSummaryFooter.innerHTML = `
        <tr class="totals-row">
            <th>Total</th>
            <th>${totals.registeredInLms.toLocaleString()}</th>
            <th>${totals.completed50.toLocaleString()}</th>
            <th>${totals.completed100.toLocaleString()}</th>
            <th>${totals.completedPreSurvey.toLocaleString()}</th>
        </tr>
    `;

    $avpnTargetBody.innerHTML = safeRows.map((row) => {
        const current = row.registeredInLms;
        const target = AVPN_TARGETS[row.category] ?? 0;
        const baseline = AVPN_BASELINES[row.category] ?? 0;
        const weeklyIncrement = calculateWeeklyIncrement(current, baseline);

        return `
            <tr>
                <td>${escapeHtml(row.category)}</td>
                <td>${current.toLocaleString()}</td>
                <td>${target.toLocaleString()}</td>
                <td>${baseline > 0 ? formatPercentInteger(weeklyIncrement) : '-'}</td>
            </tr>
        `;
    }).join('');
}

function updateAvpnGeneratedAt(rawTimestamp) {
    if (!rawTimestamp) {
        $avpnGeneratedAt.textContent = 'Updated: --';
        return;
    }

    const parsed = new Date(rawTimestamp);
    if (Number.isNaN(parsed.getTime())) {
        $avpnGeneratedAt.textContent = 'Updated: --';
        return;
    }

    $avpnGeneratedAt.textContent = `Updated: ${parsed.toLocaleString()}`;
}

/* ===== Rendering ===== */
function buildChartEntries(totals) {
    const eligibleStatuses = getChartEligibleStatuses();
    const eligibleSet = new Set(eligibleStatuses);

    const allEntries = TRAINING_STATUSES
        .filter((status) => eligibleSet.has(status))
        .map((status) => ({
            status,
            label: status,
            value: Number(totals?.[status]) || 0,
            color: STATUS_COLORS[status] || '#9ca3af',
            group: WORKFLOW_STATUS_SET.has(status) ? 'workflow' : 'data-quality',
        }));

    let visibleEntries = allEntries.filter((entry) => chartStatusVisibility.get(entry.status) !== false);
    if (visibleEntries.length === 0 && allEntries.length > 0) {
        chartStatusVisibility.set(allEntries[0].status, true);
        visibleEntries = [allEntries[0]];
    }

    const sortedVisible = [...visibleEntries]
        .filter((entry) => entry.value > 0)
        .sort((a, b) => b.value - a.value || (STATUS_ORDER.get(a.status) ?? 999) - (STATUS_ORDER.get(b.status) ?? 999));

    const sourceForChart = sortedVisible.length > 0
        ? sortedVisible
        : [...visibleEntries].sort((a, b) => (STATUS_ORDER.get(a.status) ?? 999) - (STATUS_ORDER.get(b.status) ?? 999));

    const displayEntries = sourceForChart.slice(0, CHART_MAX_SEGMENTS).map((entry) => ({
        ...entry,
        members: null,
    }));

    if (sourceForChart.length > CHART_MAX_SEGMENTS) {
        const remainder = sourceForChart.slice(CHART_MAX_SEGMENTS);
        const otherValue = remainder.reduce((sum, entry) => sum + entry.value, 0);
        displayEntries.push({
            status: CHART_OTHER_LABEL,
            label: CHART_OTHER_LABEL,
            value: otherValue,
            color: CHART_OTHER_COLOR,
            group: 'mixed',
            members: remainder.map((entry) => entry.label),
        });
    }

    return {
        allEntries,
        displayEntries,
    };
}

function renderStatusLegend(allEntries) {
    const eligibleTotal = allEntries.reduce((sum, entry) => sum + entry.value, 0);
    const byStatus = new Map(allEntries.map((entry) => [entry.status, entry]));

    const groupedMarkup = STATUS_GROUPS.map((group) => {
        const entries = group.statuses
            .map((status) => byStatus.get(status))
            .filter((entry) => entry && entry.value > 0);

        if (entries.length === 0) return '';

        const groupTotal = entries.reduce((sum, entry) => sum + entry.value, 0);
        const rows = entries.map((entry) => {
            const isHidden = chartStatusVisibility.get(entry.status) === false;
            const isIsolated = chartIsolatedStatus === entry.status;
            return `
                <button type="button" class="status-legend-item${isHidden ? ' is-muted' : ''}${isIsolated ? ' is-isolated' : ''}" data-status="${escapeHtml(entry.status)}" aria-pressed="${String(!isHidden)}">
                    <span class="status-legend-left">
                        <span class="status-legend-dot" style="--legend-color:${entry.color}"></span>
                        <span class="status-legend-label">${escapeHtml(entry.label)}</span>
                    </span>
                    <span class="status-legend-metric">${entry.value.toLocaleString()} (${formatLegendPercent(entry.value, eligibleTotal)})</span>
                </button>
            `;
        }).join('');

        return `
            <section class="status-legend-group">
                <div class="status-legend-group-head">
                    <h3>${escapeHtml(group.label)}</h3>
                    <span>${groupTotal.toLocaleString()}</span>
                </div>
                <div class="status-legend-list">${rows}</div>
            </section>
        `;
    }).join('');

    $statusLegendGroups.innerHTML = groupedMarkup || '<p class="summary-empty-note">No status data available.</p>';
}

function renderChart(totals) {
    latestChartTotals = totals;
    const canvas = document.getElementById('status-chart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    renderAnalyticsKpis(totals);
    ensureChartHasVisibleStatus();
    const { allEntries, displayEntries } = buildChartEntries(totals);
    renderStatusLegend(allEntries);

    const labels = displayEntries.map((entry) => entry.label);
    const data = displayEntries.map((entry) => entry.value);
    const backgroundColors = displayEntries.map((entry) => entry.color);
    const otherMembers = displayEntries.map((entry) => entry.members);

    const chartDataset = {
        data,
        backgroundColor: backgroundColors,
        otherMembers,
        borderWidth: 0,
        hoverOffset: 8,
        spacing: 4,
        borderRadius: 6,
        cutout: '74%',
        radius: '82%',
    };

    // Destroy any chart already attached to the canvas (e.g. after HMR / module reload)
    const existingChart = Chart.getChart(canvas);
    if (existingChart && existingChart !== chartInstance) {
        existingChart.destroy();
        chartInstance = null;
    }

    if (!chartInstance) {
        chartInstance = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [chartDataset]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: {
                    duration: 650,
                    easing: 'easeOutCubic',
                },
                layout: {
                    padding: 8,
                },
                plugins: {
                    legend: {
                        display: false,
                    },
                    tooltip: {
                        backgroundColor: 'rgba(255, 250, 240, 0.98)',
                        titleColor: '#4b3a2a',
                        bodyColor: '#6d5841',
                        borderColor: 'rgba(154, 119, 67, 0.28)',
                        borderWidth: 1,
                        padding: 12,
                        cornerRadius: 10,
                        displayColors: true,
                        titleFont: {
                            weight: 700,
                        },
                        bodyFont: {
                            weight: 600,
                        },
                        callbacks: {
                            title(context) {
                                return context[0]?.label || 'Status';
                            },
                            label(context) {
                                const value = context.parsed || 0;
                                const sum = context.dataset.data.reduce((acc, n) => acc + n, 0);
                                const pct = sum > 0 ? ((value / sum) * 100).toFixed(1) : '0.0';
                                return `Count: ${value.toLocaleString()} (${pct}%)`;
                            },
                            footer(context) {
                                const index = context?.[0]?.dataIndex ?? -1;
                                const members = context?.[0]?.dataset?.otherMembers?.[index];
                                if (!Array.isArray(members) || members.length === 0) return '';

                                const preview = members.slice(0, 3).join(', ');
                                const suffix = members.length > 3 ? ` +${members.length - 3} more` : '';
                                return `Includes: ${preview}${suffix}`;
                            }
                        }
                    }
                }
            }
        });
        return;
    }

    chartInstance.data.labels = labels;
    chartInstance.data.datasets[0] = chartDataset;
    chartInstance.update();
}

function getInvoiceStatusColor(status) {
    const normalized = String(status || '').toLowerCase();

    if (normalized.includes('payment complete') || normalized.includes('paid') || normalized.includes('completed')) {
        return '#7c8f6f';
    }

    if (normalized.includes('invoice issued') || normalized.includes('uploaded') || normalized.includes('portal')) {
        return '#e97824';
    }

    if (normalized.includes('no invoice') || normalized.includes('missing') || normalized.includes('not')) {
        return '#9a9488';
    }

    return '#8f7350';
}

async function renderAcceptedApplicantsCard() {
    const rows = acceptedApplicantSummaries.filter((row) => row.total > 0);
    const totals = calculateGrandTotals(rows);

    $acceptedTableCount.textContent = `${rows.length} source row${rows.length === 1 ? '' : 's'}`;
    $acceptedCount.textContent = (totals.Accepted || 0).toLocaleString();

    if (rows.length === 0) {
        $acceptedInvoiceList.innerHTML = '<p class="summary-empty-note">No invoice data yet.</p>';
        return;
    }

    const invoiceCounts = {};
    let sawInvoiceField = false;

    await Promise.all(rows.map(async (summary) => {
        try {
            const records = await getApplicantRecords(summary);
            if (records.length === 0) return;

            const { invoiceField } = resolveApplicantInsightFields(records);
            if (!invoiceField) return;

            sawInvoiceField = true;
            records.forEach((record) => {
                const raw = formatFieldValue(record.fields?.[invoiceField]);
                const status = raw === '-' ? 'No Invoice' : raw;
                invoiceCounts[status] = (invoiceCounts[status] || 0) + 1;
            });
        } catch (error) {
            console.warn(`Unable to compute invoice statuses for accepted source "${summary.name}".`, error);
        }
    }));

    const invoiceRows = Object.entries(invoiceCounts)
        .sort((a, b) => b[1] - a[1] || NAME_COLLATOR.compare(a[0], b[0]));

    if (invoiceRows.length === 0) {
        $acceptedInvoiceList.innerHTML = `<p class="summary-empty-note">${sawInvoiceField ? 'No invoice statuses found.' : 'Invoice status field not found in accepted source tables.'}</p>`;
        return;
    }

    const invoiceTotal = invoiceRows.reduce((sum, [, count]) => sum + count, 0);
    const invoiceMarkup = invoiceRows
        .map(([status, count]) => {
            const pct = invoiceTotal > 0 ? (count / invoiceTotal) * 100 : 0;
            const pctLabel = pct > 0 && pct < 1 ? '<1%' : `${pct.toFixed(1)}%`;
            const barWidthPct = pct <= 0 ? 0 : Math.max(3, Math.min(100, pct));
            const rowColor = getInvoiceStatusColor(status);

            return `
                <div class="accepted-invoice-row">
                    <div class="accepted-invoice-main">
                        <span class="accepted-invoice-dot" style="--invoice-color:${rowColor}"></span>
                        <span class="accepted-invoice-status">${escapeHtml(status)}</span>
                    </div>
                    <span class="accepted-invoice-count">${count.toLocaleString()}</span>
                    <span class="accepted-invoice-share">${pctLabel}</span>
                </div>
                <div class="accepted-invoice-bar" aria-hidden="true">
                    <span style="width:${barWidthPct}%;--invoice-color:${rowColor}"></span>
                </div>
            `;
        })
        .join('');

    $acceptedInvoiceList.innerHTML = `
        <div class="accepted-invoice-head" role="presentation">
            <span>Status</span>
            <span>Count</span>
            <span>Share</span>
        </div>
        ${invoiceMarkup}
    `;
}

function renderTable() {
    const { rows: filteredRows } = getTableViewState();
    const visibleStatuses = TABLE_STATUS_COLUMNS;
    const nonEmptyRowCount = tableSummaries.reduce((count, row) => count + (row.total > 0 ? 1 : 0), 0);

    const totalPages = Math.max(1, Math.ceil(filteredRows.length / rowsPerPage));
    currentPage = Math.min(currentPage, totalPages);

    const start = (currentPage - 1) * rowsPerPage;
    const end = start + rowsPerPage;
    const pagedData = filteredRows.slice(start, end);
    const visibleColumnCount = visibleStatuses.length + 2;

    if (pagedData.length === 0) {
        $tableBody.innerHTML = `<tr><td colspan="${visibleColumnCount}" class="empty-state">No data available</td></tr>`;
        $tableFooter.innerHTML = '';
        updatePaginationControls(totalPages);
        updateHeaderStats(nonEmptyRowCount, filteredRows.length);
        return;
    }

    $tableBody.innerHTML = pagedData.map((row) => `
        <tr>
            <td class="table-name-cell sticky-col-left">
                <a class="applicant-link" data-applicant-name="${escapeHtml(row.name)}" href="${buildApplicantHref(row.name)}" title="Open ${escapeHtml(row.name)} details">${escapeHtml(row.name)}</a>
            </td>
            ${visibleStatuses.map((status) => {
        const count = row.counts[status] || 0;
        if (count === 0) {
            return `
                    <td class="status-cell is-zero">
                        <span class="status-zero">-</span>
                    </td>
                `;
        }
        const badgeClass = `badge-${toCssToken(status)}`;
        return `
                    <td class="status-cell">
                        <span class="badge ${badgeClass}">${count.toLocaleString()}</span>
                    </td>
                `;
    }).join('')}
            <td class="sticky-col-right total-cell">${row.total.toLocaleString()}</td>
        </tr>
    `).join('');

    renderTableFooter(filteredRows, visibleStatuses);
    updatePaginationControls(totalPages);
    updateHeaderStats(nonEmptyRowCount, filteredRows.length);
}

function renderTableFooter(rows, visibleStatuses) {
    if (rows.length === 0) {
        $tableFooter.innerHTML = '';
        return;
    }

    const totals = calculateGrandTotals(rows);
    $tableFooter.innerHTML = `
        <tr class="totals-row">
            <th class="sticky-col-left">Filtered Total</th>
            ${visibleStatuses.map((status) => `<td class="status-cell footer-cell">${totals[status].toLocaleString()}</td>`).join('')}
            <th class="sticky-col-right total-cell">${totals.totalRecords.toLocaleString()}</th>
        </tr>
    `;
}

function updatePaginationControls(totalPages) {
    $currentPageLabel.textContent = String(currentPage);
    $totalPagesLabel.textContent = String(totalPages);
    $btnPrev.disabled = currentPage <= 1;
    $btnNext.disabled = currentPage >= totalPages;
}

function changePage(direction) {
    const totalPages = Math.max(1, Math.ceil(getTableViewState().rows.length / rowsPerPage));
    const newPage = currentPage + direction;

    if (newPage >= 1 && newPage <= totalPages) {
        currentPage = newPage;
        renderTable();
        $tableSection?.scrollIntoView({ behavior: 'smooth' });
    }
}

function updateHeaderStats(totalCount, filteredCount = totalCount) {
    if (!$tableTotalCount) return;

    if (filteredCount === totalCount) {
        $tableTotalCount.textContent = `${totalCount} applicant${totalCount !== 1 ? 's' : ''}`;
        return;
    }

    $tableTotalCount.textContent = `${filteredCount} of ${totalCount} applicants`;
}

function updateLastUpdated() {
    const now = new Date();
    $lastUpdated.textContent = `Last updated: ${now.toLocaleTimeString()}`;
}

/* ===== UI Helpers ===== */
function showLoading() { $loadingOverlay.classList.remove('hidden'); }
function hideLoading() { $loadingOverlay.classList.add('hidden'); }

function showError(message) {
    $errorMessage.textContent = message || 'Something went wrong while loading dashboard data.';
    $errorToast.classList.add('visible');
    if (errorHideTimerId !== null) {
        window.clearTimeout(errorHideTimerId);
    }
    errorHideTimerId = window.setTimeout(() => {
        $errorToast.classList.remove('visible');
        errorHideTimerId = null;
    }, 6000);
}

init();
