import { defineConfig, loadEnv } from 'vite';

const THINKIFIC_API_BASE = 'https://api.thinkific.com/api/public/v1';
const THINKIFIC_PAGE_SIZE = 100;
const THINKIFIC_FETCH_CONCURRENCY = 4;
const CATEGORY_ORDER = ['Students', 'Educators', 'SMEs'];
const AVPN_COURSE_CATEGORY_ENTRIES = [
  ['AVPN AI Opportunity Fund: Asia Pacific Phase 2 — Student Edition', 'Students'],
  ['AVPN AI Opportunity Fund: Asia Pacific Phase 2 —Education Workforce Edition', 'Educators'],
  ['AVPN AI Opportunity Fund: Asia Pacific Phase 2 — MSME Workforce Edition', 'SMEs'],
  ['AVPN AI Opportunity Fund: Asia Pacific Fasa 2 - Edisi Pelajar', 'Students'],
  ['AVPN AI Opportunity Fund: Asia Pacific Fasa 2 - Edisi Tenaga Kerja MSME', 'SMEs'],
  ['AVPN AI Opportunity Fund: Asia Pacific Fasa 2 - Edisi Tenaga Kerja Pendidikan', 'Educators'],
];
const AVPN_SUMMARY_CACHE_TTL_MS = 5 * 60 * 1000;
const avpnSummaryCache = {
  expiresAt: 0,
  payload: null,
  inFlight: null,
};

function parseJwtPayload(token) {
  if (!token || typeof token !== 'string') return null;
  const segments = token.split('.');
  if (segments.length < 2) return null;

  try {
    const payload = segments[1].replace(/-/g, '+').replace(/_/g, '/');
    const decoded = Buffer.from(payload, 'base64').toString('utf8');
    return JSON.parse(decoded);
  } catch {
    return null;
  }
}

function resolveSubdomain(env, accessToken) {
  const fromEnv = String(env.THINKIFIC_SUBDOMAIN || '').trim();
  if (fromEnv) return fromEnv;

  const jwtPayload = parseJwtPayload(accessToken);
  const fromToken = String(jwtPayload?.subdomain || '').trim();
  return fromToken;
}

function getThinkificAuthHeaders(env) {
  const accessToken = String(env.THINKIFIC_ACCESS_TOKEN || '').trim();
  const apiKey = String(env.THINKIFIC_API_KEY || '').trim();
  const subdomain = resolveSubdomain(env, accessToken);

  if (!accessToken && (!apiKey || !subdomain)) {
    throw new Error(
      'Missing Thinkific credentials. Set THINKIFIC_ACCESS_TOKEN or THINKIFIC_API_KEY + THINKIFIC_SUBDOMAIN.'
    );
  }

  const headers = {
    Accept: 'application/json',
  };

  if (accessToken) {
    headers.Authorization = `Bearer ${accessToken}`;
  }

  if (apiKey && subdomain) {
    headers['X-Auth-API-Key'] = apiKey;
    headers['X-Auth-Subdomain'] = subdomain;
  }

  return headers;
}

async function fetchThinkificPage(pathname, env, query = {}) {
  const headers = getThinkificAuthHeaders(env);
  const url = new URL(`${THINKIFIC_API_BASE}/${pathname}`);
  Object.entries(query).forEach(([key, value]) => {
    if (value === undefined || value === null || value === '') return;
    url.searchParams.set(key, String(value));
  });

  const response = await fetch(url, { headers });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Thinkific API request failed (${response.status}): ${body}`);
  }

  return response.json();
}

function extractCollection(data, collectionKey) {
  if (Array.isArray(data?.[collectionKey])) return data[collectionKey];
  if (Array.isArray(data?.items)) return data.items;
  if (Array.isArray(data?.data)) return data.data;
  return [];
}

function getPaginationTotalPages(data) {
  const pages = Number(
    data?.meta?.pagination?.total_pages ??
    data?.pagination?.total_pages ??
    data?.total_pages ??
    0
  );
  return Number.isFinite(pages) && pages > 0 ? pages : 0;
}

async function runWithConcurrencyLimit(items, limit, worker) {
  if (!Array.isArray(items) || items.length === 0) return [];

  const output = new Array(items.length);
  const workerCount = Math.max(1, Math.min(limit, items.length));
  let nextIndex = 0;

  async function runWorker() {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex++;
      output[currentIndex] = await worker(items[currentIndex], currentIndex);
    }
  }

  await Promise.all(Array.from({ length: workerCount }, () => runWorker()));
  return output;
}

async function fetchAllThinkificRecords(pathname, collectionKey, env, query = {}) {
  const records = [];
  let page = 1;

  while (true) {
    const data = await fetchThinkificPage(pathname, env, { ...query, page, limit: THINKIFIC_PAGE_SIZE });
    const pageItems = extractCollection(data, collectionKey);
    if (pageItems.length === 0) {
      break;
    }

    records.push(...pageItems);
    const totalPages = getPaginationTotalPages(data);
    if (totalPages > 0) {
      if (page >= totalPages) break;
    } else if (pageItems.length < THINKIFIC_PAGE_SIZE) {
      break;
    }

    page += 1;
  }

  return records;
}

function getCourseTitle(enrollment, courseTitleById) {
  const directTitle = enrollment?.course_name || enrollment?.course_title || enrollment?.product_name;
  if (directTitle) return String(directTitle);

  const courseId = enrollment?.course_id ?? enrollment?.course?.id;
  return courseTitleById.get(String(courseId ?? '')) || '';
}

function normalizeCourseTitle(titleText) {
  return String(titleText || '')
    .normalize('NFKC')
    .replace(/[–-]/g, '—')
    .replace(/\s*—\s*/g, ' — ')
    .replace(/\s+/g, ' ')
    .trim();
}

const AVPN_ALLOWED_COURSE_CATEGORY_BY_TITLE = new Map(
  AVPN_COURSE_CATEGORY_ENTRIES.map(([title, category]) => [normalizeCourseTitle(title), category])
);

function resolveAllowedAvpnCourseCategory(titleText) {
  return AVPN_ALLOWED_COURSE_CATEGORY_BY_TITLE.get(normalizeCourseTitle(titleText)) || null;
}

function isPreSurveyCourse(titleText) {
  const text = String(titleText || '').toLowerCase();
  return /\bpre[\s-]?survey\b/.test(text);
}

function getCompletionPercentage(enrollment) {
  const raw = enrollment?.percentage_completed
    ?? enrollment?.completion_percentage
    ?? enrollment?.percent_completed
    ?? enrollment?.progress_percentage
    ?? 0;
  const value = Number(raw);
  if (!Number.isFinite(value)) return 0;

  // Thinkific commonly returns progress as a 0..1 fraction (e.g. 0.42 = 42%).
  return value <= 1 ? value * 100 : value;
}

function getEnrollmentUserId(enrollment) {
  const raw = enrollment?.user_id ?? enrollment?.student_id ?? enrollment?.user?.id;
  return raw === null || raw === undefined ? '' : String(raw);
}

function buildEmptyCategoryBuckets() {
  return CATEGORY_ORDER.reduce((acc, category) => {
    acc[category] = {
      registeredUsers: new Set(),
      completed50Users: new Set(),
      completed100Users: new Set(),
      completedPreSurveyUsers: new Set(),
    };
    return acc;
  }, {});
}

async function buildAvpnSummary(env) {
  const courses = await fetchAllThinkificRecords('courses', 'courses', env);

  const courseTitleById = new Map(
    courses.map((course) => [String(course?.id ?? ''), String(course?.name ?? course?.title ?? '')])
  );
  const courseCategoryById = new Map();
  const matchedCatalogTitles = new Set();
  courses.forEach((course) => {
    const courseId = String(course?.id ?? '');
    if (!courseId) return;

    const courseTitle = String(course?.name ?? course?.title ?? '').trim();
    const category = resolveAllowedAvpnCourseCategory(courseTitle);
    if (!category) return;

    const normalizedTitle = normalizeCourseTitle(courseTitle);
    courseCategoryById.set(courseId, category);
    matchedCatalogTitles.add(normalizedTitle);
  });

  const allowedCourseIds = Array.from(courseCategoryById.keys()).filter(Boolean);
  const enrollmentPagesPerCourse = await runWithConcurrencyLimit(
    allowedCourseIds,
    THINKIFIC_FETCH_CONCURRENCY,
    async (courseId) => fetchAllThinkificRecords('enrollments', 'enrollments', env, { 'query[course_id]': courseId })
  );
  const enrollments = enrollmentPagesPerCourse.flat();

  const buckets = buildEmptyCategoryBuckets();

  enrollments.forEach((enrollment) => {
    const userId = getEnrollmentUserId(enrollment);
    if (!userId) return;

    const courseId = String(enrollment?.course_id ?? enrollment?.course?.id ?? '');
    const courseTitle = getCourseTitle(enrollment, courseTitleById);
    const categoryFromCourseId = courseCategoryById.get(courseId);
    const categoryFromTitle = resolveAllowedAvpnCourseCategory(courseTitle);
    const category = categoryFromCourseId || categoryFromTitle;
    if (!category || !buckets[category]) return;

    const completionPercentage = getCompletionPercentage(enrollment);
    const completed100 = completionPercentage >= 100 || Boolean(enrollment?.completed_at);

    buckets[category].registeredUsers.add(userId);
    // "Completed 50%" column is defined as completion progress between 40% and 59%.
    if (completionPercentage >= 40 && completionPercentage < 60) {
      buckets[category].completed50Users.add(userId);
    }
    if (completed100) {
      buckets[category].completed100Users.add(userId);
    }
    if (isPreSurveyCourse(courseTitle) && completed100) {
      buckets[category].completedPreSurveyUsers.add(userId);
    }
  });

  const rows = CATEGORY_ORDER.map((category) => {
    const bucket = buckets[category];
    return {
      category,
      registeredInLms: bucket.registeredUsers.size,
      completed50: bucket.completed50Users.size,
      completed100: bucket.completed100Users.size,
      completedPreSurvey: bucket.completedPreSurveyUsers.size,
    };
  });

  return {
    rows,
    diagnostics: {
      allowedCourseTitles: AVPN_COURSE_CATEGORY_ENTRIES.map(([title]) => title),
      allowedCourseIdsFetched: allowedCourseIds.length,
      allowedCourseCatalogMatches: matchedCatalogTitles.size,
      coursesFetched: courses.length,
      enrollmentsFetched: enrollments.length,
    },
    generatedAt: new Date().toISOString(),
  };
}

function sendJson(res, statusCode, payload) {
  res.statusCode = statusCode;
  res.setHeader('Content-Type', 'application/json');
  res.end(JSON.stringify(payload));
}

async function getCachedAvpnSummary(env, forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && avpnSummaryCache.payload && now < avpnSummaryCache.expiresAt) {
    return avpnSummaryCache.payload;
  }

  if (!forceRefresh && avpnSummaryCache.inFlight) {
    return avpnSummaryCache.inFlight;
  }

  const requestPromise = buildAvpnSummary(env)
    .then((payload) => {
      avpnSummaryCache.payload = payload;
      avpnSummaryCache.expiresAt = Date.now() + AVPN_SUMMARY_CACHE_TTL_MS;
      return payload;
    })
    .finally(() => {
      avpnSummaryCache.inFlight = null;
    });

  avpnSummaryCache.inFlight = requestPromise;
  return requestPromise;
}

async function thinkificMiddleware(req, res, next, env) {
  const requestUrl = new URL(req.url || '/', 'http://localhost');
  if (req.method !== 'GET' || requestUrl.pathname !== '/api/thinkific/avpn-summary') {
    next();
    return;
  }

  try {
    const refreshValue = String(requestUrl.searchParams.get('refresh') || '').toLowerCase();
    const forceRefresh = refreshValue === '1' || refreshValue === 'true' || refreshValue === 'yes';
    const payload = await getCachedAvpnSummary(env, forceRefresh);
    sendJson(res, 200, payload);
  } catch (error) {
    sendJson(res, 500, {
      error: error?.message || 'Unable to fetch AVPN Thinkific summary.',
    });
  }
}

function thinkificProxyPlugin(env) {
  const middlewareHandler = (req, res, next) => {
    void thinkificMiddleware(req, res, next, env).catch((error) => {
      sendJson(res, 500, {
        error: error?.message || 'Unexpected Thinkific proxy failure.',
      });
    });
  };

  return {
    name: 'thinkific-proxy-plugin',
    configureServer(server) {
      server.middlewares.use(middlewareHandler);
    },
    configurePreviewServer(server) {
      server.middlewares.use(middlewareHandler);
    },
  };
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '');

  return {
    server: {
      open: true,
      watch: {
        ignored: ['**/.env', '**/.env.*'],
      },
    },
    plugins: [thinkificProxyPlugin(env)],
  };
});
