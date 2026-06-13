// ============================================================
// QualityService.gs — Aggregations and logic for Quality Audits
// Source: 'PLX Raw data' tab
// ============================================================

var QUALITY_SHEET_NAME = 'PLX Raw data';
var CACHE_VERSION = 'v1.8'; // Global cache invalidation (bumped for new source)
var SOURCE_SPREADSHEET_ID = '1YDz16oRc2yi3sjyxtRPmaJbljyRPfYmwxVzuqiXu4Vw';

// ── SCHEMA MAPPING ────────────────────────────────────────────────────────

var Q_COLS = null; // Will be mapped dynamically

function getColMapping() {
  if (Q_COLS) return Q_COLS;

  var ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(QUALITY_SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + QUALITY_SHEET_NAME + "' not found.");

  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};

  var find = function(pattern) {
    var p = pattern.toLowerCase().replace(/[^a-z0-9]/g, '');

    // 1. Precise Match (stripped of symbols/spaces)
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === p) return i;
    }

    // 2. Fuzzy Match
    for (var j = 0; j < headers.length; j++) {
      var hj = String(headers[j]).toLowerCase().replace(/[^a-z0-9]/g, '');
      if (hj.indexOf(p) !== -1 || p.indexOf(hj) !== -1) return j;
    }
    return -1;
  };

  map.CASE_ID = find('__case_id__1');
map.ENTITY_GROUP = find('__billing_entity_group_name__1');
map.AGENT_LDAP = find('__agent_ldap__1');
map.OPENING_CHANNEL = find('__opening_channel__1');
map.REVIEW_DATE = find('__qplus_review_date__1');
map.REVIEW_WEEK = find('__qplus_review_week__1');
map.REVIEW_MONTH = find('__qplus_review_month__1');
map.CASE_DATE = find('__case_start_day__1');
map.CASE_WEEK = find('__case_start_week__1');
map.CASE_MONTH = find('__case_start_month__1');
map.CUSTOMER_CRITICAL = find('Customer');
map.BUSINESS_CRITICAL = find('Business');
map.COMPLIANCE_CRITICAL = find('Compliance');
map.REVIEWER_COMMENTS = find('Comment');

map.LISTENING = find('Listening');
map.PROBING = find('Probing');
map.COMPLETE_RESOLUTION = find('Providing Complete Resolution');
map.TROUBLESHOOTING = find('Troubleshooting');
map.USER_EXPECTATIONS = find('User Expectations');
map.EMPATHY = find('Empathizing');
map.OWNERSHIP = find('Ownership');
map.REFUNDS = find('Refunds');
map.RESPONSIVENESS = find('Responsiveness');

map.CONSULTS_ESCALATIONS = find('Consults');
map.CASE_DETAILS = find('Case details');
map.CATEGORIZATION = find('Categorization');
map.CSAT_REMINDER = find('CSAT reminder');
map.CASE_STATE = find('Case state');
map.OPENING_CLOSING = find('Opening');
map.LANGUAGE_PROFICIENCY = find('Language Proficiency');

map.AUTHENTICATION = find('Authentication');
map.GOOGLE_ONLY_INFO = find('Google-only');
map.PROFESSIONAL_CONDUCT = find('Professional Conduct');
map.PAYMENT_COMPLAINTS = find('Payment Complaints');

map.TEAM = find('Team');
map.SUPERVISOR = find('Supervisor LDAP');
map.MANAGER = find('Manager');
map.LOB = find('LOB');

  Q_COLS = map;
  return map;
}

var Q_TARGETS = {
  CUSTOMER: 95,
  BUSINESS: 90,
  COMPLIANCE: 99.50
};

var Q_PARAM_GROUPS = {
  customer: ['LISTENING', 'PROBING', 'COMPLETE_RESOLUTION', 'TROUBLESHOOTING', 'USER_EXPECTATIONS', 'EMPATHY', 'OWNERSHIP', 'REFUNDS', 'RESPONSIVENESS'],
  business: ['CONSULTS_ESCALATIONS', 'CASE_DETAILS', 'CATEGORIZATION', 'CSAT_REMINDER', 'CASE_STATE', 'OPENING_CLOSING', 'LANGUAGE_PROFICIENCY'],
  compliance: ['AUTHENTICATION', 'GOOGLE_ONLY_INFO', 'PROFESSIONAL_CONDUCT', 'PAYMENT_COMPLAINTS']
};

var Q_PARAM_COLS = [].concat(Q_PARAM_GROUPS.customer, Q_PARAM_GROUPS.business, Q_PARAM_GROUPS.compliance);

var _MEMOIZED_RAW_DATA = null;

// ── APPS SCRIPT WEB APP ───────────────────────────────────────────────────

function doGet() {
  var template = HtmlService.createTemplateFromFile('QualityView');

  try {
    var initialData = clientGetInitialData();
    template.bootstrap = JSON.stringify(initialData);
  } catch(e) {
    Logger.log('doGet Error: ' + e.message + ' | Stack: ' + e.stack);
    template.bootstrap = JSON.stringify({ error: e.message, stack: e.stack });
  }

  return template.evaluate()
    .setTitle('GenQA Scores')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── DATA LOADING ──────────────────────────────────────────────────────────

/**
 * Fetches raw data from the spreadsheet.
 * Optimization: We skip CacheService for the full raw dataset if it's large,
 * as the overhead of 100+ cache chunks often exceeds the time to read directly from the Sheet.
 */
/**
 * Fetches specific columns from the spreadsheet.
 * This is much faster than reading all columns for large sheets.
 */
function getColumnsFromSheet(colIndices, forceRefresh) {
  var ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(QUALITY_SHEET_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Sort and unique indices to fetch efficiently
  var uniqueIndices = Array.from(new Set(colIndices)).sort(function(a, b){return a-b});
  if (uniqueIndices[0] === -1) uniqueIndices.shift(); // Remove -1

  if (uniqueIndices.length === 0) return [];

  var startCol = uniqueIndices[0] + 1;
  var endCol = uniqueIndices[uniqueIndices.length - 1] + 1;
  var numCols = endCol - startCol + 1;

  // If we are fetching almost everything, just get the range
  // Otherwise, if columns are sparse, we might still just get the whole block
  // for simplicity in Apps Script, but limited by the actual used columns.
  var MAX_ROWS = 10000;
  var effectiveRows = Math.min(lastRow - 1, MAX_ROWS);
  var raw = sheet.getRange(3, startCol, effectiveRows, numCols).getValues();

  // Map back to the original order/indices requested
  return raw.map(function(row) {
    var mappedRow = {};
    colIndices.forEach(function(origIdx, i) {
      if (origIdx === -1) {
        mappedRow[i] = null;
      } else {
        mappedRow[i] = row[origIdx - (startCol - 1)];
      }
    });
    return mappedRow;
  });
}

function getRawQualityDataForMonth(month, forceRefresh) {
  var start = new Date().getTime();
  var cache = CacheService.getScriptCache();
  var cacheKey = 'raw_data_' + month + '_' + CACHE_VERSION;

  if (!forceRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) {
      var data = JSON.parse(cached);
      Logger.log('getRawQualityDataForMonth: Cache HIT (' + data.length + ' rows) for ' + month + ' in ' + (new Date().getTime() - start) + 'ms');
      return data;
    }
  }

  getColMapping();
  var ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(QUALITY_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // 1. Fetch Month column to identify relevant rows
  var monthIndices = [Q_COLS.REVIEW_MONTH];
  var monthRows = getColumnsFromSheet(monthIndices);

  var relevantMap = {};
  var firstMatch = -1, lastMatch = -1;
  for (var i = 0; i < monthRows.length; i++) {
    if (normalizeQualityMonth(monthRows[i][0]) === month) {
      var rowIdx = i + 3;
      relevantMap[rowIdx] = true;
      if (firstMatch === -1) firstMatch = rowIdx;
      lastMatch = rowIdx;
    }
  }

  if (firstMatch === -1) return [];

  // 2. Fetch the actual data for these rows
  // To be efficient, we fetch the bounding box of these rows if they are mostly contiguous
  var startRow = firstMatch;
  var endRow = lastMatch;

  var neededCols = [];
  for (var key in Q_COLS) {
    if (Q_COLS[key] !== -1) neededCols.push({key: key, idx: Q_COLS[key]});
  }
  var colIndices = neededCols.map(function(c) { return c.idx; });
  var maxCol = Math.max.apply(null, colIndices);

  var rawRange = sheet.getRange(startRow, 1, endRow - startRow + 1, maxCol + 1).getValues();
  var data = [];
  var caseIdIdx = Q_COLS.CASE_ID;

  for (var j = 0; j < rawRange.length; j++) {
    var row = rawRange[j];
    var rowIndex = startRow + j;
    if (relevantMap[rowIndex] && row[caseIdIdx]) {
      var obj = {};
      neededCols.forEach(function(c) {
        obj[c.idx] = row[c.idx];
      });
      data.push(obj);
    }
  }

  // 3. Cache the result (if not too large)
  try {
    var json = JSON.stringify(data);
    if (json.length < 100000) {
      cache.put(cacheKey, json, 3600);
    }
  } catch(e) {}

  Logger.log('getRawQualityDataForMonth: Sheet READ (' + data.length + ' rows) for ' + month + ' in ' + (new Date().getTime() - start) + 'ms');
  return data;
}

// Keep for compatibility or multi-month views
function getRawQualityData(forceRefresh) {
  // If we really need all data, we fetch it month by month or just do a full read
  // For 11k rows, a full read is okay IF cached.
  // But let's stick to month-based for the dashboard.
  return []; // Should not be called directly by the optimized dashboard
}

var _TIMEZONE = null;
function getTz() {
  if (!_TIMEZONE) _TIMEZONE = Session.getScriptTimeZone();
  return _TIMEZONE;
}

function getAvailableQualityMonths() {
  getColMapping();
  var rows = getColumnsFromSheet([Q_COLS.REVIEW_MONTH]);
  var seen = {};
  var tz = getTz();
  for (var i = 0; i < rows.length; i++) {
    var month = rows[i][0];
    if (month) {
      if (month instanceof Date) {
        try {
          month = Utilities.formatDate(month, tz, 'yyyy-MM');
        } catch(e) {
          month = month.getFullYear() + '-' + ('0' + (month.getMonth() + 1)).slice(-2);
        }
      }
      seen[month] = true;
    }
  }
  return Object.keys(seen).sort().reverse();
}

function normalizeQualityMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, getTz(), 'yyyy-MM');
    } catch(e) {
      return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2);
    }
  }
  var s = String(val).trim();
  // Handle ISO string like "2026-03-31T16:00:00.000Z"
  if (s.length >= 7 && s.indexOf('T') !== -1) {
    var d = new Date(s);
    try {
      return Utilities.formatDate(d, getTz(), 'yyyy-MM');
    } catch(e) {}
  }
  return s.substring(0, 7);
}

function normalizeLdap(val) {
  if (!val) return '';
  return String(val).trim().toLowerCase().split('@')[0];
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, getTz(), 'yyyy-MM-dd');
    } catch(e) {
      var d = val.getDate();
      var m = val.getMonth() + 1;
      var y = val.getFullYear();
      return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
    }
  }
  return String(val);
}

// ── AGGREGATION ───────────────────────────────────────────────────────────

function parseSheetScore(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  var s = String(val).toLowerCase().trim();
  if (s === 'yes' || s === 'pass' || s === '1') return 1;
  if (s === 'no' || s === 'fail' || s === '0') return 0;
  var n = parseFloat(s.replace('%', ''));
  if (!isNaN(n)) return n > 1 ? n / 100 : n;
  return 0;
}

function aggregateQualityRows(rows) {
  if (!rows || rows.length === 0) return null;

  var customerSum = 0, businessSum = 0, complianceSum = 0;
  var params = {};

  for (var i = 0; i < Q_PARAM_COLS.length; i++) {
    params[Q_PARAM_COLS[i]] = { yes: 0, total: 0 };
  }

  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    customerSum += parseSheetScore(r[Q_COLS.CUSTOMER_CRITICAL]);
    businessSum += parseSheetScore(r[Q_COLS.BUSINESS_CRITICAL]);
    complianceSum += parseSheetScore(r[Q_COLS.COMPLIANCE_CRITICAL]);

    for (var k = 0; k < Q_PARAM_COLS.length; k++) {
      var p = Q_PARAM_COLS[k];
      var val = String(r[Q_COLS[p]]).trim().toLowerCase();
      if (val === 'yes' || val === 'no' || val === '1' || val === '0') {
        params[p].total++;
        if (val === 'yes' || val === '1') params[p].yes++;
      }
    }
  }

  var count = rows.length;
  var paramScores = {};
  for (var l = 0; l < Q_PARAM_COLS.length; l++) {
    var pCol = Q_PARAM_COLS[l];
    paramScores[pCol] = params[pCol].total > 0 ? (params[pCol].yes / params[pCol].total) * 100 : null;
  }

  var groupedParams = {
    customer: Q_PARAM_GROUPS.customer.map(function(p) { return { name: p, score: paramScores[p] }; }),
    business: Q_PARAM_GROUPS.business.map(function(p) { return { name: p, score: paramScores[p] }; }),
    compliance: Q_PARAM_GROUPS.compliance.map(function(p) { return { name: p, score: paramScores[p] }; })
  };

  return {
    customer: (customerSum / count) * 100,
    business: (businessSum / count) * 100,
    compliance: (complianceSum / count) * 100,
    count: count,
    params: paramScores,
    groupedParams: groupedParams,
    targets: Q_TARGETS
  };
}

function aggregateTrends(rows) {
  var daily = {};
  var weekly = {};

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var dateRaw = r[Q_COLS.REVIEW_DATE];
    var date = formatDate(dateRaw);
    var week = formatDate(r[Q_COLS.REVIEW_WEEK]);

    var trends = [ {obj: daily, key: date}, {obj: weekly, key: week} ];
    for (var j = 0; j < trends.length; j++) {
      var t = trends[j];
      if (!t.key) continue;
      if (!t.obj[t.key]) t.obj[t.key] = { customer: 0, business: 0, compliance: 0, count: 0 };
      t.obj[t.key].customer += parseSheetScore(r[Q_COLS.CUSTOMER_CRITICAL]);
      t.obj[t.key].business += parseSheetScore(r[Q_COLS.BUSINESS_CRITICAL]);
      t.obj[t.key].compliance += parseSheetScore(r[Q_COLS.COMPLIANCE_CRITICAL]);
      t.obj[t.key].count++;
    }
  }

  var formatTrend = function(obj) {
    return Object.keys(obj).sort().map(function(k) {
      var d = obj[k];
      return {
        label: k,
        customer: (d.customer / d.count) * 100,
        business: (d.business / d.count) * 100,
        compliance: (d.compliance / d.count) * 100,
        avg: ((d.customer + d.business + d.compliance) / (d.count * 3)) * 100
      };
    });
  };

  return { daily: formatTrend(daily), weekly: formatTrend(weekly) };
}

// ── CLIENT WRAPPERS ───────────────────────────────────────────────────────

function clientGetAvailableQualityMonths(forceRefresh) {
  if (forceRefresh) {
    _MEMOIZED_RAW_DATA = null;
    var cache = CacheService.getScriptCache();
    cache.remove('quality_raw_v3_chunks');
    cache.remove('quality_hierarchy_v1_chunks');
  }
  return getAvailableQualityMonths();
}

function clientGetMyQuality(ldap, month, forceRefresh) {
  if (!ldap) ldap = Session.getActiveUser().getEmail().split('@')[0];

  var cache = CacheService.getScriptCache();
  var cacheKey = 'q_agent_' + normalizeLdap(ldap) + '_' + month + '_' + CACHE_VERSION;

  if (!forceRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  var allRows = getRawQualityDataForMonth(month, forceRefresh);
  var filtered = [];
  for (var i = 0; i < allRows.length; i++) {
    var r = allRows[i];
    if (normalizeLdap(r[Q_COLS.AGENT_LDAP]) === normalizeLdap(ldap)) {
      filtered.push(r);
    }
  }

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  // Calculate Team Average for benchmarking
  var teamAvg = { customer: 0, business: 0, compliance: 0, hasData: false };
  if (filtered.length > 0) {
    var supervisor = String(filtered[0][Q_COLS.SUPERVISOR]).trim();
    var teamRows = [];
    for (var ti = 0; ti < allRows.length; ti++) {
        var tr = allRows[ti];
        if (String(tr[Q_COLS.SUPERVISOR]).trim() === supervisor &&
            normalizeQualityMonth(tr[Q_COLS.REVIEW_MONTH]) === month) {
            teamRows.push(tr);
        }
    }
    if (teamRows.length > 0) {
        var teamStats = aggregateQualityRows(teamRows);
        teamAvg.customer = teamStats.customer;
        teamAvg.business = teamStats.business;
        teamAvg.compliance = teamStats.compliance;
        teamAvg.hasData = true;
    }
  }

  var caseLog = filtered.map(function(r) {
    var cId = r[Q_COLS.CASE_ID];
    return {
      caseId: cId ? String(cId).trim() : '',
      reviewDate: formatDate(r[Q_COLS.REVIEW_DATE]),
      customer: r[Q_COLS.CUSTOMER_CRITICAL],
      business: r[Q_COLS.BUSINESS_CRITICAL],
      compliance: r[Q_COLS.COMPLIANCE_CRITICAL],
      comments: r[Q_COLS.REVIEWER_COMMENTS],
      details: {
        customer: {
          LISTENING: r[Q_COLS.LISTENING],
          PROBING: r[Q_COLS.PROBING],
          COMPLETE_RESOLUTION: r[Q_COLS.COMPLETE_RESOLUTION],
          TROUBLESHOOTING: r[Q_COLS.TROUBLESHOOTING],
          USER_EXPECTATIONS: r[Q_COLS.USER_EXPECTATIONS],
          EMPATHY: r[Q_COLS.EMPATHY],
          OWNERSHIP: r[Q_COLS.OWNERSHIP],
          REFUNDS: r[Q_COLS.REFUNDS],
          RESPONSIVENESS: r[Q_COLS.RESPONSIVENESS]
        },
        business: {
          CONSULTS_ESCALATIONS: r[Q_COLS.CONSULTS_ESCALATIONS],
          CASE_DETAILS: r[Q_COLS.CASE_DETAILS],
          CATEGORIZATION: r[Q_COLS.CATEGORIZATION],
          CSAT_REMINDER: r[Q_COLS.CSAT_REMINDER],
          CASE_STATE: r[Q_COLS.CASE_STATE],
          OPENING_CLOSING: r[Q_COLS.OPENING_CLOSING],
          LANGUAGE_PROFICIENCY: r[Q_COLS.LANGUAGE_PROFICIENCY]
        },
        compliance: {
          AUTHENTICATION: r[Q_COLS.AUTHENTICATION],
          GOOGLE_ONLY_INFO: r[Q_COLS.GOOGLE_ONLY_INFO],
          PROFESSIONAL_CONDUCT: r[Q_COLS.PROFESSIONAL_CONDUCT],
          PAYMENT_COMPLAINTS: r[Q_COLS.PAYMENT_COMPLAINTS]
        }
      }
    };
  });

  var metadata = { lob: '', supervisor: '', team: '', manager: '' };
  if (filtered.length > 0) {
    var r0 = filtered[0];
    metadata.lob = String(r0[Q_COLS.LOB] || '').trim();
    metadata.supervisor = String(r0[Q_COLS.SUPERVISOR] || '').trim();
    metadata.team = String(r0[Q_COLS.TEAM] || '').trim();
    metadata.manager = String(r0[Q_COLS.MANAGER] || '').trim();
  }

  var result = {
    ldap: ldap,
    month: month,
    stats: stats,
    trends: trends,
    caseLog: caseLog,
    metadata: metadata,
    teamAvg: teamAvg,
    hasData: filtered.length > 0
  };

  try { cache.put(cacheKey, JSON.stringify(result), 3600); } catch(e) {} // 6 hours
  return result;
}

function clientGetTeamQuality(supervisor, month, forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'q_team_' + supervisor.replace(/\s/g, '_') + '_' + month + '_' + CACHE_VERSION;

  if (!forceRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  var allRows = getRawQualityDataForMonth(month, forceRefresh);
  var filtered = [];
  for (var i = 0; i < allRows.length; i++) {
    var r = allRows[i];
    if (String(r[Q_COLS.SUPERVISOR]).trim() === String(supervisor).trim()) {
      filtered.push(r);
    }
  }

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var agentStats = {};
  var uniqueLdaps = [];
  for (var j = 0; j < filtered.length; j++) {
    var row = filtered[j];
    var ldap = normalizeLdap(row[Q_COLS.AGENT_LDAP]);
    if (!agentStats[ldap]) {
      agentStats[ldap] = [];
      uniqueLdaps.push(ldap);
    }
    agentStats[ldap].push(row);
  }

  var agents = uniqueLdaps.map(function(ldap) {
    return {
      ldap: ldap,
      stats: aggregateQualityRows(agentStats[ldap])
    };
  }).sort(function(a, b) {
    return (b.stats.customer + b.stats.business + b.stats.compliance) - (a.stats.customer + a.stats.business + a.stats.compliance);
  });

  var result = {
    supervisor: supervisor,
    month: month,
    stats: stats,
    trends: trends,
    agents: agents,
    hasData: filtered.length > 0
  };

  try { cache.put(cacheKey, JSON.stringify(result), 3600); } catch(e) {} // 6 hours
  return result;
}

function clientGetClusterQuality(manager, month, forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'q_cluster_' + manager.replace(/\s/g, '_') + '_' + month + '_' + CACHE_VERSION;

  if (!forceRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  var allRows = getRawQualityDataForMonth(month, forceRefresh);
  var filtered = [];
  for (var i = 0; i < allRows.length; i++) {
    var r = allRows[i];
    if (String(r[Q_COLS.MANAGER]).trim() === String(manager).trim()) {
      filtered.push(r);
    }
  }

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var supervisorStats = {};
  var uniqueSupervisors = [];
  for (var j = 0; j < filtered.length; j++) {
    var row = filtered[j];
    var sup = String(row[Q_COLS.SUPERVISOR]).trim();
    if (!supervisorStats[sup]) {
      supervisorStats[sup] = [];
      uniqueSupervisors.push(sup);
    }
    supervisorStats[sup].push(row);
  }

  var supervisors = uniqueSupervisors.map(function(sup) {
    return {
      name: sup,
      stats: aggregateQualityRows(supervisorStats[sup])
    };
  }).sort(function(a, b) {
    return (b.stats.customer + b.stats.business + b.stats.compliance) - (a.stats.customer + a.stats.business + a.stats.compliance);
  });

  var result = {
    manager: manager,
    month: month,
    stats: stats,
    trends: trends,
    supervisors: supervisors,
    hasData: filtered.length > 0
  };

  try { cache.put(cacheKey, JSON.stringify(result), 3600); } catch(e) {} // 6 hours
  return result;
}

function clientGetAllAgents() {
  getColMapping();
  var rows = getColumnsFromSheet([Q_COLS.AGENT_LDAP]);
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var ldap = normalizeLdap(rows[i][0]);
    if (ldap) seen[ldap] = true;
  }
  return Object.keys(seen).sort().map(function(ldap) {
    return { ldap: ldap };
  });
}

function clientGetAllSupervisors() {
  getColMapping();
  var rows = getColumnsFromSheet([Q_COLS.SUPERVISOR]);
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var sup = String(rows[i][0]).trim();
    if (sup) seen[sup] = true;
  }
  return Object.keys(seen).sort();
}

function clientGetAllManagers() {
  getColMapping();
  var rows = getColumnsFromSheet([Q_COLS.MANAGER]);
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var mgr = String(rows[i][0]).trim();
    if (mgr) seen[mgr] = true;
  }
  return Object.keys(seen).sort();
}

function clientGetSession() {
  var email = Session.getActiveUser().getEmail();
  return {
    ldap: email.split('@')[0],
    email: email
  };
}

function clientGetInitialData(forceRefresh) {
  var session = clientGetSession();
  var hierarchy = clientGetHierarchy(forceRefresh);
  var months = getAvailableQualityMonths();

  var preloadedData = null;
  try {
    if (months && months.length > 0) {
      preloadedData = clientGetMyQuality(session.ldap, months[0], false);
    }
  } catch(e) {
    Logger.log('Preload failed: ' + e.message);
  }

  return {
    session: session,
    targets: Q_TARGETS,
    cols: getColMapping(),
    paramGroups: Q_PARAM_GROUPS,
    paramCols: Q_PARAM_COLS,
    cacheVersion: CACHE_VERSION,
    hierarchy: hierarchy.tree,
    managers: hierarchy.managers,
    months: months,
    preloadedData: preloadedData
  };
}

/**
 * Optimized Hierarchy fetching.
 * Caches the small result set instead of the whole raw data.
 */
function clientGetHierarchy(forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'quality_hierarchy_' + CACHE_VERSION;

  if (!forceRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  getColMapping();
  // Fetch only the columns needed for hierarchy
  var indices = [Q_COLS.LOB, Q_COLS.SUPERVISOR, Q_COLS.AGENT_LDAP, Q_COLS.MANAGER];
  var rows = getColumnsFromSheet(indices);

  var hierarchy = {};
  var managers = {};

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var lob = String(r[Q_COLS.LOB] || 'Unknown LOB').trim();
    var sup = String(r[Q_COLS.SUPERVISOR] || 'Unknown Supervisor').trim();
    var agent = normalizeLdap(r[Q_COLS.AGENT_LDAP]);
    var mgr = String(r[Q_COLS.MANAGER] || '').trim();

    if (agent) {
      if (!hierarchy[lob]) hierarchy[lob] = {};
      if (!hierarchy[lob][sup]) hierarchy[lob][sup] = {};
      hierarchy[lob][sup][agent] = true;
    }
    if (mgr) managers[mgr] = true;
  }

  // Format Hierarchy
  var result = {
    tree: {},
    managers: Object.keys(managers).sort()
  };

  var lobs = Object.keys(hierarchy).sort();
  for (var j = 0; j < lobs.length; j++) {
    var l = lobs[j];
    result.tree[l] = {};
    var sups = Object.keys(hierarchy[l]).sort();
    for (var k = 0; k < sups.length; k++) {
      var s = sups[k];
      result.tree[l][s] = Object.keys(hierarchy[l][s]).sort();
    }
  }

  try {
    cache.put(cacheKey, JSON.stringify(result), 3600);
  } catch(e) {}

  return result;
}

function clearAllCache() {
  var cache = CacheService.getScriptCache();
  cache.remove('quality_hierarchy_' + CACHE_VERSION);

  // We can't easily clear all month-based caches without knowing the keys,
  // but changing CACHE_VERSION will effectively invalidate them.

  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('raw_data_hash_' + CACHE_VERSION + '_fp');
  props.deleteProperty('raw_data_hash_' + CACHE_VERSION + '_chunks');
  _MEMOIZED_RAW_DATA = null;
  Q_COLS = null;
  Logger.log('All caches cleared. (Note: Month-based caches rely on CACHE_VERSION increment for full invalidation)');
}

function prewarmAllCaches() {
  var start = new Date().getTime();
  Logger.log('Cache prewarm started...');

  var months = getAvailableQualityMonths();
  if (!months || months.length === 0) return;
  var latestMonth = months[0];

  // Prewarm hierarchy (used by everyone on load)
  clientGetHierarchy(true);

  // Prewarm all agent caches for the latest month
  getColMapping();
  var agentRows = getColumnsFromSheet([Q_COLS.AGENT_LDAP, Q_COLS.SUPERVISOR, Q_COLS.MANAGER]);
  var seenAgents = {};
  var seenSups = {};
  var seenMgrs = {};

  for (var i = 0; i < agentRows.length; i++) {
    var ldap = normalizeLdap(agentRows[i][0]);
    var sup  = String(agentRows[i][1] || '').trim();
    var mgr  = String(agentRows[i][2] || '').trim();
    if (ldap) seenAgents[ldap] = true;
    if (sup)  seenSups[sup]    = true;
    if (mgr)  seenMgrs[mgr]   = true;
  }

  // Prewarm per-agent
  var agents = Object.keys(seenAgents);
  Logger.log('Prewarming ' + agents.length + ' agents...');
  for (var a = 0; a < agents.length; a++) {
    try {
      clientGetMyQuality(agents[a], latestMonth, true);
    } catch(e) {
      Logger.log('Agent prewarm failed: ' + agents[a] + ' — ' + e.message);
    }
    if (a % 10 === 0) Utilities.sleep(200); // avoid quota bursts
  }

  // Prewarm per-supervisor
  var sups = Object.keys(seenSups);
  Logger.log('Prewarming ' + sups.length + ' supervisors...');
  for (var s = 0; s < sups.length; s++) {
    try {
      clientGetTeamQuality(sups[s], latestMonth, true);
    } catch(e) {
      Logger.log('Sup prewarm failed: ' + sups[s] + ' — ' + e.message);
    }
    Utilities.sleep(100);
  }

  // Prewarm per-manager
  var mgrs = Object.keys(seenMgrs);
  Logger.log('Prewarming ' + mgrs.length + ' managers...');
  for (var m = 0; m < mgrs.length; m++) {
    try {
      clientGetClusterQuality(mgrs[m], latestMonth, true);
    } catch(e) {
      Logger.log('Mgr prewarm failed: ' + mgrs[m] + ' — ' + e.message);
    }
    Utilities.sleep(100);
  }

  Logger.log('Cache prewarm done in ' + (new Date().getTime() - start) + 'ms');
}

function installPrewarmTrigger() {
  // Delete any existing prewarm triggers first
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'prewarmAllCaches') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Install new trigger every 30 minutes
  ScriptApp.newTrigger('prewarmAllCaches')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('Prewarm trigger installed.');
}

function testNewSheet() {
  try {
    var ss = SpreadsheetApp.openById('1YDz16oRc2yi3sjyxtRPmaJbljyRPfYmwxVzuqiXu4Vw');
    var sheet = ss.getSheetByName('PLX Raw data');
    if (!sheet) {
      Logger.log('ERROR: Sheet tab "PLX Raw data" not found.');
      Logger.log('Available tabs: ' + ss.getSheets().map(s => s.getName()).join(', '));
    } else {
      Logger.log('SUCCESS: Found sheet with ' + sheet.getLastRow() + ' rows.');
    }
  } catch(e) {
    Logger.log('EXCEPTION: ' + e.message);
  }
}
function debugHeaders() {
  var ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(QUALITY_SHEET_NAME);
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log(JSON.stringify(headers));
}

function debugMyQuality() {
  var months = getAvailableQualityMonths();
  Logger.log('Months: ' + JSON.stringify(months));
  
  var result = clientGetMyQuality('stevenjosephc', months[0], true);
  Logger.log('hasData: ' + result.hasData);
  Logger.log('rowCount: ' + (result.caseLog ? result.caseLog.length : 0));
  Logger.log('stats: ' + JSON.stringify(result.stats));
  Logger.log('metadata: ' + JSON.stringify(result.metadata));
}

function debugLdapMatch() {
  getColMapping();
  var rows = getColumnsFromSheet([Q_COLS.AGENT_LDAP]);
  // Print first 5 non-empty values
  var count = 0;
  for (var i = 0; i < rows.length; i++) {
    var val = rows[i][Q_COLS.AGENT_LDAP];
    if (val) {
      Logger.log('Row ' + i + ': [' + val + ']');
      if (++count >= 5) break;
    }
  }
}

function debugAgents() {
  var agents = clientGetAllAgents();
  Logger.log('Count: ' + agents.length);
  Logger.log('First 5: ' + JSON.stringify(agents.slice(0, 5)));
}

function debugAgents2() {
  var agents = clientGetAllAgents();
  Logger.log('Count: ' + agents.length);
  Logger.log('First 5: ' + JSON.stringify(agents.slice(0, 5)));
}

function debugMonths() {
  getColMapping();
  Logger.log('REVIEW_MONTH col index: ' + Q_COLS.REVIEW_MONTH);
  var rows = getColumnsFromSheet([Q_COLS.REVIEW_MONTH]);
  Logger.log('Total rows: ' + rows.length);
  Logger.log('First 5 raw: ' + JSON.stringify(rows.slice(0, 5)));
}

function debugFindAgent() {
  getColMapping();
  var allRows = getRawQualityDataForMonth('2026-04', true);
  Logger.log('Total rows for 2026-04: ' + allRows.length);
  var found = allRows.filter(function(r) {
    return String(r[Q_COLS.AGENT_LDAP]).toLowerCase().indexOf('steven') !== -1;
  });
  Logger.log('Matching rows: ' + found.length);
  if (found.length > 0) Logger.log('Sample: ' + JSON.stringify(found[0]));
}
