/**
 * Interactive Turkey provinces map (RTL)
 * - Colors provinces based on region data
 * - Shows ONLY manager in left panel on click
 * - Adds Arabic city labels for colored provinces
 *
 * Data: ./data/regions.json
 * Map:  ./data/turkey-provinces.geo.json
 */

const PATHS = {
  regions: "./data/regions.json",
  geo: "./data/turkey-provinces.geo.json",
};

const COLOR_MAP = {
  red: "#e11d48",
  yellow: "#f59e0b",
  blue: "#2563eb",
  orange: "#f97316",
  green: "#16a34a",
};

// Apps Script Web App endpoint (replace if needed)
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx_IjQNg_aqq6TGesSRdKVGwlg6sbCnb2zYCg2PsHEzbROe0Ayru8knVww4_CHMLc4/exec";
const PRESIDENTS_API_URL = "https://script.google.com/macros/s/AKfycbyiXHAcF2ZBJgDRPSXP4m326S_v3jK0BbBFYc8bwAq0M4NOvT_lWCCtOa9Md8TIAdFB/exec";
// Quick test 1 (ping): افتح الرابط التالي في المتصفح وتحقق أن ok:true
// `${PRESIDENTS_API_URL}?action=ping`
// Quick test 2 (add president): أرسل النموذج وتحقق من إضافة صف جديد في Sheet3
// Offices API deployment checklist:
// 1) Deploy -> New deployment -> Web app
// 2) Execute as: Me
// 3) Who has access: Anyone (or Anyone with Google account)
// 4) Use /exec URL only (not /dev)
const OFFICES_API_URL = "https://script.google.com/macros/s/AKfycbw0DBNAdU6T_eiZ2pNFExa0-v7Szp7E-o7aewYAzmvwxoyTwSYXHzeISsueN9Zwk7S4/exec";
const ADMIN_SESSION_KEY = 'isAdminLoggedIn';
const CITY_DATA_CACHE_TTL_MS = 5 * 60 * 1000;
const ISTANBUL_CITY = 'إسطنبول';
const ISTANBUL_UNIONS = [
  'اتحاد طلبة في جامعة جليشيم',
  'اتحاد طلبة في جامعة نيشانتشي',
  'اتحاد طلبة في جامعة إيستينيا',
];
const UNION_MARKER_PREFIX = '[UNION]:';

// Events storage by Arabic city name -> array of events
const eventsByCity = new Map();
const officesByCity = new Map();
const officesCacheMeta = new Map();
const presidentByCity = new Map();
const presidentsCacheMeta = new Map();
const officeRequestsInFlight = new Map();
const presidentRequestsInFlight = new Map();
let presidentsReadActionUnsupported = false;
const CITY_WARMUP_CONCURRENCY = 4;
const BULK_DATA_TTL_MS = 5 * 60 * 1000;
let isAdminLoggedIn = sessionStorage.getItem(ADMIN_SESSION_KEY) === 'true';
const selectedUnionByCity = new Map();
const bulkOfficesState = { rows: [], fetchedAt: 0, promise: null };
const bulkPresidentsState = { rows: [], fetchedAt: 0, promise: null };

function isBulkStateFresh(state) {
  return !!state && !!state.fetchedAt && (Date.now() - state.fetchedAt) < BULK_DATA_TTL_MS;
}

function getCachedValue(cacheMap, metaMap, cityKey) {
  const meta = metaMap.get(cityKey);
  if (!meta || meta.expiresAt <= Date.now()) return undefined;
  return cacheMap.get(cityKey);
}

function setCachedValue(cacheMap, metaMap, cityKey, value) {
  cacheMap.set(cityKey, value);
  metaMap.set(cityKey, { expiresAt: Date.now() + CITY_DATA_CACHE_TTL_MS });
}

function invalidateCityCache(cacheMap, metaMap, inFlightMap, cityName) {
  const cityKey = normalizeCityToArabic(cityName);
  if (!cityKey) return;
  cacheMap.delete(cityKey);
  metaMap.delete(cityKey);
  inFlightMap.delete(cityKey);
}

function invalidatePresidentCacheForCity(cityName, unionName = '') {
  const cityKey = normalizeCityToArabic(cityName);
  if (!cityKey) return;

  const selectedUnion = String(unionName || getSelectedUnionForCity(cityKey) || '').trim();
  const scopeKey = getScopeKey(cityKey, selectedUnion);
  const effectiveCity = cityKey === ISTANBUL_CITY && selectedUnion
    ? `${selectedUnion} - ${ISTANBUL_CITY}`
    : cityKey;

  const keysToClear = [cityKey, scopeKey, effectiveCity];
  keysToClear.forEach((key) => {
    if (!key) return;
    presidentByCity.delete(key);
    presidentsCacheMeta.delete(key);
    presidentRequestsInFlight.delete(key);
  });
}

// Extra safety: ensure modal/fab/admin are hidden on DOM ready
document.addEventListener('DOMContentLoaded', () => {
  const addModal = document.getElementById('addModal'); if (addModal) addModal.hidden = true;
  const presidentModal = document.getElementById('presidentModal'); if (presidentModal) presidentModal.hidden = true;
  const officeModal = document.getElementById('officeModal'); if (officeModal) officeModal.hidden = true;
  const eventsContainer = document.getElementById('events'); if (eventsContainer) eventsContainer.hidden = true;
  const adminSectionInit = document.getElementById('adminSection'); if (adminSectionInit) adminSectionInit.hidden = true;

  const legendSection = document.getElementById('legendSection');
  const legendWrapper = document.getElementById('legendWrapper');
  if (legendSection && legendWrapper) {
    legendSection.appendChild(legendWrapper);
  }
});

// Arabic -> Turkish province names (canonical)
const AR_TO_TR = {
  "إسطنبول": "İstanbul",
  "كوجالي": "Kocaeli",
  "بورصة": "Bursa",
  "سكاريا": "Sakarya",
  "دوزجة": "Düzce",

  "إزمير": "İzmir",
  "كوتاهيا": "Kütahya",
  "إسبارطة": "Isparta",
  "دنيزلي": "Denizli",
  "قونية": "Konya",
  "أوشاك": "Uşak",

  "بولو": "Bolu",
  "أنقرة": "Ankara",
  "كارابوك": "Karabük",
  "سامسون": "Samsun",
  "كاستامونو": "Kastamonu",

  "توكات": "Tokat",
  "قيصري": "Kayseri",
  "سيواس": "Sivas",
  "إيلازيغ": "Elazığ",
  "ملاطيا": "Malatya",

  "مرسين": "Mersin",
  "أضنة": "Adana",
  "هاتاي": "Hatay",
  "كيليس": "Kilis",
  "غازي عنتاب": "Gaziantep",
  "شانلي أورفا": "Şanlıurfa",
  "كهرمان مرعش": "Kahramanmaraş",
  "أديامان": "Adıyaman",
};

// Reverse map TR -> AR for labels
const TR_TO_AR = Object.fromEntries(Object.entries(AR_TO_TR).map(([ar, tr]) => [tr, ar]));
const AR_NORMALIZED_TO_AR = new Map(Object.keys(AR_TO_TR).map((ar) => [normalizeKey(ar), ar]));
const TR_NORMALIZED_TO_AR = new Map(Object.entries(TR_TO_AR).map(([tr, ar]) => [normalizeTR(tr), ar]));

// Normalize helper (handles Turkish chars + case + spaces)
function normalizeKey(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function normalizeTR(s) {
  // convert Turkish-specific letters to ASCII-ish for robust matching
  // Handle Turkish İ first (before toLowerCase)
  s = String(s || "")
    .replace(/İ/g, "i")  // Turkish capital İ -> i
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
  
  // Now handle lowercase Turkish characters
  s = s
    .replace(/ş/g, "s")
    .replace(/ğ/g, "g")
    .replace(/ü/g, "u")
    .replace(/ö/g, "o")
    .replace(/ç/g, "c")
    .replace(/ı/g, "i")
    .replace(/â/g, "a")
    .replace(/î/g, "i")
    .replace(/û/g, "u");
  
  return s;
}

function safe(v) {
  const text = String(v ?? '').trim();
  return text ? text : '—';
}

function normalizeCityToArabic(cityName) {
  const raw = String(cityName || '').trim();
  if (!raw) return '';

  if (raw.includes(ISTANBUL_CITY)) return ISTANBUL_CITY;

  if (AR_TO_TR[raw]) return raw;

  const arFromNormalized = AR_NORMALIZED_TO_AR.get(normalizeKey(raw));
  if (arFromNormalized) return arFromNormalized;

  if (TR_TO_AR[raw]) return TR_TO_AR[raw];

  const trFromNormalized = TR_NORMALIZED_TO_AR.get(normalizeTR(raw));
  if (trFromNormalized) return trFromNormalized;

  return raw;
}

function getEventsForCity(cityName) {
  const cityKey = normalizeCityToArabic(cityName);
  if (!cityKey) return [];

  if (eventsByCity.has(cityKey)) return eventsByCity.get(cityKey) || [];

  for (const [storedCity, storedEvents] of eventsByCity.entries()) {
    if (normalizeCityToArabic(storedCity) === cityKey) {
      return storedEvents || [];
    }
  }

  return [];
}

function getOfficesForCity(cityName) {
  const cityKey = normalizeCityToArabic(cityName);
  if (!cityKey) return [];

  if (officesByCity.has(cityKey)) return officesByCity.get(cityKey) || [];

  for (const [storedCity, storedOffices] of officesByCity.entries()) {
    if (normalizeCityToArabic(storedCity) === cityKey) {
      return storedOffices || [];
    }
  }

  return [];
}

function pickOfficeValue(office, keys) {
  for (const key of keys) {
    const value = office && office[key];
    const text = String(value ?? '').trim();
    if (text) return text;
  }
  return '';
}

function pickProvinceNameFromProps(props) {
  // try common keys across different Turkey GeoJSONs
  const candidates = [
    props?.name,
    props?.NAME,
    props?.Name,
    props?.il_adi,
    props?.province,
    props?.admin,
    props?.id,        // sometimes contains name
    props?.City,      // rare
    props?.city,      // rare
  ].filter(Boolean);

  return candidates[0] ? String(candidates[0]).trim() : "";
}

async function loadJSON(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to load ${url}: ${res.status}`);
  return res.json();
}

function buildLegend(regions) {
  const legend = document.getElementById("legendItems");
  legend.innerHTML = "";

  for (const r of regions) {
    const item = document.createElement("div");
    item.className = "legendItem";

    const badge = document.createElement("div");
    badge.className = "badge";
    badge.style.background = COLOR_MAP[r.color] || r.color;

    const text = document.createElement("div");
    text.className = "legendText";
    text.textContent = r.regionName;

    const mgr = document.createElement("div");
    mgr.className = "legendMgr";
    mgr.textContent = r.manager;

    const left = document.createElement("div");
    left.style.display = "flex";
    left.style.alignItems = "center";
    left.style.gap = "10px";
    left.appendChild(badge);
    left.appendChild(text);

    item.appendChild(left);
    item.appendChild(mgr);

    legend.appendChild(item);
  }
}

function setPanel(manager) {
  const hint = document.getElementById("hint");
  const card = document.getElementById("infoCard");
  const name = document.getElementById("managerName");

  if (!manager) {
    hint.hidden = false;
    card.hidden = true;
    name.textContent = "—";
    return;
  }
  hint.hidden = true;
  card.hidden = false;
  name.textContent = manager;
}

function normalizeArabicLoose(value) {
  return String(value || '')
    .replace(/[\u0640]/g, '')
    .replace(/[\u0622\u0623\u0625]/g, 'ا')
    .replace(/[\u0649]/g, 'ي')
    .replace(/[\u0629]/g, 'ه')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function getScopeKey(cityName, unionName = '') {
  const city = normalizeCityToArabic(cityName);
  if (!city) return '';
  if (city !== ISTANBUL_CITY) return city;
  const union = String(unionName || '').trim();
  return union ? `${city}::${normalizeArabicLoose(union)}` : city;
}

function getUnionFromRecord(row) {
  const candidates = [
    row?.union_name,
    row?.Union_Name,
    row?.union,
    row?.Union,
    row?.university,
    row?.University,
    row?.university_name,
    row?.University_Name,
    row?.uni,
    row?.Uni,
  ];
  for (const candidate of candidates) {
    const text = String(candidate || '').trim();
    if (text) return text;
  }

  const notesCandidates = [
    row?.notes,
    row?.Notes,
    row?.notes_2,
    row?.Notes2,
    row?.description,
    row?.Description,
  ];

  for (const noteValue of notesCandidates) {
    const noteText = String(noteValue || '').trim();
    if (!noteText) continue;
    const regex = /\[UNION\]:\s*(.+)$/im;
    const match = noteText.match(regex);
    if (match && match[1]) {
      return String(match[1]).trim();
    }
  }

  return '';
}

function appendUnionMarkerToNotes(notesText, unionName) {
  const notes = String(notesText || '').trim();
  const union = String(unionName || '').trim();
  if (!union) return notes;

  const withoutMarker = notes.replace(/\n?\[UNION\]:\s*.+$/im, '').trim();
  return withoutMarker
    ? `${withoutMarker}\n${UNION_MARKER_PREFIX} ${union}`
    : `${UNION_MARKER_PREFIX} ${union}`;
}

function stripUnionMarker(textValue) {
  const text = String(textValue || '').trim();
  if (!text) return '';
  return text.replace(/\n?\[UNION\]:\s*.+$/im, '').trim();
}

function getSelectedUnionForCity(cityName) {
  const city = normalizeCityToArabic(cityName);
  if (city !== ISTANBUL_CITY) return '';
  return String(selectedUnionByCity.get(city) || '').trim();
}

function getEffectiveCityName(cityName) {
  const city = normalizeCityToArabic(cityName);
  if (!city) return '';
  if (city !== ISTANBUL_CITY) return city;
  const selectedUnion = getSelectedUnionForCity(city);
  return selectedUnion ? `${selectedUnion} - ${ISTANBUL_CITY}` : city;
}

function getCityForApi(cityName, unionName = '') {
  const city = normalizeCityToArabic(cityName);
  if (!city) return '';
  const union = String(unionName || '').trim();
  if (city === ISTANBUL_CITY && union) {
    return `${union} - ${ISTANBUL_CITY}`;
  }
  return city;
}

function filterRowsBySelectedUnion(cityName, rows) {
  const city = normalizeCityToArabic(cityName);
  const list = Array.isArray(rows) ? rows : [];
  if (city !== ISTANBUL_CITY) return list;

  const selectedUnion = getSelectedUnionForCity(city);
  if (!selectedUnion) return list;
  const target = normalizeArabicLoose(selectedUnion);

  return list.filter((row) => normalizeArabicLoose(getUnionFromRecord(row)) === target);
}

function hydrateOfficesCacheFromRows(rows) {
  const grouped = new Map();
  const list = Array.isArray(rows) ? rows : [];

  list.forEach((office) => {
    const cityKey = normalizeCityToArabic(office?.city || office?.City || '');
    if (!cityKey) return;
    if (!grouped.has(cityKey)) grouped.set(cityKey, []);
    grouped.get(cityKey).push({ ...office, city: cityKey });
  });

  grouped.forEach((offices, cityKey) => {
    setCachedValue(officesByCity, officesCacheMeta, cityKey, offices);
  });
}

function hydratePresidentsCacheFromRows(rows) {
  const list = Array.isArray(rows) ? rows : [];

  list.forEach((row) => {
    const cityRaw = row?.city || row?.City || '';
    const cityKey = normalizeCityToArabic(cityRaw);
    if (!cityKey) return;

    const unionName = getUnionFromRecord(row);
    const scopeKey = getScopeKey(cityKey, unionName);
    const effectiveCity = getCityForApi(cityKey, unionName) || cityKey;

    setCachedValue(presidentByCity, presidentsCacheMeta, scopeKey, row);
    if (!presidentByCity.has(cityKey)) {
      setCachedValue(presidentByCity, presidentsCacheMeta, cityKey, row);
    }
    if (effectiveCity) {
      setCachedValue(presidentByCity, presidentsCacheMeta, effectiveCity, row);
    }
  });
}

function pickPresidentFromRows(rows, cityArabic, selectedUnion, effectiveCity) {
  const list = Array.isArray(rows) ? rows : [];
  if (!cityArabic || !list.length) return null;

  const cityLoose = normalizeArabicLoose(cityArabic);
  const effectiveLoose = normalizeArabicLoose(effectiveCity || cityArabic);

  let cityMatched = list.filter((row) => {
    const rowCity = row?.city || row?.City || '';
    const rowArabic = normalizeCityToArabic(rowCity);
    const rowLoose = normalizeArabicLoose(rowArabic || rowCity);
    if (cityArabic === ISTANBUL_CITY && selectedUnion) {
      return rowLoose === cityLoose || rowLoose === effectiveLoose;
    }
    return rowLoose === cityLoose;
  });

  if (cityArabic === ISTANBUL_CITY && selectedUnion) {
    const selectedUnionLoose = normalizeArabicLoose(selectedUnion);
    if (cityMatched.length) {
      const byUnion = cityMatched.filter((row) => normalizeArabicLoose(getUnionFromRecord(row)) === selectedUnionLoose);
      if (byUnion.length) cityMatched = byUnion;
    } else {
      const unionMatched = list.filter((row) => normalizeArabicLoose(getUnionFromRecord(row)) === selectedUnionLoose);
      if (unionMatched.length) cityMatched = unionMatched;
    }
  }

  return cityMatched.length ? cityMatched[0] : null;
}

async function queryPresidentsByCity(cityQuery, withAction = true) {
  const cityValue = String(cityQuery || '').trim();
  const cityParam = cityValue ? `&city=${encodeURIComponent(cityValue)}` : '';
  const url = withAction
    ? `${PRESIDENTS_API_URL}?action=presidents${cityParam}`
    : (cityValue
      ? `${PRESIDENTS_API_URL}?city=${encodeURIComponent(cityValue)}`
      : `${PRESIDENTS_API_URL}`);
  console.log('Presidents API request URL:', url);
  const res = await fetch(url, { method: 'GET', redirect: 'follow' });

  const contentType = (res.headers.get('content-type') || '').toLowerCase();
  const text = await res.text();
  console.log('Presidents API response content-type:', contentType);
  console.log('Presidents API raw response (first 200):', String(text || '').slice(0, 200));

  if (!res.ok) throw new Error(`PRESIDENTS_HTTP_${res.status}`);
  const trimmed = String(text || '').trim();
  if (contentType.includes('text/html') || /^<!doctype html|^<html/i.test(trimmed)) {
    throw new Error('PRESIDENTS_HTML_RESPONSE');
  }

  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error('PRESIDENTS_INVALID_JSON');
  }

  if (!json) {
    throw new Error('PRESIDENTS_INVALID_JSON_OBJECT');
  }

  if (json.ok !== true) {
    const apiError = String((json.error || json.message || '')).trim();
    if (/unknown action/i.test(apiError)) {
      console.warn('Presidents API read action unsupported on this deployment:', apiError);
      return { presidents: [], unsupportedAction: true };
    }
    throw new Error(apiError || 'PRESIDENTS_API_ERROR');
  }

  return {
    presidents: Array.isArray(json.presidents) ? json.presidents : [],
    unsupportedAction: false,
  };
}

async function ensureAllPresidentsLoaded(forceRefresh = false) {
  if (!forceRefresh && isBulkStateFresh(bulkPresidentsState)) {
    return bulkPresidentsState.rows;
  }

  if (bulkPresidentsState.promise) {
    return bulkPresidentsState.promise;
  }

  bulkPresidentsState.promise = (async () => {
    const attempts = presidentsReadActionUnsupported ? [false] : [true, false];
    let rows = [];
    let loaded = false;

    for (const withAction of attempts) {
      try {
        const result = await queryPresidentsByCity('', withAction);
        loaded = true;
        if (result?.unsupportedAction) {
          presidentsReadActionUnsupported = true;
        }
        rows = Array.isArray(result?.presidents) ? result.presidents : [];
        if (rows.length) break;
      } catch {
        // try next strategy
      }
    }

    if (!loaded) {
      throw new Error('PRESIDENTS_BULK_LOAD_FAILED');
    }

    bulkPresidentsState.rows = Array.isArray(rows) ? rows : [];
    bulkPresidentsState.fetchedAt = Date.now();
    hydratePresidentsCacheFromRows(bulkPresidentsState.rows);
    return bulkPresidentsState.rows;
  })();

  try {
    return await bulkPresidentsState.promise;
  } finally {
    bulkPresidentsState.promise = null;
  }
}

async function fetchPresidentByCity(cityName) {
  const cityArabic = normalizeCityToArabic(cityName);
  if (!cityArabic) return null;
  const selectedUnion = getSelectedUnionForCity(cityArabic);
  const effectiveCity = getCityForApi(cityArabic, selectedUnion) || getEffectiveCityName(cityArabic);
  const scopeKey = getScopeKey(cityArabic, selectedUnion);

  const cachedPresident = getCachedValue(presidentByCity, presidentsCacheMeta, scopeKey);
  if (cachedPresident !== undefined) return cachedPresident;

  try {
    const bulkRows = await ensureAllPresidentsLoaded(false);
    if (Array.isArray(bulkRows) && bulkRows.length) {
      const localMatch = pickPresidentFromRows(bulkRows, cityArabic, selectedUnion, effectiveCity);
      const localResult = localMatch || null;
      setCachedValue(presidentByCity, presidentsCacheMeta, scopeKey, localResult);
      return localResult;
    }
  } catch {
    // fall back to targeted city query
  }

  if (presidentRequestsInFlight.has(scopeKey)) {
    return presidentRequestsInFlight.get(scopeKey);
  }

  const requestPromise = (async () => {
    const queryCandidates = [];
    if (effectiveCity) queryCandidates.push(effectiveCity);
    if (!queryCandidates.includes(cityArabic)) queryCandidates.push(cityArabic);
    if (!(cityArabic === ISTANBUL_CITY && selectedUnion)) {
      const trName = AR_TO_TR[cityArabic] || '';
      if (trName && !queryCandidates.includes(trName)) queryCandidates.push(trName);
    }

    const targetLoose = normalizeArabicLoose(cityArabic);
    const effectiveTargetLoose = normalizeArabicLoose(effectiveCity || cityArabic);
    let lastHardError = null;
    let bestMatch = null;

    for (const queryCity of queryCandidates) {
      const attempts = presidentsReadActionUnsupported ? [false] : [true, false];
      for (const withAction of attempts) {
        let result;
        try {
          result = await queryPresidentsByCity(queryCity, withAction);
        } catch (err) {
          lastHardError = err;
          continue;
        }

        if (result?.unsupportedAction) {
          presidentsReadActionUnsupported = true;
        }

        const presidents = Array.isArray(result?.presidents) ? result.presidents : [];
        if (!presidents.length) continue;

        const cityMatched = presidents.filter((row) => {
          const rowCity = row?.city || row?.City || '';
          const rowArabic = normalizeCityToArabic(rowCity);
          const rowLoose = normalizeArabicLoose(rowArabic || rowCity);
          if (cityArabic === ISTANBUL_CITY && selectedUnion) {
            return rowLoose === targetLoose || rowLoose === effectiveTargetLoose;
          }
          return rowLoose === targetLoose;
        });

        let scopedRows = cityMatched;

        if (cityArabic === ISTANBUL_CITY && selectedUnion) {
          const selectedUnionLoose = normalizeArabicLoose(selectedUnion);
          const unionMatched = presidents.filter((row) => normalizeArabicLoose(getUnionFromRecord(row)) === selectedUnionLoose);

          if (scopedRows.length) {
            const cityAndUnion = scopedRows.filter((row) => normalizeArabicLoose(getUnionFromRecord(row)) === selectedUnionLoose);
            if (cityAndUnion.length) scopedRows = cityAndUnion;
          } else if (unionMatched.length) {
            scopedRows = unionMatched;
          }
        }

        if (!scopedRows.length) continue;

        const exact = scopedRows[0];

        if (exact) {
          bestMatch = exact;
          break;
        }

        if (!bestMatch) bestMatch = presidents[0];
      }

      if (bestMatch) break;
    }

    if (lastHardError && !bestMatch) throw lastHardError;

    const normalizedResult = bestMatch || null;
    setCachedValue(presidentByCity, presidentsCacheMeta, scopeKey, normalizedResult);
    return normalizedResult;
  })();

  presidentRequestsInFlight.set(scopeKey, requestPromise);
  try {
    return await requestPromise;
  } finally {
    presidentRequestsInFlight.delete(scopeKey);
  }
}

function renderPresidentState(state, president) {
  const statusEl = document.getElementById('presidentStatus');
  const nameEl = document.getElementById('presidentName');
  const phoneEl = document.getElementById('presidentPhone');
  if (!statusEl || !nameEl || !phoneEl) return;

  if (state === 'loading') {
    statusEl.textContent = 'جارٍ تحميل رئيس الاتحاد...';
    nameEl.textContent = '—';
    phoneEl.textContent = '—';
    return;
  }

  if (state === 'success') {
    const presidentName = safe(
      president?.president_name
      || president?.President_Name
      || president?.name
      || president?.Name
      || president?.president
      || president?.President
    );
    const presidentPhone = safe(
      president?.phone
      || president?.Phone
      || president?.phone_number
      || president?.Phone_Number
      || president?.contact_number
      || president?.Contact_Number
    );
    statusEl.textContent = 'رئيس الاتحاد الحالي';
    nameEl.textContent = `الاسم: ${presidentName}`;
    phoneEl.textContent = `الهاتف: ${presidentPhone}`;
    return;
  }

  if (state === 'not-found') {
    statusEl.textContent = 'لم يتم تعيين رئيس الاتحاد بعد.';
    nameEl.textContent = '—';
    phoneEl.textContent = '—';
    return;
  }

  if (state === 'error') {
    statusEl.textContent = 'تعذر تحميل رئيس الاتحاد حالياً.';
    nameEl.textContent = '—';
    phoneEl.textContent = '—';
    return;
  }

  statusEl.textContent = 'اختر مدينة لعرض رئيس الاتحاد.';
  nameEl.textContent = '—';
  phoneEl.textContent = '—';
}

async function updatePresidentForCity(cityName) {
  const cityKey = normalizeCityToArabic(cityName);
  const selectedUnion = getSelectedUnionForCity(cityKey);
  const scopeKey = getScopeKey(cityKey, selectedUnion);
  const cachedPresident = getCachedValue(presidentByCity, presidentsCacheMeta, scopeKey);
  if (cachedPresident !== undefined) {
    if (cachedPresident) renderPresidentState('success', cachedPresident);
    else renderPresidentState('not-found');
    return;
  }
  renderPresidentState('loading');
  try {
    const president = await fetchPresidentByCity(cityKey);
    if (president) renderPresidentState('success', president);
    else renderPresidentState('not-found');
  } catch (err) {
    console.error('Could not load president for city', cityName, err);
    renderPresidentState('error');
  }
}

// Example event data for each province
const EVENTS = {
  "إسطنبول": [
    { name: "ورشة عمل القيادة", office: "مكتب التدريب", day: "الإثنين" },
    { name: "ندوة ثقافية", office: "مكتب الثقافة", day: "الأربعاء" },
  ],
  "كوجالي": [
    { name: "دورة تطوير الذات", office: "مكتب التنمية", day: "الثلاثاء" },
  ],
  // Add more provinces and their events here
};

function showEvents(provinceName, managerName) {
  const cityKey = normalizeCityToArabic(provinceName);
  const events = filterRowsBySelectedUnion(cityKey, (getEventsForCity(cityKey) || EVENTS[cityKey]) || []);
  const eventsContainer = document.getElementById("events");
  const eventList = document.getElementById("eventList");

  // Clear previous events
  eventList.innerHTML = "";

  // Debugging info
  console.log("showEvents called for:", provinceName, "manager:", managerName, "found events:", events.length, events);
  if (!eventsContainer) {
    console.warn("events container not found in DOM");
    return;
  }

  if (events.length > 0) {
    events.forEach(ev => {
      const eventTitle = safe(ev.title ?? ev.name);
      const eventOffice = safe(ev.office_name ?? ev.office);
      const eventWhen = safe(ev.date ?? ev.day);
      const eventDescription = safe(ev.description);
      const descriptionRow = eventDescription !== '—'
        ? `<div class="eventDesc"><strong>الوصف:</strong> ${safe(stripUnionMarker(eventDescription))}</div>`
        : '';

      const li = document.createElement("li");
      li.className = "event-item";
      li.innerHTML = `
        <div class="eventInfo">
          <strong>اسم الحدث:</strong> ${eventTitle}<br>
          <strong>المكتب المسؤول:</strong> ${eventOffice}<br>
          <strong>اليوم:</strong> ${eventWhen}
          ${descriptionRow}
        </div>`;

      eventList.appendChild(li);
    });
  } else {
    // Show a friendly 'no events' message so panel is visible and user sees feedback
    const li = document.createElement("li");
    li.textContent = "لا توجد أحداث قريبة";
    eventList.appendChild(li);
  }
  eventsContainer.hidden = false;
}

function renderOfficesSection(state, cityName, offices = [], errorMessage = '') {
  const officesSection = document.getElementById('officesSection');
  const officesStatus = document.getElementById('officesStatus');
  const officeList = document.getElementById('officeList');
  if (!officesSection || !officesStatus || !officeList) return;
  const isMobile = (() => {
    try {
      return window.matchMedia('(max-width: 900px)').matches;
    } catch {
      return false;
    }
  })();

  officeList.innerHTML = '';

  if (state === 'idle') {
    officesStatus.hidden = false;
    officesStatus.textContent = 'اختر مدينة لعرض المكاتب.';
    return;
  }

  if (state === 'loading') {
    officesStatus.hidden = false;
    officesStatus.textContent = 'جارٍ تحميل مكاتب المدينة...';
    return;
  }

  if (state === 'error') {
    officesStatus.hidden = false;
    officesStatus.textContent = `تعذر تحميل المكاتب: ${errorMessage || 'خطأ غير معروف'}`;
    return;
  }

  if (!Array.isArray(offices) || offices.length === 0) {
    officesStatus.hidden = false;
    officesStatus.textContent = 'لا توجد مكاتب مسجلة لهذه المدينة';
    return;
  }

  officesStatus.hidden = isMobile;
  officesStatus.textContent = isMobile ? '' : `عدد المكاتب: ${offices.length}`;
  offices.forEach((office) => {
    const officeName = safe(pickOfficeValue(office, ['office_name', 'Office Name', 'officeName']));
    const officeManager = safe(pickOfficeValue(office, ['office_manager', 'Office Manager', 'officeManager', 'responsible']));
    const contactNumber = safe(pickOfficeValue(office, ['contact_number', 'Contact Number', 'contactNumber', 'phone']));
    const officeNotes = pickOfficeValue(office, ['notes', 'Notes']);
    const officeNotes2 = pickOfficeValue(office, ['notes_2', 'Notes 2', 'Notes2', 'notes2']);

    const li = document.createElement('li');
    li.className = 'officeItem';

    const notesHtml = [officeNotes, officeNotes2].filter(Boolean)
      .map((note) => `<div class="officeItemRow"><strong>ملاحظات:</strong> ${safe(stripUnionMarker(note))}</div>`)
      .join('');

    li.innerHTML = `
      <div class="officeItemTitle">${officeName}</div>
      <div class="officeItemRow"><strong>مدير المكتب:</strong> ${officeManager}</div>
      <div class="officeItemRow"><strong>رقم التواصل:</strong> ${contactNumber}</div>
      ${notesHtml}`;

    officeList.appendChild(li);
  });
}

async function ensureAllOfficesLoaded(forceRefresh = false) {
  if (!forceRefresh && isBulkStateFresh(bulkOfficesState)) {
    return bulkOfficesState.rows;
  }

  if (bulkOfficesState.promise) {
    return bulkOfficesState.promise;
  }

  bulkOfficesState.promise = (async () => {
    const url = `${OFFICES_API_URL}?action=offices`;
    const res = await fetch(url, { method: 'GET', redirect: 'follow' });
    if (!res.ok) throw new Error(`OFFICES_HTTP_${res.status}`);

    const text = await res.text();
    let json;
    try {
      json = JSON.parse(text);
    } catch {
      throw new Error('OFFICES_INVALID_JSON');
    }

    if (!json || json.ok !== true) {
      throw new Error(String(json?.error || json?.message || 'OFFICES_API_ERROR').trim());
    }

    const rows = Array.isArray(json.offices)
      ? json.offices
      : Array.isArray(json.data)
        ? json.data
        : [];

    bulkOfficesState.rows = rows;
    bulkOfficesState.fetchedAt = Date.now();
    hydrateOfficesCacheFromRows(rows);
    return rows;
  })();

  try {
    return await bulkOfficesState.promise;
  } finally {
    bulkOfficesState.promise = null;
  }
}

async function fetchOffices(city) {
  const cityKey = normalizeCityToArabic(city);
  if (!cityKey) {
    officesByCity.set('', []);
    return [];
  }

  const cachedOffices = getCachedValue(officesByCity, officesCacheMeta, cityKey);
  if (cachedOffices !== undefined) return cachedOffices;

  try {
    await ensureAllOfficesLoaded(false);
    const bulkCached = getCachedValue(officesByCity, officesCacheMeta, cityKey);
    if (bulkCached !== undefined) return bulkCached;
  } catch {
    // fall back to city-specific request
  }

  if (officeRequestsInFlight.has(cityKey)) {
    return officeRequestsInFlight.get(cityKey);
  }

  const requestPromise = (async () => {
    const url = `${OFFICES_API_URL}?action=offices&city=${encodeURIComponent(cityKey)}`;
    console.log('Offices API request URL:', url);

    const res = await fetch(url, { method: 'GET', redirect: 'follow' });
    const text = await res.text();
    console.log('Offices API response status:', res.status);

    let json;
    try {
      json = JSON.parse(text);
    } catch {
      console.error('Offices API invalid JSON:', text);
      throw new Error(`Invalid JSON response: ${text}`);
    }

    console.log('Offices API response JSON:', json);

    if (!res.ok) {
      const errMsg = String(json?.error || json?.message || `HTTP ${res.status}`).trim();
      throw new Error(errMsg);
    }

    if (!json || json.ok !== true) {
      const errMsg = String(json?.error || json?.message || 'API returned ok:false').trim();
      throw new Error(errMsg);
    }

    const offices = Array.isArray(json.offices)
      ? json.offices
      : Array.isArray(json.data)
        ? json.data
        : [];

    const normalized = offices.map((office) => {
      const officeCity = normalizeCityToArabic(office?.city || office?.City || cityKey) || cityKey;
      return { ...office, city: officeCity };
    });

    setCachedValue(officesByCity, officesCacheMeta, cityKey, normalized);
    return normalized;
  })();

  officeRequestsInFlight.set(cityKey, requestPromise);
  try {
    return await requestPromise;
  } finally {
    officeRequestsInFlight.delete(cityKey);
  }
}

async function pingOfficesApiOnLoad() {
  const officesStatus = document.getElementById('officesStatus');
  const url = `${OFFICES_API_URL}?action=ping`;
  console.log('Offices API ping URL:', url);

  try {
    const res = await fetch(url, {
      method: 'GET',
      mode: 'cors',
      redirect: 'follow',
    });
    const text = await res.text();
    console.log('Offices API ping status:', res.status);
    console.log('Offices API ping raw response text:', text);

    let json = null;
    try { json = JSON.parse(text); } catch (e) { console.warn('Offices ping JSON parse failed:', e); }
    console.log('Offices API ping parsed JSON:', json);

    if (!res.ok || (json && json.ok === false)) {
      const apiError = String(json?.error || json?.message || `HTTP ${res.status}`).trim();
      if (officesStatus) {
        officesStatus.textContent = `تعذر الوصول إلى واجهة المكاتب أو لم يتم نشرها بصلاحية عامة: ${apiError}`;
        officesStatus.style.color = '#b91c1c';
      }
    }
  } catch (err) {
    console.error('Offices API ping failed:', err);
    if (officesStatus) {
      officesStatus.textContent = 'تعذر الوصول إلى واجهة المكاتب أو لم يتم نشرها بصلاحية عامة';
      officesStatus.style.color = '#b91c1c';
    }
  }
}

async function updateOfficesForCity(cityName) {
  const cityKey = normalizeCityToArabic(cityName);
  if (!cityKey) {
    renderOfficesSection('idle', '');
    return;
  }

  const cachedOffices = getCachedValue(officesByCity, officesCacheMeta, cityKey);
  if (cachedOffices !== undefined) {
    renderOfficesSection('success', cityKey, filterRowsBySelectedUnion(cityKey, cachedOffices || []));
    return;
  }
  try {
    const offices = await fetchOffices(cityKey);
    renderOfficesSection('success', cityKey, filterRowsBySelectedUnion(cityKey, offices));
  } catch (err) {
    console.error('Could not load offices for city', cityKey, err);
    renderOfficesSection('error', cityKey, [], String(err?.message || err));
  }
}

(async function main() {
  const [regions, geo] = await Promise.all([loadJSON(PATHS.regions), loadJSON(PATHS.geo)]);

  // Fetch events from Apps Script backend (populate eventsByCity)
  // Fetch events from Apps Script backend. If `city` supplied, fetch only that city's events.
  async function fetchEvents(city) {
    try {
      const url = APPS_SCRIPT_URL + (city ? `?action=events&city=${encodeURIComponent(city)}` : `?action=events`);
      const res = await fetch(url);
      if (!res.ok) throw new Error('Failed to fetch events: ' + res.status);
      const json = await res.json();
      if (!json || json.ok !== true) throw new Error('API returned error');
      const arr = Array.isArray(json.events) ? json.events : [];
      const requestedCity = normalizeCityToArabic(city || '');

      if (city) {
        // replace only this city's events using canonical Arabic city key
        const scoped = arr.filter((ev) => {
          const evCity = normalizeCityToArabic(ev.city || ev.City || '');
          return !requestedCity || evCity === requestedCity;
        });

        const normalizedScoped = scoped.map((ev) => ({
          ...ev,
          city: normalizeCityToArabic(ev.city || ev.City || requestedCity || '') || requestedCity,
        }));

        if (requestedCity) {
          eventsByCity.set(requestedCity, normalizedScoped);
        } else if (city) {
          eventsByCity.set(String(city).trim(), normalizedScoped);
        }
      } else {
        // rebuild full index
        eventsByCity.clear();
        arr.forEach(ev => {
          const c = ev.city || ev.City || '';
          const key = normalizeCityToArabic(c) || String(c || '').trim();
          const normalizedEvent = {
            ...ev,
            city: normalizeCityToArabic(c) || String(c || '').trim(),
          };

          if (!eventsByCity.has(key)) eventsByCity.set(key, []);
          eventsByCity.get(key).push(normalizedEvent);
        });
      }
      console.log('Fetched events:', eventsByCity);
      return arr;
    } catch (err) {
      console.warn('Could not fetch events from API, using local EVENTS as fallback', err);
      return [];
    }
  }

async function addEvent(eventObj) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    // ✅ avoid CORS preflight with Apps Script
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(eventObj),
  });

  const text = await res.text();
  console.log("POST status:", res.status);
  console.log("POST raw response:", text);

  let json;
  try { json = JSON.parse(text); } catch { json = { ok: false, error: text }; }

  if (!res.ok) throw new Error(`HTTP ${res.status}: ${text}`);
  if (json && json.ok === false) throw new Error(`API ok:false: ${text}`);

  return json;
}

async function addOffice(officeObj) {
  const url = OFFICES_API_URL;
  const payload = officeObj;

  console.log('POST url', url);
  console.log('payload', payload);

  const parseOfficeResponse = async (res, strategy) => {
    console.log(`[${strategy}] response status`, res.status);
    const text = await res.text();
    console.log(`[${strategy}] raw response text`, text);

    let json = null;
    try {
      json = JSON.parse(text);
    } catch (e) {
      console.warn(`[${strategy}] POST JSON parse failed:`, e);
    }
    console.log(`[${strategy}] parsed json`, json);

    const contentType = (res.headers.get('content-type') || '').toLowerCase();
    const allowHeader = (res.headers.get('allow') || '').toUpperCase();
    const trimmed = String(text || '').trim();
    const isHtml = contentType.includes('text/html') || /^<!doctype html|^<html/i.test(trimmed);
    const apiError = String(json?.error || json?.message || '').trim();

    if (!res.ok || (json && json.ok === false)) {
      let details = apiError || `HTTP ${res.status}`;
      if (res.status === 405 || (allowHeader && !allowHeader.includes('POST'))) {
        details = 'هذا الرابط يدعم GET فقط حالياً (405). فعّل doPost ثم أعد نشر Web App بصيغة /exec.';
      } else if (isHtml) {
        details = 'الخادم أعاد HTML/Redirect بدل JSON. تحقق من رابط /exec وصلاحية Anyone.';
      }

      const err = new Error(details);
      err.apiError = apiError;
      err.status = res.status;
      err.rawText = text;
      err.parsedJson = json;
      err.allow = allowHeader;
      throw err;
    }

    if (!json) {
      const err = new Error('API returned non-JSON response.');
      err.status = res.status;
      err.rawText = text;
      throw err;
    }

    return json;
  };

  try {
    const res = await fetch(url, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      redirect: 'follow',
    });
    return await parseOfficeResponse(res, 'application/json');
  } catch (err) {
    const isFetchError = err && (err.name === 'TypeError' || /Failed to fetch/i.test(String(err.message || '')));
    if (!isFetchError) throw err;

    console.warn('Primary JSON POST failed at network/CORS layer. Retrying with text/plain to bypass preflight...', err);
    try {
      const fallbackRes = await fetch(url, {
        method: 'POST',
        mode: 'cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        redirect: 'follow',
      });
      return await parseOfficeResponse(fallbackRes, 'text/plain');
    } catch (fallbackErr) {
      const stillFetchError = fallbackErr && (fallbackErr.name === 'TypeError' || /Failed to fetch/i.test(String(fallbackErr.message || '')));
      if (stillFetchError) {
        const e = new Error('API unreachable (CORS/permission). Check Apps Script deployment: Anyone + Web App /exec');
        e.cause = fallbackErr;
        throw e;
      }
      throw fallbackErr;
    }
  }
}

async function addPresident(presidentObj) {
  const requestBody = JSON.stringify(presidentObj);
  console.log('President payload:', presidentObj);
  const logPresidentRequestDiagnostics = (err, extra = {}) => {
    console.error('President API error object:', err);
    console.error('President API diagnostics:', {
      endpoint: PRESIDENTS_API_URL,
      ...extra,
    });
    console.warn('Possible causes to verify:');
    console.warn('1) CORS restrictions');
    console.warn('2) OPTIONS preflight failure');
    console.warn('3) Web App not deployed as Anyone');
    console.warn('4) Redirect/login page response');
    console.warn('5) Browser blocked non-simple cross-origin request');
  };

  try {
    const res = await fetch(PRESIDENTS_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: requestBody,
      redirect: 'follow',
    });

    const contentType = (res.headers.get('content-type') || '').toLowerCase();
    const text = await res.text();
    const textPreview = String(text || '').slice(0, 300);
    console.log('President API POST diagnostics:', {
      endpoint: PRESIDENTS_API_URL,
      status: res.status,
      contentType,
      preview: textPreview,
    });

    if (!res.ok) {
      const err = new Error('HTTP_NOT_OK');
      err.code = 'HTTP_NOT_OK';
      err.status = res.status;
      err.statusText = res.statusText;
      err.responseText = text;
      err.contentType = contentType;
      throw err;
    }

    const trimmed = String(text || '').trim();
    if (contentType.includes('text/html') || /^<!doctype html|^<html/i.test(trimmed)) {
      const err = new Error('HTML_RESPONSE');
      err.code = 'HTML_RESPONSE';
      err.responseText = text;
      err.contentType = contentType;
      throw err;
    }

    let json;
    try {
      json = JSON.parse(text);
    } catch (parseErr) {
      const err = new Error('INVALID_JSON_RESPONSE');
      err.code = 'INVALID_JSON_RESPONSE';
      err.responseText = text;
      err.contentType = contentType;
      err.parseError = parseErr;
      throw err;
    }

    if (!json || json.ok !== true) {
      const serverError = (json && (json.error || json.message)) || 'API_LOGICAL_ERROR';
      const err = new Error(serverError);
      err.code = 'API_LOGICAL_ERROR';
      err.apiError = json && (json.error || json.message || json.code);
      err.responseText = text;
      err.contentType = contentType;
      throw err;
    }

    return json;
  } catch (err) {
    logPresidentRequestDiagnostics(err, {
      strategy: 'text/plain;charset=utf-8',
      method: 'POST',
      redirect: 'follow',
      contentType: 'text/plain;charset=utf-8',
      status: err && err.status,
      statusText: err && err.statusText,
      responseText: err && String(err.responseText || '').slice(0, 300),
    });
    throw err;
  }
}

  // Fetch all events (no city filter) and return array
  async function fetchAllEvents() {
    try {
      const url = APPS_SCRIPT_URL + '?action=events';
      const res = await fetch(url);
      if (!res.ok) throw new Error('Failed to fetch all events: ' + res.status);
      const json = await res.json();
      if (!json || json.ok !== true) return [];
      return Array.isArray(json.events) ? json.events : [];
    } catch (err) {
      console.warn('fetchAllEvents failed', err);
      return [];
    }
  }

  // Generate next event id by fetching current events from API (ensures up-to-date)
  async function generateNextEventId() {
    let all = [];

    if (eventsByCity.size > 0) {
      eventsByCity.forEach((rows) => {
        if (Array.isArray(rows)) all.push(...rows);
      });
    }

    if (!all.length) {
      all = await fetchAllEvents();
    }

    let max = 0;
    all.forEach(ev => {
      const id = (ev.event_id || ev.eventId || ev.id || '').toString().trim();
      const m = id.match(/^EVT0*?(\d+)$/i);
      if (m) {
        const n = parseInt(m[1], 10);
        if (!Number.isNaN(n) && n > max) max = n;
      }
    });
    const next = max + 1;
    const padded = String(next).padStart(3, '0');
    return `EVT${padded}`;
  }

  ensureAllOfficesLoaded(false).catch((err) => {
    console.warn('Initial offices preload failed:', err);
  });

  ensureAllPresidentsLoaded(false).catch((err) => {
    console.warn('Initial presidents preload failed:', err);
  });

  // initial fetch of events
  await fetchEvents();

  // Ensure admin and events UI are hidden on initial load
  const eventsContainerInit = document.getElementById('events');
  if (eventsContainerInit) eventsContainerInit.hidden = true;
  const adminSectionInitEl = document.getElementById('adminSection');
  if (adminSectionInitEl) adminSectionInitEl.hidden = true;
  const addModalInit = document.getElementById('addModal'); if (addModalInit) addModalInit.hidden = true;
  const presidentModalInit = document.getElementById('presidentModal'); if (presidentModalInit) presidentModalInit.hidden = true;
  const officeModalInit = document.getElementById('officeModal'); if (officeModalInit) officeModalInit.hidden = true;

  buildLegend(regions);
  setPanel(null);
  renderPresidentState('idle');
  renderOfficesSection('idle', '');
  pingOfficesApiOnLoad();

  // Lookup: normalized(TR name) -> region
  const provinceToRegion = new Map();
  // Lookup: normalized(TR name) -> Arabic label
  const provinceToArabic = new Map();

  for (const r of regions) {
    for (const cityAr of r.cities) {
      const tr = AR_TO_TR[cityAr] || cityAr;
      const keyA = normalizeTR(tr);
      provinceToRegion.set(keyA, r);
      provinceToArabic.set(keyA, cityAr);
    }
  }

  async function warmupCityData(regionsList) {
    const uniqueCities = [];
    const seen = new Set();

    for (const region of regionsList || []) {
      const cities = Array.isArray(region?.cities) ? region.cities : [];
      for (const city of cities) {
        const cityKey = normalizeCityToArabic(city);
        if (!cityKey || seen.has(cityKey)) continue;
        seen.add(cityKey);
        uniqueCities.push(cityKey);
      }
    }

    if (!uniqueCities.length) return;

    const hydratePresidentCache = (rows) => {
      const rowsList = Array.isArray(rows) ? rows : [];
      const byCity = new Map();

      rowsList.forEach((row) => {
        const cityRaw = row?.city || row?.City || '';
        const cityKey = normalizeCityToArabic(cityRaw);
        if (!cityKey || byCity.has(cityKey)) return;
        byCity.set(cityKey, row);
      });

      uniqueCities.forEach((city) => {
        const item = byCity.has(city) ? byCity.get(city) : null;
        setCachedValue(presidentByCity, presidentsCacheMeta, city, item);
      });
    };

    const hydrateOfficesCache = (rows) => {
      const grouped = new Map();
      const rowsList = Array.isArray(rows) ? rows : [];

      rowsList.forEach((office) => {
        const cityKey = normalizeCityToArabic(office?.city || office?.City || '');
        if (!cityKey) return;
        if (!grouped.has(cityKey)) grouped.set(cityKey, []);
        grouped.get(cityKey).push({ ...office, city: cityKey });
      });

      uniqueCities.forEach((city) => {
        const offices = grouped.get(city) || [];
        setCachedValue(officesByCity, officesCacheMeta, city, offices);
      });
    };

    const fetchAllPresidentsForWarmup = async () => {
      try {
        const rows = await ensureAllPresidentsLoaded(false);
        if (!Array.isArray(rows) || !rows.length) return false;
        hydratePresidentCache(rows);
        return true;
      } catch {
        return false;
      }
    };

    const fetchAllOfficesForWarmup = async () => {
      try {
        const rows = await ensureAllOfficesLoaded(false);
        if (!Array.isArray(rows)) return false;
        hydrateOfficesCache(rows);
        return rows.length > 0;
      } catch {
        return false;
      }
    };

    await Promise.allSettled([
      fetchAllPresidentsForWarmup(),
      fetchAllOfficesForWarmup(),
    ]);

    const remainingCities = uniqueCities.filter((city) => {
      const hasPresident = getCachedValue(presidentByCity, presidentsCacheMeta, city) !== undefined;
      const hasOffices = getCachedValue(officesByCity, officesCacheMeta, city) !== undefined;
      return !hasPresident || !hasOffices;
    });

    if (!remainingCities.length) {
      console.log('City data warmup completed from bulk fetch for presidents/offices:', uniqueCities.length);
      return;
    }

    let nextIndex = 0;
    const workerCount = Math.min(CITY_WARMUP_CONCURRENCY, remainingCities.length);
    const workers = Array.from({ length: workerCount }, async () => {
      while (nextIndex < remainingCities.length) {
        const city = remainingCities[nextIndex];
        nextIndex += 1;
        await Promise.allSettled([
          fetchPresidentByCity(city),
          fetchOffices(city),
        ]);
      }
    });

    await Promise.allSettled(workers);
    console.log('City data warmup completed for presidents/offices:', uniqueCities.length, 'remaining fetched:', remainingCities.length);
  }

  const box = document.getElementById("mapBox");
  const leafletContainer = document.getElementById('leafletMap');
  const resetBtn = document.getElementById("resetBtn");

  const isTouchEnvironment = () => {
    try {
      return window.matchMedia('(pointer: coarse)').matches
        || ('ontouchstart' in window)
        || (navigator && navigator.maxTouchPoints > 0);
    } catch {
      return ('ontouchstart' in window);
    }
  };

  function installMapScrollLock(targetEl) {
    if (!targetEl || targetEl.dataset.scrollLockInstalled === '1') return;
    targetEl.dataset.scrollLockInstalled = '1';
    const touchDevice = isTouchEnvironment();

    const stopGestureScroll = (e) => {
      e.preventDefault();
      e.stopPropagation();
    };

    const stopNativeDrag = (e) => {
      e.preventDefault();
    };

    if (!touchDevice) {
      targetEl.addEventListener('wheel', stopGestureScroll, { passive: false });
    }
    targetEl.addEventListener('dragstart', stopNativeDrag);
  }

  // Admin UI elements
  const adminSection = document.getElementById('adminSection');
  const adminForm = document.getElementById('adminForm');
  const adminUser = document.getElementById('adminUser');
  const adminPass = document.getElementById('adminPass');
  const adminError = document.getElementById('adminError');
  const adminContent = document.getElementById('adminContent');
  const adminInfo = document.getElementById('adminInfo');
  const adminLogout = document.getElementById('adminLogout');
  const adminCancel = document.getElementById('adminCancel');
  const manageCityBtn = document.getElementById('manageCityBtn');
  const manageHint = document.getElementById('manageHint');
  const istanbulUnionModal = document.getElementById('istanbulUnionModal');
  const istanbulUnionCancel = document.getElementById('istanbulUnionCancel');
  const istanbulUnionButtons = istanbulUnionModal
    ? Array.from(istanbulUnionModal.querySelectorAll('[data-union]'))
    : [];
  let currentAdminProvince = null;
  let selectedCity = null;
  let lastManager = null;
  let clearSelection = () => {};
  let pendingIstanbulResolver = null;

  function promptIstanbulUnionSelection() {
    if (!istanbulUnionModal) return Promise.resolve(ISTANBUL_UNIONS[0]);

    return new Promise((resolve) => {
      const finalize = (value) => {
        istanbulUnionModal.hidden = true;
        pendingIstanbulResolver = null;
        resolve(value);
      };

      pendingIstanbulResolver = finalize;
      istanbulUnionModal.hidden = false;
    });
  }

  if (istanbulUnionCancel && !istanbulUnionCancel.dataset.bound) {
    istanbulUnionCancel.addEventListener('click', () => {
      if (pendingIstanbulResolver) pendingIstanbulResolver(null);
    });
    istanbulUnionCancel.dataset.bound = '1';
  }

  istanbulUnionButtons.forEach((button) => {
    if (button.dataset.bound) return;
    button.addEventListener('click', () => {
      const unionName = String(button.dataset.union || '').trim();
      if (pendingIstanbulResolver) pendingIstanbulResolver(unionName || null);
    });
    button.dataset.bound = '1';
  });

  function renderCityActionButtons(city, loggedIn) {
    const actionButtons = [
      document.getElementById('actionAddOffice'),
      document.getElementById('actionAddPresident'),
      document.getElementById('actionAddEvent'),
    ].filter(Boolean);

    const hasCity = !!String(city || '').trim();
    actionButtons.forEach((button) => {
      button.hidden = !hasCity;
      button.classList.toggle('is-disabled', !loggedIn);
      button.setAttribute('aria-disabled', loggedIn ? 'false' : 'true');
    });
  }

  function setAdminHeader(city) {
    const cityText = String(city || '').trim();
    const title = adminSection ? adminSection.querySelector('.admin-title') : null;
    if (title) {
      title.textContent = cityText ? `الإدارة — ${cityText}` : 'الإدارة';
    }
    if (adminInfo) {
      adminInfo.textContent = cityText ? `لوحة إدارة ${cityText}` : 'لوحة إدارة';
    }
  }

  function syncAdminUI(options = {}) {
    const forceShow = !!options.forceShow;
    const city = String(selectedCity || currentAdminProvince || '').trim();
    const cityLabel = getEffectiveCityName(city);

    setAdminHeader(cityLabel);
    renderCityActionButtons(city, isAdminLoggedIn);

    if (!adminSection) return;
    if (forceShow) {
      adminSection.hidden = !city;
    }

    if (isAdminLoggedIn) {
      if (adminForm) adminForm.hidden = true;
      if (adminContent) adminContent.hidden = false;
      if (city && adminInfo) adminInfo.textContent = `لوحة إدارة ${cityLabel}`;
    } else {
      if (adminForm) adminForm.hidden = false;
      if (adminContent) adminContent.hidden = true;
    }

    if (adminError) adminError.style.display = 'none';
  }

  function setSelectedCity(cityAr) {
    const city = normalizeCityToArabic(cityAr);
    selectedCity = city;
    currentAdminProvince = city;
    const cityLabel = getEffectiveCityName(city);
    const selectedUnion = getSelectedUnionForCity(city);

    if (manageCityBtn) {
      manageCityBtn.disabled = !city;
      manageCityBtn.textContent = city ? `إدارة المدينة — ${cityLabel}` : 'إدارة المدينة';
    }
    if (manageHint) {
      manageHint.textContent = !city
        ? 'اختر مدينة أولاً'
        : (selectedUnion ? `الاتحاد المختار: ${selectedUnion}` : '');
    }

    setAdminHeader(cityLabel);
    if (adminSection && !adminSection.hidden) {
      syncAdminUI();
    }
  }

  function showAdminSection(provinceName) {
    setSelectedCity(provinceName);
    syncAdminUI({ forceShow: true });
    if (!isAdminLoggedIn && adminForm) {
      adminForm.reset();
    }
  }

  if (adminForm) {
    adminForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const u = (adminUser && adminUser.value || '').trim();
      const p = (adminPass && adminPass.value || '');
      if (u === 'admin' && p === 'admin') {
        // mark globally logged in and show FAB for all cities
        isAdminLoggedIn = true;
        sessionStorage.setItem(ADMIN_SESSION_KEY, 'true');
        syncAdminUI({ forceShow: true });
        if (selectedCity) {
          renderOfficesSection('success', selectedCity, getOfficesForCity(selectedCity));
        }
      } else {
        if (adminError) { adminError.textContent = 'اسم المستخدم أو كلمة المرور غير صحيحة'; adminError.style.display = 'block'; }
      }
    });
  }

  if (adminLogout) {
    adminLogout.addEventListener('click', () => {
      // clear global admin session and hide FAB
      sessionStorage.removeItem(ADMIN_SESSION_KEY);
      isAdminLoggedIn = false;
      syncAdminUI({ forceShow: true });
      if (adminForm) adminForm.reset();
      if (selectedCity) {
        renderOfficesSection('success', selectedCity, getOfficesForCity(selectedCity));
      }
    });
  }

  if (adminCancel) {
    adminCancel.addEventListener('click', () => {
      clearSelection();
    });
  }

  // Open admin panel explicitly via button
  // Manage-city button opens admin login for the selected city
  if (manageCityBtn) {
    manageCityBtn.addEventListener('click', () => {
      if (!selectedCity) {
        if (manageHint) manageHint.textContent = 'اختر مدينة أولاً';
        return;
      }
      showAdminSection(selectedCity);
      if (!isAdminLoggedIn && adminUser && typeof adminUser.focus === 'function') {
        adminUser.focus();
      }
    });
  }

  let map = null;
  let provincesLayer = null;
  let selectedProvinceLayer = null;
  let labelsLayer = null;
  let turkeyBounds = null;
  const provinceLabels = new Map();

  function provinceStyle(feature) {
    const trName = pickProvinceNameFromProps(feature?.properties || {});
    const key = normalizeTR(trName);
    const region = provinceToRegion.get(key);
    return {
      color: '#111827',
      weight: 1,
      bubblingMouseEvents: false,
      className: 'province-shape',
      fillColor: region ? (COLOR_MAP[region.color] || region.color) : '#eef2f1',
      fillOpacity: region ? 0.9 : 0.7,
    };
  }

  function buildLabelRect(entry) {
    const p = map.latLngToContainerPoint(entry.latlng);
    const fontSize = 11;
    const width = Math.max(36, entry.arName.length * (fontSize * 0.62) + 10);
    const height = 18;
    return {
      left: p.x - width / 2,
      right: p.x + width / 2,
      top: p.y - height / 2,
      bottom: p.y + height / 2,
    };
  }

  function rectsOverlap(a, b) {
    return !(a.right < b.left || a.left > b.right || a.bottom < b.top || a.top > b.bottom);
  }

  function polygonRingArea(ring) {
    if (!Array.isArray(ring) || ring.length < 3) return 0;
    let area = 0;
    for (let index = 0; index < ring.length; index += 1) {
      const [x1, y1] = ring[index] || [0, 0];
      const [x2, y2] = ring[(index + 1) % ring.length] || [0, 0];
      area += (x1 * y2) - (x2 * y1);
    }
    return Math.abs(area / 2);
  }

  function getProvinceLabelLatLng(feature, layer) {
    const geometry = feature?.geometry;
    const polylabelFn = (typeof window !== 'undefined' && typeof window.polylabel === 'function')
      ? window.polylabel
      : null;

    if (geometry && polylabelFn) {
      try {
        if (geometry.type === 'Polygon' && Array.isArray(geometry.coordinates)) {
          const point = polylabelFn(geometry.coordinates, 1.0);
          if (Array.isArray(point) && point.length >= 2) {
            return L.latLng(point[1], point[0]);
          }
        }

        if (geometry.type === 'MultiPolygon' && Array.isArray(geometry.coordinates)) {
          let bestPoint = null;
          let bestArea = -1;

          geometry.coordinates.forEach((polygon) => {
            if (!Array.isArray(polygon) || !Array.isArray(polygon[0])) return;
            const area = polygonRingArea(polygon[0]);
            const point = polylabelFn(polygon, 1.0);
            if (!Array.isArray(point) || point.length < 2) return;
            if (area > bestArea) {
              bestArea = area;
              bestPoint = point;
            }
          });

          if (bestPoint) {
            return L.latLng(bestPoint[1], bestPoint[0]);
          }
        }
      } catch (error) {
        console.warn('polylabel failed, fallback to bounds center', error);
      }
    }

    return layer.getBounds().getCenter();
  }

  function placeProvinceLabels() {
    if (!map || !labelsLayer) return;

    labelsLayer.clearLayers();
    const placedRects = [];
    const entries = Array.from(provinceLabels.values());

    entries.forEach((entry) => {
      entry.visible = false;
    });

    entries.forEach((entry) => {
      const rect = buildLabelRect(entry);
      const hasOverlap = placedRects.some((otherRect) => rectsOverlap(rect, otherRect));
      if (hasOverlap) return;

      labelsLayer.addLayer(entry.marker);
      entry.visible = true;
      placedRects.push(rect);
    });
  }

  clearSelection = () => {
    if (selectedProvinceLayer && provincesLayer) {
      provincesLayer.resetStyle(selectedProvinceLayer);
      selectedProvinceLayer = null;
    }
    setPanel(null);
    const eventsContainer = document.getElementById("events");
    const eventList = document.getElementById("eventList");
    if (eventList) eventList.innerHTML = "";
    if (eventsContainer) eventsContainer.hidden = true;
    const adminSectionEl = document.getElementById('adminSection');
    if (adminSectionEl) adminSectionEl.hidden = true;
    selectedCity = null;
    currentAdminProvince = null;
    selectedUnionByCity.clear();
    lastManager = null;
    if (manageCityBtn) { manageCityBtn.disabled = true; manageCityBtn.textContent = 'إدارة المدينة'; }
    if (manageHint) { manageHint.textContent = 'اختر مدينة أولاً'; }
    setAdminHeader('');
    renderPresidentState('idle');
    renderOfficesSection('idle', '');
    syncAdminUI();
  };

  function initLeafletMap() {
    if (!leafletContainer || typeof L === 'undefined') return;

    map = L.map('leafletMap', {
      zoomControl: true,
      scrollWheelZoom: false,
      dragging: true,
      touchZoom: true,
      tap: true,
    });
    map.setView([39.0, 35.0], 6);

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; OpenStreetMap contributors',
    }).addTo(map);

    map.createPane('labels');
    map.getPane('labels').style.zIndex = '620';
    map.getPane('labels').style.pointerEvents = 'none';

    labelsLayer = L.layerGroup().addTo(map);

    provincesLayer = L.geoJSON(geo, {
      style: provinceStyle,
      onEachFeature: (feature, lyr) => {
        const trName = pickProvinceNameFromProps(feature.properties);
        const key = normalizeTR(trName);
        const region = provinceToRegion.get(key);
        const arName = provinceToArabic.get(key) || TR_TO_AR[trName] || trName;
        const centroid = getProvinceLabelLatLng(feature, lyr);
        const labelMarker = L.marker(centroid, {
          icon: L.divIcon({
            className: 'province-label',
            html: `<span>${arName}</span>`,
            iconAnchor: [0, 0],
          }),
          pane: 'labels',
          interactive: false,
          keyboard: false,
        });

        const labelEntry = {
          id: L.stamp(lyr),
          arName,
          marker: labelMarker,
          layer: lyr,
          latlng: centroid,
          visible: false,
        };
        provinceLabels.set(labelEntry.id, labelEntry);

        if (region) {
          lyr.on('mouseover', () => {
            const layerEl = lyr.getElement();
            if (layerEl) layerEl.classList.add('province-hover-glow');

            if (selectedProvinceLayer !== lyr) {
              lyr.setStyle({ weight: 2, color: '#0f172a' });
            }
          });

          lyr.on('mouseout', () => {
            const layerEl = lyr.getElement();
            if (layerEl) layerEl.classList.remove('province-hover-glow');

            if (selectedProvinceLayer !== lyr) {
              provincesLayer.resetStyle(lyr);
            }
          });
        }

        lyr.on('click', async (evt) => {
          if (evt && evt.originalEvent) {
            L.DomEvent.stopPropagation(evt.originalEvent);
          }

          if (!region) {
            clearSelection();
            return;
          }

          const cityKey = normalizeCityToArabic(arName);
          if (cityKey === ISTANBUL_CITY) {
            const selectedUnion = await promptIstanbulUnionSelection();
            if (!selectedUnion) return;
            selectedUnionByCity.set(ISTANBUL_CITY, selectedUnion);
          } else {
            selectedUnionByCity.delete(ISTANBUL_CITY);
          }

          map.fitBounds(lyr.getBounds(), {
            padding: [40, 40],
            maxZoom: 7,
            animate: true,
            duration: 0.85,
            easeLinearity: 0.25,
          });

          if (selectedProvinceLayer && provincesLayer) {
            provincesLayer.resetStyle(selectedProvinceLayer);
          }
          selectedProvinceLayer = lyr;
          lyr.setStyle({ weight: 2, fillOpacity: 1 });
          lyr.bringToFront();

          const manager = region.manager || null;
          const selectedUnion = getSelectedUnionForCity(arName);
          const managerLabel = manager
            ? (selectedUnion ? `${manager} — ${getEffectiveCityName(arName)}` : manager)
            : null;
          setPanel(managerLabel);
          lastManager = manager;
          showEvents(arName, manager);
          updatePresidentForCity(arName);
          updateOfficesForCity(arName);

          setSelectedCity(arName);
          syncAdminUI();
        });
      },
    }).addTo(map);

    turkeyBounds = provincesLayer.getBounds();
    map.fitBounds(turkeyBounds, { padding: [20, 20] });
    placeProvinceLabels();

    map.on('click', () => clearSelection());
    map.on('focus click', () => map.scrollWheelZoom.enable());
    map.on('zoomend moveend', () => placeProvinceLabels());

    const mapEl = map.getContainer();
    installMapScrollLock(mapEl);
    mapEl.addEventListener('mouseleave', () => map.scrollWheelZoom.disable());

    if (resetBtn) {
      resetBtn.onclick = () => {
        if (turkeyBounds) {
          map.fitBounds(turkeyBounds, {
            padding: [20, 20],
            animate: true,
            duration: 0.5,
          });
        }
      };
      resetBtn.disabled = false;
      resetBtn.style.opacity = '1';
      resetBtn.style.cursor = 'pointer';
      resetBtn.title = 'إعادة ضبط الخريطة';
    }
  }

  warmupCityData(regions).catch((err) => {
    console.warn('City data warmup failed:', err);
  });
  initLeafletMap();

      // City panel action buttons + modals wiring
      const actionAddOffice = document.getElementById('actionAddOffice');
      const actionAddPresident = document.getElementById('actionAddPresident');
      const actionAddEvent = document.getElementById('actionAddEvent');

      const addModal = document.getElementById('addModal');
      const addForm = document.getElementById('addEventForm');
      const addCancel = document.getElementById('addCancel');
      const addMsg = document.getElementById('addMsg');

      const presidentModal = document.getElementById('presidentModal');
      const presidentForm = document.getElementById('addPresidentForm');
      const presidentCancel = document.getElementById('presidentCancel');
      let presidentStatusTimer = null;

      const officeModal = document.getElementById('officeModal');
      const officeForm = document.getElementById('addOfficeForm');
      const officeCancel = document.getElementById('officeCancel');
      const officeStatus = document.getElementById('officeStatus');

      const presidentSaveBtn = document.getElementById('presidentSave');
      const officeSaveBtn = document.getElementById('officeSave');
      if (presidentSaveBtn) presidentSaveBtn.textContent = 'حفظ';
      if (presidentCancel) presidentCancel.textContent = 'إلغاء';
      if (officeSaveBtn) officeSaveBtn.textContent = 'حفظ';
      if (officeCancel) officeCancel.textContent = 'إلغاء';

      const requireAdminForAction = () => {
        if (isAdminLoggedIn) return true;
        alert('يجب تسجيل الدخول كمسؤول');
        return false;
      };

      const fillCityInForm = (formEl) => {
        if (!formEl) return;
        const cityInput = formEl.querySelector('input[name="city"]');
        if (!cityInput) return;
        cityInput.value = getEffectiveCityName(selectedCity || currentAdminProvince || '');
      };

      const openAddEventModal = async () => {
        if (!addModal || !addForm) return;
        if (!selectedCity && !currentAdminProvince) return alert('اختر مدينة أولاً');

        addForm.reset();
        addModal.hidden = false;
        addMsg.style.display = 'none';
        addMsg.textContent = '';
        const cityInput = addForm.querySelector('input[name="city"]');
        const idInput = addForm.querySelector('input[name="event_id"]');
        const saveBtnOpen = document.getElementById('addSave');
        if (saveBtnOpen) {
          saveBtnOpen.disabled = false;
          saveBtnOpen.classList.remove('is-loading');
          saveBtnOpen.textContent = 'حفظ';
        }
        const cityVal = selectedCity || currentAdminProvince || '';
        const effectiveCityVal = getEffectiveCityName(cityVal);
        if (cityInput && effectiveCityVal) cityInput.value = effectiveCityVal;
        if (idInput) {
          let candidate = await generateNextEventId();
          const existing = new Set();
          for (const arr of eventsByCity.values()) { (arr||[]).forEach(e=> existing.add((e.event_id||e.eventId||'').toString())); }
          while (existing.has(candidate)) {
            const num = parseInt(candidate.replace(/^EVT/i, ''), 10) || 0;
            candidate = `EVT${String(num+1).padStart(3,'0')}`;
          }
          idInput.value = candidate;
          idInput.readOnly = true;
        }
      };

      if (actionAddEvent && !actionAddEvent.dataset.bound) {
        actionAddEvent.addEventListener('click', async () => {
          if (!requireAdminForAction()) return;
          await openAddEventModal();
        });
        actionAddEvent.dataset.bound = '1';
      }

      if (actionAddPresident && !actionAddPresident.dataset.bound) {
        actionAddPresident.addEventListener('click', () => {
          if (!requireAdminForAction()) return;
          if (!presidentModal || !presidentForm) return;
          presidentForm.reset();
          fillCityInForm(presidentForm);
          const presidentStatus = document.getElementById('presidentModalStatus');
          if (presidentStatus) {
            presidentStatus.textContent = '';
            presidentStatus.style.opacity = '0';
            presidentStatus.style.display = 'none';
          }
          presidentModal.hidden = false;
        });
        actionAddPresident.dataset.bound = '1';
      }

      if (actionAddOffice && !actionAddOffice.dataset.bound) {
        actionAddOffice.addEventListener('click', () => {
          if (!requireAdminForAction()) return;
          if (!officeModal || !officeForm) return;
          officeForm.reset();
          fillCityInForm(officeForm);
          if (officeStatus) {
            officeStatus.style.display = 'none';
            officeStatus.textContent = '';
          }
          officeModal.hidden = false;
        });
        actionAddOffice.dataset.bound = '1';
      }

      if (presidentCancel && !presidentCancel.dataset.bound) {
        presidentCancel.addEventListener('click', () => {
          if (presidentModal) presidentModal.hidden = true;
          if (presidentForm) presidentForm.reset();
          const presidentStatus = document.getElementById('presidentModalStatus');
          if (presidentStatus) {
            presidentStatus.style.display = 'none';
            presidentStatus.style.opacity = '0';
            presidentStatus.textContent = '';
          }
        });
        presidentCancel.dataset.bound = '1';
      }
      if (officeCancel && !officeCancel.dataset.bound) {
        officeCancel.addEventListener('click', () => {
          if (officeModal) officeModal.hidden = true;
          if (officeForm) officeForm.reset();
          if (officeStatus) {
            officeStatus.style.display = 'none';
            officeStatus.textContent = '';
          }
        });
        officeCancel.dataset.bound = '1';
      }

      if (presidentForm) {
        let presidentStatus = document.getElementById('presidentModalStatus');
        if (!presidentStatus) {
          presidentStatus = document.createElement('div');
          presidentStatus.id = 'presidentModalStatus';
          presidentStatus.style.display = 'none';
          presidentStatus.style.opacity = '0';
          presidentStatus.style.transition = 'opacity .25s ease';
          presidentStatus.style.marginTop = '10px';
          presidentStatus.style.padding = '8px 10px';
          presidentStatus.style.borderRadius = '10px';
          presidentStatus.style.fontWeight = '700';
          presidentStatus.style.fontSize = '14px';
          const actions = presidentForm.querySelector('.modalActions');
          if (actions && actions.parentNode) {
            actions.parentNode.insertBefore(presidentStatus, actions.nextSibling);
          } else {
            presidentForm.appendChild(presidentStatus);
          }
        }

        const showPresidentStatus = (text, isError = false, autoHideMs = 2600) => {
          if (!presidentStatus) return;
          if (presidentStatusTimer) {
            clearTimeout(presidentStatusTimer);
            presidentStatusTimer = null;
          }
          presidentStatus.textContent = text;
          presidentStatus.style.display = 'block';
          presidentStatus.style.background = isError ? 'rgba(220,38,38,0.08)' : 'rgba(22,163,74,0.10)';
          presidentStatus.style.border = isError ? '1px solid rgba(220,38,38,0.25)' : '1px solid rgba(22,163,74,0.25)';
          presidentStatus.style.color = isError ? '#b91c1c' : '#166534';
          requestAnimationFrame(() => { presidentStatus.style.opacity = '1'; });

          if (autoHideMs > 0) {
            presidentStatusTimer = setTimeout(() => {
              presidentStatus.style.opacity = '0';
              setTimeout(() => {
                presidentStatus.style.display = 'none';
                presidentStatus.textContent = '';
              }, 250);
            }, autoHideMs);
          }
        };

        if (!presidentForm.dataset.bound) {
          presidentForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!isAdminLoggedIn) {
              showPresidentStatus('يجب تسجيل الدخول كمسؤول', true);
              return;
            }
            const formData = new FormData(presidentForm);
            const payloadRaw = Object.fromEntries(formData.entries());
            const citySource = selectedCity || currentAdminProvince || payloadRaw.city || '';
            const canonicalCity = normalizeCityToArabic(citySource) || String(payloadRaw.city || '').trim();
            const selectedUnion = getSelectedUnionForCity(canonicalCity || citySource);
            const cityForApi = String(getCityForApi(canonicalCity || citySource, selectedUnion) || '').trim();
            const payload = {
              id: Date.now().toString(),
              city: cityForApi,
              union_name: selectedUnion,
              president_name: String(payloadRaw.president_name || '').trim(),
              phone: String(payloadRaw.phone || '').trim(),
              notes: appendUnionMarkerToNotes(payloadRaw.notes, selectedUnion),
            };

            if (!payload.city || !payload.president_name) {
              showPresidentStatus('الرجاء تعبئة الحقول المطلوبة', true);
              return;
            }

            try {
              try {
                let duplicateCheck = await queryPresidentsByCity(cityForApi, !presidentsReadActionUnsupported);
                if (duplicateCheck?.unsupportedAction) {
                  presidentsReadActionUnsupported = true;
                  duplicateCheck = await queryPresidentsByCity(cityForApi, false);
                }
                const rows = Array.isArray(duplicateCheck?.presidents) ? duplicateCheck.presidents : [];
                const cityLoose = normalizeArabicLoose(normalizeCityToArabic(payload.city) || payload.city);
                const targetUnionLoose = normalizeArabicLoose(payload.union_name || '');
                const hasExistingPresident = rows.some((row) => {
                  const rowCityRaw = row?.city || row?.City || '';
                  const rowCityLoose = normalizeArabicLoose(normalizeCityToArabic(rowCityRaw) || rowCityRaw);
                  if (rowCityLoose !== cityLoose) return false;
                  if (normalizeCityToArabic(payload.city) === ISTANBUL_CITY && targetUnionLoose) {
                    const rowUnionLoose = normalizeArabicLoose(getUnionFromRecord(row));
                    if (rowUnionLoose !== targetUnionLoose) return false;
                  }
                  const rowPresidentName = String(row?.president_name || row?.President_Name || '').trim();
                  return rowPresidentName.length > 0;
                });

                if (hasExistingPresident) {
                  showPresidentStatus('يوجد رئيس اتحاد مضاف لهذه المدينة بالفعل', true, 0);
                  return;
                }
              } catch (dupErr) {
                console.warn('Duplicate check skipped due read error:', dupErr);
              }

              const result = await addPresident(payload);
              if (!result || result.ok !== true) {
                throw new Error('API_LOGICAL_ERROR');
              }

              bulkPresidentsState.fetchedAt = Date.now();
              bulkPresidentsState.rows = [
                ...bulkPresidentsState.rows,
                {
                  city: payload.city,
                  union_name: payload.union_name,
                  president_name: payload.president_name,
                  phone: payload.phone,
                  notes: payload.notes,
                },
              ];
              hydratePresidentsCacheFromRows(bulkPresidentsState.rows);

              invalidatePresidentCacheForCity(payload.city, payload.union_name);

              showPresidentStatus('تمت إضافة رئيس الاتحاد بنجاح', false, 2600);
              const cityToRefresh = normalizeCityToArabic(selectedCity || currentAdminProvince || payload.city || '');
              if (cityToRefresh) await updatePresidentForCity(cityToRefresh);
            } catch (err) {
              console.error('President submit failed:', err);
              const apiError = String((err && (err.apiError || err.message)) || '').trim();
              if (
                apiError === 'CITY_ALREADY_HAS_PRESIDENT'
                || /already\s*has\s*president/i.test(apiError)
                || /يوجد\s*رئيس\s*اتحاد/i.test(apiError)
              ) {
                showPresidentStatus('يوجد رئيس اتحاد مضاف لهذه المدينة بالفعل', true, 0);
              } else {
                showPresidentStatus('فشل الحفظ — تحقق من رابط الـAPI وصلاحيات النشر', true, 0);
              }
            }
          });
          presidentForm.dataset.bound = '1';
        }
      }

      if (officeForm) {
        const showOfficeStatus = (text, isError = false) => {
          if (!officeStatus) return;
          officeStatus.textContent = text;
          officeStatus.style.display = 'block';
          officeStatus.style.color = isError ? '#b91c1c' : '#166534';
        };

        if (!officeForm.dataset.bound) {
          officeForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!isAdminLoggedIn) {
              showOfficeStatus('يجب تسجيل الدخول كمسؤول', true);
              return;
            }
            const formData = new FormData(officeForm);
            const cityValue = normalizeCityToArabic(selectedCity || currentAdminProvince || formData.get('city') || '');
            const selectedUnion = getSelectedUnionForCity(cityValue);
            const cityForApi = getCityForApi(cityValue, selectedUnion) || cityValue;
            const payload = {
              city: String(cityForApi || '').trim(),
              union_name: selectedUnion,
              office_name: String(formData.get('office_name') || '').trim(),
              office_manager: String(formData.get('office_manager') || '').trim(),
              contact_number: String(formData.get('contact_number') || '').trim(),
              notes: appendUnionMarkerToNotes(formData.get('notes'), selectedUnion),
            };

            if (!payload.city || !payload.office_name) {
              showOfficeStatus('الرجاء تعبئة الحقول المطلوبة', true);
              return;
            }

            const saveBtn = document.getElementById('officeSave');
            if (saveBtn) {
              saveBtn.disabled = true;
              saveBtn.classList.add('is-loading');
              saveBtn.textContent = 'جارٍ الحفظ...';
            }

            showOfficeStatus('جارٍ الحفظ...');

            try {
              await addOffice(payload);

              const cityToRender = normalizeCityToArabic(cityValue || payload.city || '');
              const existing = getOfficesForCity(cityToRender);
              const normalizedOffice = {
                ...payload,
                city: cityToRender,
              };

              officesByCity.set(cityToRender, [...existing, normalizedOffice]);
              setCachedValue(officesByCity, officesCacheMeta, cityToRender, officesByCity.get(cityToRender));
              renderOfficesSection('success', cityToRender, filterRowsBySelectedUnion(cityToRender, officesByCity.get(cityToRender) || []));

              bulkOfficesState.fetchedAt = Date.now();
              bulkOfficesState.rows = [...bulkOfficesState.rows, { ...payload }];

              showOfficeStatus('تمت إضافة المكتب بنجاح');

              setTimeout(() => {
                if (officeModal) officeModal.hidden = true;
                officeForm.reset();
                if (officeStatus) {
                  officeStatus.style.display = 'none';
                  officeStatus.textContent = '';
                }
              }, 900);
            } catch (err) {
              console.error('Office submit failed:', err);
              const exactError = String(err?.apiError || err?.message || '').trim();
              const isReachability = /API unreachable|CORS\/permission|Failed to fetch/i.test(exactError);
              if (isReachability) {
                showOfficeStatus('تعذر الوصول إلى الواجهة (CORS/الصلاحيات). تحقق من نشر Apps Script بصلاحية عامة ورابط /exec', true);
              } else if (exactError) {
                showOfficeStatus(`فشل حفظ المكتب: ${exactError}`, true);
              } else {
                showOfficeStatus('فشل حفظ المكتب: خطأ غير معروف من الخادم', true);
              }
            } finally {
              if (saveBtn) {
                saveBtn.disabled = false;
                saveBtn.classList.remove('is-loading');
                saveBtn.textContent = 'حفظ';
              }
            }
          });
          officeForm.dataset.bound = '1';
        }
      }

      if (addCancel && !addCancel.dataset.bound) {
        addCancel.addEventListener('click', () => {
          if (addModal) addModal.hidden = true;
          if (addForm) addForm.reset();
          if (addMsg) {
            addMsg.style.display = 'none';
            addMsg.textContent = '';
          }
        });
        addCancel.dataset.bound = '1';
      }

      if (addForm) {
        if (!addForm.dataset.bound) {
          addForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!isAdminLoggedIn) { alert('يجب تسجيل الدخول كمسؤول'); return; }
          const submitBtn = document.getElementById('addSave') || addForm.querySelector('button[type="submit"]');
          const formData = new FormData(addForm);
          const ev = Object.fromEntries(formData.entries());
          const citySource = ev.city || selectedCity || currentAdminProvince || '';
          const canonicalCity = normalizeCityToArabic(citySource) || String(ev.city || '').trim();
          const selectedUnion = getSelectedUnionForCity(canonicalCity || citySource);
          ev.city = getCityForApi(canonicalCity || citySource, selectedUnion) || canonicalCity || ev.city;
          ev.union_name = selectedUnion;
          ev.notes = appendUnionMarkerToNotes(ev.notes, selectedUnion);
          // validate required fields: event_id, title, city, date
          if (!ev.event_id || !ev.title || !ev.city || !ev.date) {
            addMsg.style.display = 'block'; addMsg.style.color = '#b91c1c'; addMsg.textContent = 'الرجاء تعبئة الحقول المطلوبة';
            return;
          }
          // disable submit and show saving text
          if (submitBtn) {
            submitBtn.disabled = true;
            submitBtn.classList.add('is-loading');
            submitBtn.textContent = 'جارٍ الحفظ...';
          }
          addMsg.style.display = 'block'; addMsg.style.color = '#0f172a'; addMsg.textContent = 'جارٍ الحفظ...';
          try {
            await addEvent(ev);
            addMsg.style.color = '#16a34a'; addMsg.textContent = 'تمت إضافة الفعالية بنجاح';

            const cityToRefresh = normalizeCityToArabic(selectedCity || currentAdminProvince || ev.city || '');
            if (cityToRefresh) {
              const existingEvents = getEventsForCity(cityToRefresh);
              const normalizedEvent = {
                ...ev,
                city: cityToRefresh,
              };
              eventsByCity.set(cityToRefresh, [...existingEvents, normalizedEvent]);
              try { showEvents(cityToRefresh, lastManager); } catch (e) { console.warn('Could not refresh events view', e); }

              fetchEvents(cityToRefresh).then(() => {
                try { showEvents(cityToRefresh, lastManager); } catch (e) { console.warn('Could not refresh events view after sync', e); }
              }).catch((syncErr) => {
                console.warn('Background events sync failed:', syncErr);
              });
            }

            // close modal after short delay
            setTimeout(() => {
              if (addModal) addModal.hidden = true;
              addForm.reset();
              addMsg.style.display = 'none';
              addMsg.textContent = '';
              if (submitBtn) {
                submitBtn.disabled = false;
                submitBtn.classList.remove('is-loading');
                submitBtn.textContent = 'حفظ';
              }
            }, 900);
          } catch (err) {
            addMsg.style.color = '#b91c1c'; addMsg.textContent = 'حدث خطأ، حاول مرة أخرى';
            if (submitBtn) {
              submitBtn.disabled = false;
              submitBtn.classList.remove('is-loading');
              submitBtn.textContent = 'حفظ';
            }
            console.error('addForm submit error', err);
            // show more detail for debugging
            if (err && err.message) {
              addMsg.textContent = 'حدث خطأ: ' + err.message;
            }
          }
        });
          addForm.dataset.bound = '1';
        }
      }
})().catch((err) => {
  console.error(err);
  alert("حدث خطأ أثناء تحميل البيانات أو الخريطة. افتح Console للتفاصيل.");
});

(function initActivityStatsSection() {
  const statsEls = {
    cities: document.getElementById('statCities'),
    events: document.getElementById('statEvents'),
  };

  const hasStatsUi = [statsEls.cities, statsEls.events].every(Boolean);
  if (!hasStatsUi) return;

  const setStat = (el, value) => {
    if (!el) return;
    const n = Number(value);
    el.textContent = Number.isFinite(n) ? n.toLocaleString('en-US') : '0';
  };

  const toArray = (value) => (Array.isArray(value) ? value : []);
  let citiesCountCache = null;
  let refreshTimer = null;

  const countFromRegions = async () => {
    if (typeof citiesCountCache === 'number') return citiesCountCache;
    try {
      const regions = await loadJSON(PATHS.regions);
      const bag = new Set();
      toArray(regions).forEach((region) => {
        toArray(region?.cities).forEach((city) => {
          const cityKey = normalizeCityToArabic(city);
          if (cityKey) bag.add(cityKey);
        });
      });
      citiesCountCache = bag.size;
      return citiesCountCache;
    } catch {
      citiesCountCache = 0;
      return 0;
    }
  };

  const countEventsFromMemory = () => {
    let total = 0;
    eventsByCity.forEach((rows) => {
      total += Array.isArray(rows) ? rows.length : 0;
    });
    return total;
  };

  const refreshStats = async () => {
    const cities = await countFromRegions();
    let eventsCount = countEventsFromMemory();

    if (eventsCount === 0) {
      try {
        const res = await fetch(`${APPS_SCRIPT_URL}?action=events`, { method: 'GET', redirect: 'follow' });
        if (res.ok) {
          const json = await res.json();
          eventsCount = toArray(json?.events).length;
        }
      } catch {
        // keep memory-derived value
      }
    }

    setStat(statsEls.cities, cities);
    setStat(statsEls.events, eventsCount);
  };

  const requestRefreshSoon = (delay = 500) => {
    if (refreshTimer) clearTimeout(refreshTimer);
    refreshTimer = setTimeout(() => {
      refreshStats().catch((err) => {
        console.warn('Stats refresh failed:', err);
      });
    }, delay);
  };

  const bindRefreshOnSubmit = (formId) => {
    const form = document.getElementById(formId);
    if (!form || form.dataset.statsBound === '1') return;
    form.addEventListener('submit', () => requestRefreshSoon(800));
    form.dataset.statsBound = '1';
  };

  const boot = () => {
    setStat(statsEls.cities, 0);
    setStat(statsEls.events, 0);

    refreshStats().catch((err) => {
      console.warn('Initial stats load failed:', err);
    });

    bindRefreshOnSubmit('addEventForm');
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot, { once: true });
  } else {
    boot();
  }
})();
