/* Built from src/app.js */
const THANK_YOU_WINDOW_DAYS = 14;
const donorSearchInput = document.querySelector(".search-input");
const dateRangeFilter = document.querySelector("#date-range-filter");
const giftTypeFilter = document.querySelector("#gift-type-filter");
const segmentFilter = document.querySelector("#segment-filter");
const sortFilter = document.querySelector("#sort-filter");
const donorResults = document.querySelector("#donor-results");
const alertsFeed = document.querySelector("#alerts-feed");
const resultsChip = document.querySelector("#results-chip");
const dateChip = document.querySelector("#date-chip");
const sortChip = document.querySelector("#sort-chip");
const fuzzyChip = document.querySelector("#fuzzy-chip");
const monthOverviewTitle = document.querySelector("#month-overview-title");
const monthTotalLabel = document.querySelector("#month-total-label");
const monthSelector = document.querySelector("#month-selector");
const summaryPeriodLabel = document.querySelector("#summary-period-label");
const summaryPeriodTotal = document.querySelector("#summary-period-total");
const summaryRecurringDonors = document.querySelector("#summary-recurring-donors");

const monthTotal = document.querySelector("#month-total");
const monthTotalDetail = document.querySelector("#month-total-detail");
const recurringTotal = document.querySelector("#recurring-total");
const recurringNote = document.querySelector("#recurring-note");
const oneTimeTotal = document.querySelector("#one-time-total");
const oneTimeNote = document.querySelector("#one-time-note");
const averageGift = document.querySelector("#average-gift");
const averageGiftNote = document.querySelector("#average-gift-note");
const largestGift = document.querySelector("#largest-gift");
const largestGiftNote = document.querySelector("#largest-gift-note");
const giftMixChart = document.querySelector("#gift-mix-chart");
const giftMixCenter = document.querySelector("#gift-mix-center");
const legendRecurring = document.querySelector("#legend-recurring");
const legendOneTime = document.querySelector("#legend-one-time");
const trendChartTitle = document.querySelector("#trend-chart-title");
const trendLabels = [
  document.querySelector("#trend-label-1"),
  document.querySelector("#trend-label-2"),
  document.querySelector("#trend-label-3"),
  document.querySelector("#trend-label-4")
];
const designationSummary = document.querySelector("#designation-summary");
const designationList = document.querySelector("#designation-list");
const designationModal = document.querySelector("#designation-modal");
const designationModalOpen = document.querySelector("#designation-modal-open");
const designationModalClose = document.querySelector("#designation-modal-close");
const designationModalTitle = document.querySelector("#designation-modal-title");
const designationModalSubtitle = document.querySelector("#designation-modal-subtitle");
const designationModalList = document.querySelector("#designation-modal-list");
const weekBars = [
  document.querySelector("#trend-bar-week1"),
  document.querySelector("#trend-bar-week2"),
  document.querySelector("#trend-bar-week3"),
  document.querySelector("#trend-bar-week4")
];
const firstRepeatSummary = document.querySelector("#first-repeat-summary");
const firstRepeatBars = document.querySelector("#first-repeat-bars");
const yoySummary = document.querySelector("#yoy-summary");
const yoyCurrentPeriod = document.querySelector("#yoy-current-period");
const yoyPriorPeriod = document.querySelector("#yoy-prior-period");
const yoyDeltaValue = document.querySelector("#yoy-delta-value");
const benchmarkDeltaValue = document.querySelector("#benchmark-delta-value");
const statusMixSummary = document.querySelector("#status-mix-summary");
const statusMixList = document.querySelector("#status-mix-list");
const dashboardRoot = document.querySelector("#dashboard-root");
const donorModal = document.querySelector("#donor-modal");
const donorModalClose = document.querySelector("#donor-modal-close");
const donorModalTitle = document.querySelector("#donor-modal-title");
const modalDonorBadges = document.querySelector("#modal-donor-badges");
const modalDonorStatus = document.querySelector("#modal-donor-status");
const modalDonorAddress = document.querySelector("#modal-donor-address");
const modalLatestGift = document.querySelector("#modal-latest-gift");
const modalLatestGiftMeta = document.querySelector("#modal-latest-gift-meta");
const modalLatestGiftFund = document.querySelector("#modal-latest-gift-fund");
const modalGivingSummary = document.querySelector("#modal-giving-summary");
const modalGivingMeta = document.querySelector("#modal-giving-meta");
const modalHistoryCount = document.querySelector("#modal-history-count");
const modalHistoryRows = document.querySelector("#modal-history-rows");
const modalAlertCard = document.querySelector("#modal-alert-card");
const modalAlertContent = document.querySelector("#modal-alert-content");
const modalRestrictionsCard = document.querySelector("#modal-restrictions-card");
const modalRestrictionsContent = document.querySelector("#modal-restrictions-content");
const modalActionList = document.querySelector("#modal-action-list");
const modalNotesTextarea = document.querySelector("#modal-notes-textarea");
const modalSaveNote = document.querySelector("#modal-save-note");
const modalActivityList = document.querySelector("#modal-activity-list");
const modalTimelineList = document.querySelector("#modal-timeline-list");
const modalMarkThanked = document.querySelector("#modal-mark-thanked");

let donors = [];
const donorByGuid = new Map();
let selectedOverviewPeriod = "";
let lastLoadErrorMessage = "";
let activeModal = null;
const stewardshipState = new Map();

function shouldUseLocalSampleData() {
  const pageUrl = new URL(window.location.href);
  const hostname = pageUrl.hostname.toLowerCase();

  if (pageUrl.searchParams.get("sampleData") === "1") {
    return true;
  }

  if (pageUrl.protocol === "file:") {
    return true;
  }

  if (hostname === "localhost" || hostname === "127.0.0.1" || hostname === "[::1]") {
    return true;
  }

  if (
    hostname.endsWith(".local") ||
    hostname.endsWith(".localhost") ||
    hostname.startsWith("192.168.") ||
    hostname.startsWith("10.") ||
    hostname.startsWith("172.")
  ) {
    return true;
  }

  return false;
}

function getFetchUrl() {
  if (window.__PORTAL_DATA_URL__) {
    return new URL(window.__PORTAL_DATA_URL__, window.location.href).toString();
  }

  const pageUrl = new URL(window.location.href);
  const configuredDataUrl = pageUrl.searchParams.get("dataUrl");

  if (configuredDataUrl) {
    return new URL(configuredDataUrl, pageUrl.href).toString();
  }

  if (shouldUseLocalSampleData()) {
    return new URL("./sample-data.json", pageUrl.href).toString();
  }

  const url = new URL(window.location.href);
  url.search = "";
  url.searchParams.set("cmd", "fetch-donors");
  return url.toString();
}

function getAnonymousFetchUrl() {
  if (window.__PORTAL_ANONYMOUS_DATA_URL__) {
    return new URL(window.__PORTAL_ANONYMOUS_DATA_URL__, window.location.href).toString();
  }

  if (shouldUseLocalSampleData()) {
    return "";
  }

  const pageUrl = new URL(window.location.href);
  const configuredAnonymousUrl = pageUrl.searchParams.get("anonymousDataUrl");

  if (configuredAnonymousUrl) {
    return new URL(configuredAnonymousUrl, pageUrl.href).toString();
  }

  const url = new URL(window.location.href);
  url.search = "";
  url.searchParams.set("cmd", "fetch-anonymous");
  return url.toString();
}

function getFetchOptions(fetchUrl) {
  const fetchOrigin = new URL(fetchUrl, window.location.href).origin;
  const currentOrigin = window.location.origin;
  const credentialsMode = fetchOrigin === currentOrigin ? "same-origin" : "include";

  return { credentials: credentialsMode };
}

function wait(milliseconds) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

async function fetchJsonWithRetries(fetchUrl, label, attempts = 3) {
  let lastError = null;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetch(fetchUrl, getFetchOptions(fetchUrl));

      if (!response.ok) {
        throw new Error(`${label} request failed with status ${response.status}`);
      }

      const rawText = await response.text();

      if (!rawText.trim()) {
        throw new Error(`${label} request returned an empty response`);
      }

      try {
        return JSON.parse(rawText);
      } catch (_error) {
        const bodyPreview = rawText.slice(0, 80).trim();
        throw new Error(`${label} data URL returned non-JSON content: ${fetchUrl} (${bodyPreview || "empty response"})`);
      }
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));

      if (attempt < attempts) {
        await wait(350 * attempt);
      }
    }
  }

  throw lastError || new Error(`${label} request failed`);
}

function parseAmount(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  if (!value) {
    return null;
  }

  const parsed = new Date(`${value}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeForSearch(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function isFuzzyMatch(haystack, needle) {
  const normalizedHaystack = normalizeForSearch(haystack);
  const normalizedNeedle = normalizeForSearch(needle);

  if (!normalizedNeedle) {
    return true;
  }

  if (normalizedHaystack.includes(normalizedNeedle)) {
    return true;
  }

  let haystackIndex = 0;

  for (const character of normalizedNeedle) {
    haystackIndex = normalizedHaystack.indexOf(character, haystackIndex);

    if (haystackIndex === -1) {
      return false;
    }

    haystackIndex += 1;
  }

  return true;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0
  }).format(value);
}

function formatDate(value) {
  const date = typeof value === "string" ? parseDate(value) : value;

  if (!date) {
    return "--";
  }

  return new Intl.DateTimeFormat("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric"
  }).format(date);
}

function formatDeltaCurrency(value) {
  if (!Number.isFinite(value)) {
    return "--";
  }

  if (value === 0) {
    return "$0";
  }

  return `${value > 0 ? "+" : "-"}${formatCurrency(Math.abs(value))}`;
}

function getMonthKey(referenceDate) {
  const year = referenceDate.getFullYear();
  const month = String(referenceDate.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function getYearKey(referenceDate) {
  return String(referenceDate.getFullYear());
}

function getMonthStartFromKey(monthKey) {
  const [year, month] = monthKey.split("-").map(Number);
  return new Date(year, month - 1, 1);
}

function getOverviewMonthValue(monthKey) {
  return `month:${monthKey}`;
}

function getOverviewYearValue(yearKey) {
  return `year:${yearKey}`;
}

function parseOverviewPeriod(periodValue) {
  const [type, key] = String(periodValue || "").split(":");

  if ((type === "month" || type === "year") && key) {
    return { type, key };
  }

  return { type: "month", key: getMonthKey(new Date()) };
}

function getPeriodReferenceDate(periodValue) {
  const { type, key } = parseOverviewPeriod(periodValue);
  const now = new Date();

  if (type === "month") {
    const monthStart = getMonthStartFromKey(key);
    const isCurrentMonth = getMonthKey(now) === key;

    if (isCurrentMonth) {
      return now;
    }

    return new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0, 23, 59, 59, 999);
  }

  const isCurrentYear = getYearKey(now) === key;

  if (isCurrentYear) {
    return now;
  }

  return new Date(Number(key), 11, 31, 23, 59, 59, 999);
}

function formatMonthLabel(monthKey) {
  return new Intl.DateTimeFormat("en-US", {
    month: "long",
    year: "numeric"
  }).format(getMonthStartFromKey(monthKey));
}

function formatYearLabel(yearKey) {
  return yearKey;
}

function formatOverviewPeriodLabel(periodValue) {
  const { type, key } = parseOverviewPeriod(periodValue);
  return type === "year" ? formatYearLabel(key) : formatMonthLabel(key);
}

function giftMatchesOverviewPeriod(gift, periodValue) {
  const { type, key } = parseOverviewPeriod(periodValue);

  if (type === "year") {
    return getYearKey(gift.date) === key;
  }

  return getMonthKey(gift.date) === key;
}

function formatAddress(address) {
  if (!address || Object.keys(address).length === 0) {
    return "Address unavailable";
  }

  const street = address["Address by Type, Rank Street Combined"] || "";
  const city = address["Address by Type, Rank City"] || "";
  const region = address["Address by Type, Rank Region"] || "";
  const postal = address["Address by Type, Rank Postal"] || "";

  return [street, [city, region].filter(Boolean).join(", "), postal]
    .filter(Boolean)
    .join(" ");
}

function getRestrictionBadgeMarkup() {
  return `<span class="alert-badge alert-badge-restriction">Contact Restriction</span>`;
}

function normalizeContactRestrictions(value) {
  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value
      .flatMap((entry) => normalizeContactRestrictions(entry))
      .filter(Boolean);
  }

  if (typeof value === "string") {
    return value
      .split(/[\n;,]+/)
      .map((entry) => entry.trim())
      .filter(Boolean);
  }

  if (typeof value === "object") {
    return Object.values(value)
      .map((entry) => String(entry).trim())
      .filter(Boolean);
  }

  return [String(value).trim()].filter(Boolean);
}

function isRecurringGift(gift) {
  return String(gift.recurring) === "1";
}

function isQualifyingGift(gift) {
  return gift && gift.date && gift.amount > 0 && gift.type !== "Pledge";
}

function isPastOrPresentGift(gift, today = new Date()) {
  return Boolean(gift.date) && gift.date <= today;
}

function getDateDifferenceInDays(laterDate, earlierDate) {
  const millisecondsPerDay = 1000 * 60 * 60 * 24;
  return Math.floor((laterDate - earlierDate) / millisecondsPerDay);
}

function getDonorStatusBadgeMarkup(donorStatus) {
  const normalizedStatus = String(donorStatus || "Unknown").trim();
  const statusMap = {
    Active: "donor-status-badge-active",
    "At Risk": "donor-status-badge-at-risk",
    Lapsed: "donor-status-badge-lapsed"
  };
  const className = statusMap[normalizedStatus] || "donor-status-badge-unknown";

  return `<span class="donor-status-badge ${className}">${escapeHtml(normalizedStatus || "Unknown")}</span>`;
}

function normalizeGift(gift) {
  return {
    amount: parseAmount(gift["Gifts Amount"]),
    date: parseDate(gift["Gifts Date"]),
    dateRaw: gift["Gifts Date"] || "",
    anonymous: Boolean(gift["Gifts Anonymous"]),
    fundName: gift["Funds Name"] || "",
    recurring: isRecurringGift(gift),
    type: gift["Gifts Type"] || "",
    method: gift["Gifts Gift Method"] || "",
    paymentStatus: gift["Gifts Online Payment Status"] || ""
  };
}

function parseArrayLikeValue(value) {
  if (Array.isArray(value)) {
    return value;
  }

  if (value && typeof value === "object") {
    return Object.values(value);
  }

  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value);
      if (Array.isArray(parsed)) {
        return parsed;
      }

      if (parsed && typeof parsed === "object") {
        return Object.values(parsed);
      }

      return [];
    } catch (_error) {
      return [];
    }
  }

  return [];
}

function getRecordValue(record, keys) {
  if (!record || typeof record !== "object") {
    return "";
  }

  for (const key of keys) {
    if (record[key] !== undefined && record[key] !== null && String(record[key]).trim() !== "") {
      return record[key];
    }
  }

  const normalizedEntries = Object.entries(record).map(([key, value]) => [key.trim().toLowerCase(), value]);

  for (const key of keys) {
    const match = normalizedEntries.find(([entryKey, value]) => {
      return entryKey === key.trim().toLowerCase() && value !== undefined && value !== null && String(value).trim() !== "";
    });

    if (match) {
      return match[1];
    }
  }

  return "";
}

function getPrimaryPersonInfo(row) {
  const personInfo = parseArrayLikeValue(row.person_info);

  if (personInfo.length > 0) {
    return personInfo[0] || {};
  }

  return row;
}

function getPrimaryAddress(row, personInfo) {
  const cofoAddress = parseArrayLikeValue(personInfo.cofo_address);

  if (cofoAddress.length > 0) {
    return cofoAddress[0] || {};
  }

  const nestedAddress = parseArrayLikeValue(personInfo.address);

  if (nestedAddress.length > 0) {
    return nestedAddress[0] || {};
  }

  const mailingAddress = parseArrayLikeValue(row.mailing_address);

  if (mailingAddress.length > 0) {
    return mailingAddress[0] || {};
  }

  return {};
}

function getMostRecentGift(gifts, today = new Date()) {
  const validGifts = gifts.filter(isQualifyingGift).filter((gift) => isPastOrPresentGift(gift, today));

  validGifts.sort((left, right) => right.date - left.date || right.amount - left.amount);

  return validGifts[0] || null;
}

function normalizeDonor(row) {
  const personInfo = getPrimaryPersonInfo(row);
  const gifts = parseArrayLikeValue(row.giving_history).map(normalizeGift);
  const latestGift = getMostRecentGift(gifts);
  const preferredName = getRecordValue(personInfo, ["Person Preferred", "person_preferred"]);
  const lastName = getRecordValue(personInfo, ["Person Last", "person_last"]);
  const organizationName = getRecordValue(personInfo, [
    "Companies and Foundations Name",
    "companies_and_foundations_name",
    "Company Name",
    "company_name"
  ]);
  const status =
    getRecordValue(personInfo, ["Person Status", "person_status"]) ||
    row.person_status ||
    (organizationName ? "Company/Foundation" : "Unknown");
  const donorStatus =
    getRecordValue(personInfo, ["Person Donor Status", "donor_status", "person_donor_status"]) ||
    row.donor_status ||
    "Unknown";
  const contactRestriction =
    getRecordValue(personInfo, ["Person Contact Restriction", "contact_restriction", "person_contact_restriction"]) ||
    row.contact_restriction;
  const fullName =
    [preferredName || row.person_preferred, lastName || row.person_last].filter(Boolean).join(" ").trim() ||
    organizationName ||
    "Unknown Donor";
  const fallbackPersonGuid = [
    fullName,
    latestGift ? latestGift.dateRaw : "",
    latestGift ? latestGift.fundName : "",
    latestGift ? latestGift.amount : ""
  ]
    .filter(Boolean)
    .join("::");
  const personGuid = row.person_guid || personInfo.person_guid || fallbackPersonGuid;

  const donor = {
    personGuid,
    preferredName: preferredName || row.person_preferred || "",
    lastName: lastName || row.person_last || "",
    fullName,
    status,
    donorStatus,
    address: formatAddress(getPrimaryAddress(row, personInfo)),
    contactRestrictions: normalizeContactRestrictions(contactRestriction),
    gifts,
    latestGift
  };

  donorByGuid.set(personGuid, donor);

  return donor;
}

function normalizeAnonymousDonor(row, index) {
  const giftsSource = Array.isArray(row.anonymous_gifts)
    ? row.anonymous_gifts
    : Array.isArray(row.giving_history)
      ? row.giving_history
      : Array.isArray(row.gifts)
        ? row.gifts
        : [row];
  const gifts = giftsSource.map((gift) => normalizeGift({
    ...gift,
    "Gifts Anonymous": true
  }));
  const latestGift = getMostRecentGift(gifts);

  if (!latestGift) {
    return null;
  }

  const personGuid = `anonymous::${latestGift.dateRaw || "unknown-date"}::${latestGift.fundName || "unknown-fund"}::${latestGift.amount}::${index}`;
  const donor = {
    personGuid,
    preferredName: "",
    lastName: "",
    fullName: "Anonymous Donor",
    status: "Anonymous",
    donorStatus: "Private",
    address: "Identity withheld",
    contactRestrictions: [],
    gifts,
    latestGift
  };

  donorByGuid.set(personGuid, donor);
  return donor;
}

function getAvailableOverviewPeriods() {
  const currentYear = new Date().getFullYear();
  const allowedYears = new Set([
    String(currentYear),
    String(currentYear - 1),
    String(currentYear - 2)
  ]);
  const monthKeys = new Set();
  const yearKeys = new Set();

  for (const donor of donors) {
    for (const gift of donor.gifts) {
      if (isQualifyingGift(gift) && gift.date) {
        const yearKey = getYearKey(gift.date);

        if (!allowedYears.has(yearKey)) {
          continue;
        }

        monthKeys.add(getMonthKey(gift.date));
        yearKeys.add(yearKey);
      }
    }
  }

  const currentMonthKey = getMonthKey(new Date());

  return {
    months: [...monthKeys]
      .filter((monthKey) => monthKey <= currentMonthKey)
      .sort((left, right) => right.localeCompare(left)),
    years: [...yearKeys].sort((left, right) => right.localeCompare(left))
  };
}

function getDefaultOverviewPeriod() {
  const periods = getAvailableOverviewPeriods();

  if (periods.months.length > 0) {
    return getOverviewMonthValue(periods.months[0]);
  }

  if (periods.years.length > 0) {
    return getOverviewYearValue(periods.years[0]);
  }

  return getOverviewYearValue(String(new Date().getFullYear() - 1));
}

function isSameMonth(date, referenceDate) {
  return (
    date &&
    date.getFullYear() === referenceDate.getFullYear() &&
    date.getMonth() === referenceDate.getMonth()
  );
}

function getQuarterStart(referenceDate) {
  const quarterMonth = Math.floor(referenceDate.getMonth() / 3) * 3;
  return new Date(referenceDate.getFullYear(), quarterMonth, 1);
}

function giftMatchesDateRange(gift, rangeValue, today) {
  if (!gift.date || !isPastOrPresentGift(gift, today)) {
    return false;
  }

  if (rangeValue === "all-time") {
    return true;
  }

  if (rangeValue === "this-month") {
    return isSameMonth(gift.date, today);
  }

  if (rangeValue === "last-30") {
    return getDateDifferenceInDays(today, gift.date) <= 30;
  }

  if (rangeValue === "quarter-to-date") {
    return gift.date >= getQuarterStart(today) && gift.date <= today;
  }

  return true;
}

function donorMatchesSearch(donor, searchTerm) {
  if (!searchTerm) {
    return true;
  }

  const haystack = [donor.fullName, donor.status, donor.donorStatus, donor.address, formatCurrency(donor.latestGift.amount)].join(" ");

  return haystack.toLowerCase().includes(searchTerm) || isFuzzyMatch(haystack, searchTerm);
}

function getFilteredDonors() {
  const today = new Date();
  const searchTerm = donorSearchInput.value.trim().toLowerCase();
  const selectedDateRange = dateRangeFilter.value;
  const selectedGiftType = giftTypeFilter.value;
  const selectedSegment = segmentFilter.value;

  return donors
    .map((donor) => {
      const matchingGifts = donor.gifts
        .filter(isQualifyingGift)
        .filter((gift) => giftMatchesDateRange(gift, selectedDateRange, today))
        .filter((gift) => {
          if (selectedGiftType === "all") {
            return true;
          }

          return selectedGiftType === "recurring" ? gift.recurring : !gift.recurring;
        });

      const latestMatchingGift = getMostRecentGift(matchingGifts, today);

      return {
        ...donor,
        filteredGifts: matchingGifts,
        latestMatchingGift
      };
    })
    .filter((donor) => donor.latestMatchingGift)
    .filter((donor) => selectedSegment === "all" || donor.status === selectedSegment)
    .filter((donor) => donorMatchesSearch(donor, searchTerm));
}

function sortDonors(items) {
  const sorted = [...items];

  sorted.sort((left, right) => {
    if (sortFilter.value === "date-asc") {
      return left.latestMatchingGift.date - right.latestMatchingGift.date;
    }

    if (sortFilter.value === "amount-desc") {
      return right.latestMatchingGift.amount - left.latestMatchingGift.amount;
    }

    if (sortFilter.value === "name-asc") {
      return left.fullName.localeCompare(right.fullName);
    }

    return right.latestMatchingGift.date - left.latestMatchingGift.date;
  });

  return sorted;
}

function renderDonorRows(items) {
  if (!donorResults) {
    return;
  }

  if (items.length === 0) {
    donorResults.innerHTML = `
      <tr>
        <td class="data-td" colspan="4">No donors match the current filters.</td>
      </tr>
    `;
    return;
  }

  donorResults.innerHTML = items
    .map((donor) => {
      const isAnonymousDisplay = donor.latestMatchingGift.anonymous;
      const giftTypeLabel = donor.latestMatchingGift.recurring ? "Recurring" : "One-Time";
      const donorName = isAnonymousDisplay ? "Anonymous Donor" : donor.fullName;
      const donorSubtext = isAnonymousDisplay ? "Gift marked anonymous" : donor.address;
      const donorSegment = isAnonymousDisplay ? "Private" : donor.status;
      const rowAttributes = isAnonymousDisplay
        ? `class="donor-row donor-row-disabled"`
        : `class="donor-row" data-person-guid="${escapeHtml(donor.personGuid)}"`;

      return `
        <tr ${rowAttributes}>
          <td class="data-td">
            <div class="donor-name">${escapeHtml(donorName)}</div>
            ${!isAnonymousDisplay && donor.contactRestrictions.length > 0 ? `<div class="donor-inline-flags">${getRestrictionBadgeMarkup()}</div>` : ""}
            <div class="donor-subtext">${escapeHtml(donorSubtext)}</div>
          </td>
          <td class="data-td">${formatDate(donor.latestMatchingGift.date)}</td>
          <td class="data-td">${formatCurrency(donor.latestMatchingGift.amount)}</td>
          <td class="data-td">${escapeHtml(donorSegment)}</td>
        </tr>
      `;
    })
    .join("");
}

function renderFilterChips(count) {
  const dateLabels = {
    "this-month": "This Month",
    "last-30": "Last 30 Days",
    "quarter-to-date": "Quarter to Date",
    "all-time": "All Time"
  };

  const sortLabels = {
    "date-desc": "Sort by Date",
    "date-asc": "Sort by Oldest Date",
    "amount-desc": "Sort by Gift Amount",
    "name-asc": "Sort by Name"
  };

  dateChip.textContent = `Range: ${dateLabels[dateRangeFilter.value] || "This Month"}`;
  sortChip.textContent = sortLabels[sortFilter.value] || "Sort by Date";
  resultsChip.textContent = `${count} Results`;
}

function getSelectedOverviewMonthStats() {
  const periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()));
  const referenceDate = getPeriodReferenceDate(periodValue);
  const { type } = parseOverviewPeriod(periodValue);
  const monthGifts = donors
    .flatMap((donor) => donor.gifts)
    .filter(isQualifyingGift)
    .filter((gift) => giftMatchesOverviewPeriod(gift, periodValue))
    .filter((gift) => isPastOrPresentGift(gift, referenceDate));

  const total = monthGifts.reduce((sum, gift) => sum + gift.amount, 0);
  const recurring = monthGifts
    .filter((gift) => gift.recurring)
    .reduce((sum, gift) => sum + gift.amount, 0);
  const oneTime = total - recurring;
  const largest = monthGifts.reduce((max, gift) => Math.max(max, gift.amount), 0);
  const average = monthGifts.length > 0 ? total / monthGifts.length : 0;

  const periodTotals = [0, 0, 0, 0];

  for (const gift of monthGifts) {
    const bucketIndex =
      type === "year"
        ? Math.min(Math.floor(gift.date.getMonth() / 3), 3)
        : Math.min(Math.floor((gift.date.getDate() - 1) / 7), 3);
    periodTotals[bucketIndex] += gift.amount;
  }

  return {
    type,
    total,
    recurring,
    oneTime,
    average,
    largest,
    giftCount: monthGifts.length,
    periodTotals
  };
}

function renderCurrentMonthStats() {
  const stats = getSelectedOverviewMonthStats();
  const recurringPercent = stats.total > 0 ? Math.round((stats.recurring / stats.total) * 100) : 0;
  const oneTimePercent = stats.total > 0 ? 100 - recurringPercent : 0;
  const monthlyTarget = 150000;
  const targetPercent = monthlyTarget > 0 ? Math.round((stats.total / monthlyTarget) * 100) : 0;
  const periodLabel = formatOverviewPeriodLabel(selectedOverviewPeriod);

  monthTotal.textContent = formatCurrency(stats.total);
  monthTotalDetail.textContent = `${targetPercent}% of the ${formatCurrency(monthlyTarget)} target across ${stats.giftCount} gifts in ${periodLabel}.`;
  recurringTotal.textContent = formatCurrency(stats.recurring);
  recurringNote.textContent = `${recurringPercent}% of period giving`;
  oneTimeTotal.textContent = formatCurrency(stats.oneTime);
  oneTimeNote.textContent = `${oneTimePercent}% of period giving`;
  averageGift.textContent = formatCurrency(stats.average);
  averageGiftNote.textContent = `Across ${stats.giftCount} gifts in ${periodLabel}`;
  largestGift.textContent = formatCurrency(stats.largest);
  largestGiftNote.textContent = `Largest single gift in ${periodLabel}`;

  giftMixCenter.textContent = `${recurringPercent}%`;
  legendRecurring.textContent = formatCurrency(stats.recurring);
  legendOneTime.textContent = formatCurrency(stats.oneTime);
  giftMixChart.style.background = `conic-gradient(#003865 0 ${recurringPercent}%, #cbd5e1 ${recurringPercent}% 100%)`;

  trendChartTitle.textContent = stats.type === "year" ? "Quarterly Pace" : "Weekly Pace";
  const trendLabelText = stats.type === "year" ? ["Q1", "Q2", "Q3", "Q4"] : ["Week 1", "Week 2", "Week 3", "Week 4"];
  trendLabels.forEach((label, index) => {
    label.textContent = trendLabelText[index];
  });

  const peakPeriod = Math.max(...stats.periodTotals, 1);

  stats.periodTotals.forEach((value, index) => {
    const heightPercent = peakPeriod > 0 ? Math.max(12, Math.round((value / peakPeriod) * 100)) : 12;
    weekBars[index].style.height = `${heightPercent}%`;
  });
}

function getSelectedMonthFundStats() {
  const periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()));
  const referenceDate = getPeriodReferenceDate(periodValue);
  const monthlyGifts = donors
    .flatMap((donor) => donor.gifts)
    .filter(isQualifyingGift)
    .filter((gift) => giftMatchesOverviewPeriod(gift, periodValue))
    .filter((gift) => isPastOrPresentGift(gift, referenceDate));

  const fundTotals = new Map();

  for (const gift of monthlyGifts) {
    const fundName = gift.fundName || "Unassigned Fund";
    const currentStats = fundTotals.get(fundName) || { name: fundName, total: 0, giftCount: 0 };
    currentStats.total += gift.amount;
    currentStats.giftCount += 1;
    fundTotals.set(fundName, currentStats);
  }

  const rankedFunds = [...fundTotals.values()].sort((left, right) => right.total - left.total);

  const scholarshipTotal = rankedFunds
    .filter((fund) => fund.name.toLowerCase().includes("scholarship"))
    .reduce((sum, fund) => sum + fund.total, 0);

  return {
    monthLabel: formatOverviewPeriodLabel(periodValue),
    rankedFunds,
    totalFunds: rankedFunds.length,
    totalGifts: monthlyGifts.length,
    scholarshipTotal
  };
}

function getOverviewGiftsForPeriod(periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()))) {
  const referenceDate = getPeriodReferenceDate(periodValue);

  return donors
    .flatMap((donor) => donor.gifts.map((gift) => ({ donor, gift })))
    .filter(({ gift }) => isQualifyingGift(gift))
    .filter(({ gift }) => giftMatchesOverviewPeriod(gift, periodValue))
    .filter(({ gift }) => isPastOrPresentGift(gift, referenceDate));
}

function getComparisonPeriodValue(periodValue) {
  const { type, key } = parseOverviewPeriod(periodValue);

  if (type === "year") {
    return getOverviewYearValue(String(Number(key) - 1));
  }

  const [year, month] = key.split("-").map(Number);
  return getOverviewMonthValue(`${year - 1}-${String(month).padStart(2, "0")}`);
}

function getPeriodTotal(periodValue) {
  return getOverviewGiftsForPeriod(periodValue).reduce((sum, entry) => sum + entry.gift.amount, 0);
}

function getRecurringDonorCount(periodValue) {
  return new Set(
    getOverviewGiftsForPeriod(periodValue)
      .filter(({ gift }) => gift.recurring)
      .map(({ donor }) => donor.personGuid)
  ).size;
}

function getYearOverYearStats(periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()))) {
  const currentTotal = getPeriodTotal(periodValue);
  const comparisonPeriod = getComparisonPeriodValue(periodValue);
  const priorTotal = getPeriodTotal(comparisonPeriod);
  const delta = currentTotal - priorTotal;
  const percentDelta = priorTotal > 0 ? Math.round((delta / priorTotal) * 100) : null;
  const benchmarkDelta = currentTotal - 150000;

  return {
    currentTotal,
    priorTotal,
    delta,
    percentDelta,
    benchmarkDelta,
    comparisonLabel: formatOverviewPeriodLabel(comparisonPeriod)
  };
}

function getFirstTimeRepeatStats(periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()))) {
  const entries = getOverviewGiftsForPeriod(periodValue);
  const donorBuckets = new Map();

  for (const { donor, gift } of entries) {
    const current = donorBuckets.get(donor.personGuid) || {
      donor,
      priorGiftCount: donor.gifts.filter((entry) => isQualifyingGift(entry) && entry.date && gift.date && entry.date < gift.date).length,
      amount: 0
    };
    current.amount += gift.amount;
    donorBuckets.set(donor.personGuid, current);
  }

  let firstTimeCount = 0;
  let repeatCount = 0;
  let firstTimeAmount = 0;
  let repeatAmount = 0;

  for (const stats of donorBuckets.values()) {
    if (stats.priorGiftCount === 0) {
      firstTimeCount += 1;
      firstTimeAmount += stats.amount;
    } else {
      repeatCount += 1;
      repeatAmount += stats.amount;
    }
  }

  return { firstTimeCount, repeatCount, firstTimeAmount, repeatAmount };
}

function getPersonStatusMix(periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()))) {
  const donorMap = new Map();

  for (const { donor } of getOverviewGiftsForPeriod(periodValue)) {
    const current = donorMap.get(donor.personGuid) || { label: donor.status || "Unknown", count: 0 };
    current.count += 1;
    donorMap.set(donor.personGuid, current);
  }

  const statusTotals = new Map();

  for (const entry of donorMap.values()) {
    statusTotals.set(entry.label, (statusTotals.get(entry.label) || 0) + 1);
  }

  return [...statusTotals.entries()]
    .map(([label, count]) => ({ label, count }))
    .sort((left, right) => right.count - left.count);
}

function getStewardshipState(personGuid) {
  if (!stewardshipState.has(personGuid)) {
    stewardshipState.set(personGuid, {
      thanked: false,
      lastActionDate: "",
      notes: "",
      history: []
    });
  }

  return stewardshipState.get(personGuid);
}

function addStewardshipHistory(personGuid, label) {
  const state = getStewardshipState(personGuid);
  state.lastActionDate = new Date().toISOString();
  state.history.unshift({
    label,
    timestamp: state.lastActionDate
  });
}

function renderSummaryBar() {
  const periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()));
  const stats = getSelectedOverviewMonthStats();

  summaryPeriodLabel.textContent = formatOverviewPeriodLabel(periodValue);
  summaryPeriodTotal.textContent = formatCurrency(stats.total);
  summaryRecurringDonors.textContent = String(getRecurringDonorCount(periodValue));
}

function renderInsightCards() {
  const periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()));
  const firstRepeat = getFirstTimeRepeatStats(periodValue);
  const yoy = getYearOverYearStats(periodValue);
  const statusMix = getPersonStatusMix(periodValue);
  const totalMixAmount = Math.max(firstRepeat.firstTimeAmount + firstRepeat.repeatAmount, 1);

  firstRepeatSummary.textContent = `${firstRepeat.firstTimeCount} first-time donors and ${firstRepeat.repeatCount} repeat donors in ${formatOverviewPeriodLabel(periodValue)}.`;
  firstRepeatBars.innerHTML = [
    {
      label: "First-Time",
      amount: firstRepeat.firstTimeAmount,
      count: firstRepeat.firstTimeCount,
      percent: Math.round((firstRepeat.firstTimeAmount / totalMixAmount) * 100)
    },
    {
      label: "Repeat",
      amount: firstRepeat.repeatAmount,
      count: firstRepeat.repeatCount,
      percent: Math.round((firstRepeat.repeatAmount / totalMixAmount) * 100)
    }
  ]
    .map(
      (item) => `
        <div class="comparison-row">
          <div class="comparison-row-copy">
            <p class="comparison-row-label">${item.label}</p>
            <p class="comparison-row-meta">${item.count} donor${item.count === 1 ? "" : "s"} · ${formatCurrency(item.amount)}</p>
          </div>
          <div class="comparison-row-bar"><span style="width:${Math.max(item.percent, 8)}%"></span></div>
          <p class="comparison-row-value">${item.percent}%</p>
        </div>
      `
    )
    .join("");

  yoySummary.textContent = `${formatOverviewPeriodLabel(periodValue)} compared with ${yoy.comparisonLabel}.`;
  yoyCurrentPeriod.textContent = formatCurrency(yoy.currentTotal);
  yoyPriorPeriod.textContent = formatCurrency(yoy.priorTotal);
  yoyDeltaValue.textContent = formatDeltaCurrency(yoy.delta);
  benchmarkDeltaValue.textContent = formatDeltaCurrency(yoy.benchmarkDelta);

  const totalStatuses = statusMix.reduce((sum, item) => sum + item.count, 0);
  statusMixSummary.textContent = `${totalStatuses} donor records contributed in the selected period.`;
  statusMixList.innerHTML = statusMix
    .map(
      (item) => `
        <div class="status-mix-item">
          <div class="status-mix-copy">
            <p class="status-mix-label">${escapeHtml(item.label)}</p>
            <p class="status-mix-meta">${item.count} donor${item.count === 1 ? "" : "s"}</p>
          </div>
          <div class="status-mix-bar"><span style="width:${Math.max(Math.round((item.count / Math.max(totalStatuses, 1)) * 100), 10)}%"></span></div>
        </div>
      `
    )
    .join("");
}

function renderDesignationOverview() {
  const stats = getSelectedMonthFundStats();
  const topFunds = stats.rankedFunds.slice(0, 5);

  if (stats.totalFunds === 0) {
    designationSummary.textContent = `No qualifying designation activity recorded for ${stats.monthLabel}.`;
    designationList.innerHTML = `
      <article class="designation-item designation-item-empty">
        <p class="designation-item-name">No designations recorded</p>
        <p class="designation-item-meta">There were no qualifying gifts with designation data in ${stats.monthLabel}.</p>
      </article>
    `;
    return;
  }

  const scholarshipText =
    stats.scholarshipTotal > 0
      ? `Scholarship-designated giving totaled ${formatCurrency(stats.scholarshipTotal)}.`
      : "No scholarship-designated giving recorded.";

  designationSummary.textContent = `${stats.totalFunds} active designations across ${stats.totalGifts} qualifying gifts in ${stats.monthLabel}. ${scholarshipText}`;
  designationList.innerHTML = topFunds
    .map(
      (fund, index) => `
        <article class="designation-item">
          <div class="designation-item-copy">
            <p class="designation-item-rank">#${index + 1}</p>
            <p class="designation-item-name">${escapeHtml(fund.name)}</p>
            <p class="designation-item-meta">${fund.giftCount} gift${fund.giftCount === 1 ? "" : "s"} in ${stats.monthLabel}</p>
          </div>
          <p class="designation-item-value">${formatCurrency(fund.total)}</p>
        </article>
      `
    )
    .join("");
}

function renderDesignationModal() {
  const stats = getSelectedMonthFundStats();

  designationModalTitle.textContent = `${stats.monthLabel} Designations`;

  if (stats.totalFunds === 0) {
    designationModalSubtitle.textContent = `No qualifying designation activity recorded for ${stats.monthLabel}.`;
    designationModalList.innerHTML = `
      <article class="designation-modal-item designation-item-empty">
        <p class="designation-item-name">No designations recorded</p>
        <p class="designation-item-meta">There were no qualifying gifts with designation data in ${stats.monthLabel}.</p>
      </article>
    `;
    return;
  }

  designationModalSubtitle.textContent = `${stats.totalFunds} active designations across ${stats.totalGifts} qualifying gifts.`;
  designationModalList.innerHTML = stats.rankedFunds
    .map(
      (fund, index) => `
        <article class="designation-modal-item">
          <div class="designation-item-copy">
            <p class="designation-item-rank">#${index + 1}</p>
            <p class="designation-item-name">${escapeHtml(fund.name)}</p>
            <p class="designation-item-meta">${fund.giftCount} gift${fund.giftCount === 1 ? "" : "s"}</p>
          </div>
          <p class="designation-item-value">${formatCurrency(fund.total)}</p>
        </article>
      `
    )
    .join("");
}

function buildAlerts() {
  const periodValue = selectedOverviewPeriod || getOverviewMonthValue(getMonthKey(new Date()));
  const referenceDate = getPeriodReferenceDate(periodValue);
  const alerts = [];

  for (const donor of donors) {
    const latestGift = getMostRecentGift(donor.gifts, referenceDate);

    if (!latestGift || latestGift.anonymous) {
      continue;
    }

    const daysSinceLatestGift = getDateDifferenceInDays(referenceDate, latestGift.date);

    if (daysSinceLatestGift <= THANK_YOU_WINDOW_DAYS && giftMatchesOverviewPeriod(latestGift, periodValue)) {
      alerts.push({
        personGuid: donor.personGuid,
        donorName: donor.fullName,
        reason: "Needs a thank you",
        badgeLabel: "Thank You",
        badgeClassName: "alert-badge-thanks",
        sortDate: latestGift.date
      });
    }

    if (donor.donorStatus === "At Risk") {
      alerts.push({
        personGuid: donor.personGuid,
        donorName: donor.fullName,
        reason: "Donor is marked at risk",
        badgeLabel: "At Risk",
        badgeClassName: "alert-badge-at-risk",
        sortDate: latestGift.date
      });
    }
  }

  alerts.sort((left, right) => right.sortDate - left.sortDate);

  return alerts;
}

function renderAlerts() {
  const alerts = buildAlerts();

  if (alerts.length === 0) {
    alertsFeed.innerHTML = `
      <article class="alert-item">
        <div class="alert-item-row">
          <span class="alert-item-title">No current alerts</span>
        </div>
        <p class="alert-item-reason">No donors currently match the thank-you or at-risk rules.</p>
      </article>
    `;
    return;
  }

  const alertMarkup = alerts
    .map(
      (alert) => `
        <article class="alert-item">
          <button class="alert-item-row alert-item-button" data-person-guid="${escapeHtml(alert.personGuid)}" type="button">
            <span class="alert-item-title">${escapeHtml(alert.donorName)}</span>
            <span class="donor-inline-flags">
              ${donorByGuid.get(alert.personGuid)?.contactRestrictions.length > 0 ? getRestrictionBadgeMarkup() : ""}
              <span class="alert-badge ${alert.badgeClassName}">${alert.badgeLabel}</span>
            </span>
          </button>
          <p class="alert-item-reason">${escapeHtml(alert.reason)}</p>
        </article>
      `
    )
    .join("");

  alertsFeed.innerHTML = alertMarkup;

  const alertItems = [...alertsFeed.querySelectorAll(".alert-item")];

  if (alertItems.length === 0 || alertsFeed.scrollHeight <= alertsFeed.clientHeight) {
    return;
  }

  const moreElement = document.createElement("p");
  moreElement.className = "alerts-more";
  moreElement.textContent = "+0 More";
  alertsFeed.appendChild(moreElement);

  const availableHeight = alertsFeed.clientHeight;
  const moreHeight = moreElement.offsetHeight;
  let visibleCount = alertItems.length;

  for (let index = 0; index < alertItems.length; index += 1) {
    const item = alertItems[index];
    const itemBottom = item.offsetTop + item.offsetHeight;

    if (itemBottom + moreHeight > availableHeight) {
      visibleCount = index;
      break;
    }
  }

  if (visibleCount <= 0) {
    visibleCount = 1;
  }

  alertItems.forEach((item, index) => {
    item.hidden = index >= visibleCount;
  });

  const remainingAlerts = Math.max(0, alerts.length - visibleCount);

  if (remainingAlerts > 0) {
    moreElement.textContent = `+${remainingAlerts} More`;
  } else {
    moreElement.remove();
  }
}

function getAlertsForDonor(personGuid) {
  return buildAlerts().filter((alert) => alert.personGuid === personGuid);
}

function renderDashboard() {
  const filtered = sortDonors(getFilteredDonors());
  renderDonorRows(filtered);
  renderFilterChips(filtered.length);
}

function renderOverviewMonthSelector() {
  const periods = getAvailableOverviewPeriods();

  if (!selectedOverviewPeriod) {
    selectedOverviewPeriod = getDefaultOverviewPeriod();
  }

  const monthOptions = periods.months
    .map((monthKey) => {
      const value = getOverviewMonthValue(monthKey);
      const selected = value === selectedOverviewPeriod ? " selected" : "";
      return `<option value="${value}"${selected}>${escapeHtml(formatMonthLabel(monthKey))}</option>`;
    })
    .join("");

  const yearOptions = periods.years
    .map((yearKey) => {
      const value = getOverviewYearValue(yearKey);
      const selected = value === selectedOverviewPeriod ? " selected" : "";
      return `<option value="${value}"${selected}>${escapeHtml(formatYearLabel(yearKey))}</option>`;
    })
    .join("");

  monthSelector.innerHTML = `
    <optgroup label="Months">
      ${monthOptions}
    </optgroup>
    <optgroup label="Years">
      ${yearOptions}
    </optgroup>
  `;
}

function renderCurrentMonthLabels() {
  const periodLabel = formatOverviewPeriodLabel(selectedOverviewPeriod);
  monthOverviewTitle.textContent = `${periodLabel} Giving Overview`;
  monthTotalLabel.textContent = `${periodLabel} Total`;
}

function getGivingSummary(donor) {
  const qualifyingGifts = donor.gifts.filter(isQualifyingGift);
  const total = qualifyingGifts.reduce((sum, gift) => sum + gift.amount, 0);
  const recurringCount = qualifyingGifts.filter((gift) => gift.recurring).length;
  return {
    total,
    count: qualifyingGifts.length,
    recurringCount
  };
}

function getFocusableElements(container) {
  return [...container.querySelectorAll("button, [href], input, select, textarea, [tabindex]:not([tabindex='-1'])")].filter(
    (element) => !element.hasAttribute("disabled") && !element.getAttribute("aria-hidden")
  );
}

function trapFocus(event, modalElement) {
  if (event.key !== "Tab") {
    return;
  }

  const focusable = getFocusableElements(modalElement);

  if (focusable.length === 0) {
    event.preventDefault();
    return;
  }

  const first = focusable[0];
  const last = focusable[focusable.length - 1];

  if (event.shiftKey && document.activeElement === first) {
    event.preventDefault();
    last.focus();
    return;
  }

  if (!event.shiftKey && document.activeElement === last) {
    event.preventDefault();
    first.focus();
  }
}

function openModal(modalElement) {
  activeModal = modalElement;
  modalElement.classList.add("is-open");
  modalElement.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";
  getFocusableElements(modalElement)[0]?.focus();
}

function closeModal(modalElement) {
  modalElement.classList.remove("is-open");
  modalElement.setAttribute("aria-hidden", "true");
  if (activeModal === modalElement) {
    activeModal = null;
  }
  document.body.style.overflow = donorModal.classList.contains("is-open") || designationModal.classList.contains("is-open") ? "hidden" : "";
}

function renderModalHistoryRows(donor) {
  const recentGifts = donor.gifts
    .filter(isQualifyingGift)
    .sort((left, right) => right.date - left.date || right.amount - left.amount)
    .slice(0, 8);

  modalHistoryCount.textContent = `${recentGifts.length} recent gifts`;

  if (recentGifts.length === 0) {
    modalHistoryRows.innerHTML = `
      <tr>
        <td class="data-td" colspan="5">No recent qualifying gifts.</td>
      </tr>
    `;
    return;
  }

  modalHistoryRows.innerHTML = recentGifts
    .map((gift) => {
      const giftType = gift.recurring ? "Recurring" : gift.type || "One-Time";
      const method = gift.method || "--";
      const fundName = gift.fundName || "--";

      return `
        <tr>
          <td class="data-td">${formatDate(gift.date)}</td>
          <td class="data-td">${escapeHtml(fundName)}</td>
          <td class="data-td">${formatCurrency(gift.amount)}</td>
          <td class="data-td">${escapeHtml(giftType)}</td>
          <td class="data-td">${escapeHtml(method)}</td>
        </tr>
      `;
    })
    .join("");
}

function renderModalStewardship(donor) {
  const state = getStewardshipState(donor.personGuid);
  const donorAlerts = getAlertsForDonor(donor.personGuid);
  const daysSinceGift = donor.latestGift?.date ? getDateDifferenceInDays(new Date(), donor.latestGift.date) : null;
  const suggestedActions = [
    donorAlerts.some((alert) => alert.badgeLabel === "Thank You")
      ? `Send thank-you note for ${formatCurrency(donor.latestGift.amount)} gift`
      : `Review stewardship touchpoint for ${formatOverviewPeriodLabel(selectedOverviewPeriod)}`,
    donor.donorStatus === "Lapsed" ? "Queue reactivation outreach" : "Confirm next cultivation touchpoint",
    donor.contactRestrictions.length > 0 ? "Respect contact restrictions before outreach" : "Prepare outreach channel recommendation"
  ];

  modalActionList.innerHTML = suggestedActions
    .map(
      (action, index) => `
        <button class="modal-task-button" data-modal-task="${index}" type="button">
          <span>${escapeHtml(action)}</span>
          <span class="modal-mock-badge">Mockup</span>
        </button>
      `
    )
    .join("");

  modalNotesTextarea.value = state.notes;
  modalActivityList.innerHTML = [
    ...(state.thanked ? [{ label: "Marked as thanked", timestamp: state.lastActionDate }] : []),
    ...state.history
  ]
    .slice(0, 4)
    .map(
      (item) => `
        <div class="modal-activity-item">
          <p class="modal-activity-title">${escapeHtml(item.label)}</p>
          <p class="modal-activity-meta">${item.timestamp ? new Date(item.timestamp).toLocaleString("en-US") : "Suggested activity"}</p>
        </div>
      `
    )
    .join("") || `<div class="modal-activity-item"><p class="modal-activity-title">No recorded actions yet</p><p class="modal-activity-meta">Mock action history will appear here.</p></div>`;

  modalTimelineList.innerHTML = donor.gifts
    .filter(isQualifyingGift)
    .sort((left, right) => right.date - left.date || right.amount - left.amount)
    .slice(0, 5)
    .map(
      (gift) => `
        <div class="modal-timeline-item">
          <p class="modal-timeline-title">${formatCurrency(gift.amount)} · ${escapeHtml(gift.fundName || "Unassigned Fund")}</p>
          <p class="modal-timeline-meta">${formatDate(gift.date)} · ${gift.recurring ? "Recurring" : gift.type || "One-Time"}</p>
        </div>
      `
    )
    .join("");

  modalMarkThanked.textContent = state.thanked ? "Thanked" : "Mark Thanked";
  modalMarkThanked.disabled = state.thanked;
  modalMarkThanked.dataset.personGuid = donor.personGuid;
  modalSaveNote.dataset.personGuid = donor.personGuid;
  modalActionList.dataset.personGuid = donor.personGuid;

  if (daysSinceGift !== null) {
    modalGivingMeta.textContent = `${modalGivingMeta.textContent} · ${daysSinceGift} day${daysSinceGift === 1 ? "" : "s"} since latest gift`;
  }
}

function openDonorModal(personGuid) {
  const donor = donorByGuid.get(personGuid);

  if (!donor) {
    return;
  }

  const givingSummary = getGivingSummary(donor);
  const latestGiftType = donor.latestGift.recurring ? "Recurring" : donor.latestGift.type || "One-Time";
  const badgeMarkup = [
    getDonorStatusBadgeMarkup(donor.donorStatus),
    donor.contactRestrictions.length > 0 ? getRestrictionBadgeMarkup() : ""
  ]
    .filter(Boolean)
    .join("");

  donorModalTitle.textContent = donor.fullName;
  modalDonorBadges.innerHTML = badgeMarkup;
  modalDonorStatus.textContent = donor.status;
  modalDonorAddress.textContent = donor.address;
  modalLatestGift.textContent = `${formatCurrency(donor.latestGift.amount)} on ${formatDate(donor.latestGift.date)}`;
  modalLatestGiftMeta.textContent = `${latestGiftType} gift via ${donor.latestGift.method || "Unknown method"}`;
  modalLatestGiftFund.textContent = donor.latestGift.fundName || "Fund unavailable";
  modalGivingSummary.textContent = formatCurrency(givingSummary.total);
  modalGivingMeta.textContent = `${givingSummary.count} gifts on record · ${givingSummary.recurringCount} recurring gifts`;
  renderModalHistoryRows(donor);
  renderModalStewardship(donor);

  const donorAlerts = getAlertsForDonor(personGuid);

  if (donorAlerts.length > 0) {
    modalAlertCard.hidden = false;
    modalAlertContent.innerHTML = donorAlerts
      .map(
        (donorAlert) => `
          <div class="alert-item donor-modal-inline-alert">
            <div class="alert-item-row">
              <span class="alert-item-title">${escapeHtml(donorAlert.donorName)}</span>
              <span class="alert-badge ${donorAlert.badgeClassName}">${donorAlert.badgeLabel}</span>
            </div>
            <p class="alert-item-reason">${escapeHtml(donorAlert.reason)}</p>
          </div>
        `
      )
      .join("");
  } else {
    modalAlertCard.hidden = true;
    modalAlertContent.innerHTML = "";
  }

  if (donor.contactRestrictions.length > 0) {
    modalRestrictionsCard.hidden = false;
    modalRestrictionsContent.innerHTML = donor.contactRestrictions
      .map(
        (restriction) => `
          <div class="donor-modal-restriction-item">${escapeHtml(restriction)}</div>
        `
      )
      .join("");
  } else {
    modalRestrictionsCard.hidden = true;
    modalRestrictionsContent.innerHTML = "";
  }

  openModal(donorModal);
}

function closeDonorModal() {
  closeModal(donorModal);
}

function openDesignationModal() {
  renderDesignationModal();
  openModal(designationModal);
}

function closeDesignationModal() {
  closeModal(designationModal);
}

function bindFilters() {
  const rerender = () => {
    renderSummaryBar();
    renderInsightCards();
    renderDashboard();
  };

  let alertResizeFrame = 0;
  const rerenderAlertsForLayout = () => {
    if (alertResizeFrame) {
      cancelAnimationFrame(alertResizeFrame);
    }

    alertResizeFrame = window.requestAnimationFrame(() => {
      renderAlerts();
    });
  };

  donorSearchInput.addEventListener("input", rerender);
  dateRangeFilter.addEventListener("change", rerender);
  giftTypeFilter.addEventListener("change", rerender);
  segmentFilter.addEventListener("change", rerender);
  sortFilter.addEventListener("change", rerender);
  monthSelector.addEventListener("change", () => {
    selectedOverviewPeriod = monthSelector.value;
    renderCurrentMonthLabels();
    renderCurrentMonthStats();
    renderSummaryBar();
    renderInsightCards();
    renderDesignationOverview();
    renderAlerts();

    if (designationModal.classList.contains("is-open")) {
      renderDesignationModal();
    }
  });

  donorResults.addEventListener("click", (event) => {
    const row = event.target.closest("[data-person-guid]");

    if (!row) {
      return;
    }

    openDonorModal(row.dataset.personGuid);
  });

  alertsFeed.addEventListener("click", (event) => {
    const trigger = event.target.closest("[data-person-guid]");

    if (!trigger) {
      return;
    }

    openDonorModal(trigger.dataset.personGuid);
  });

  donorModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-modal-close='true']")) {
      closeDonorModal();
    }
  });

  donorModalClose.addEventListener("click", closeDonorModal);
  designationModalOpen.addEventListener("click", openDesignationModal);
  designationModalClose.addEventListener("click", closeDesignationModal);
  modalMarkThanked.addEventListener("click", () => {
    const personGuid = modalMarkThanked.dataset.personGuid;

    if (!personGuid) {
      return;
    }

    const state = getStewardshipState(personGuid);
    state.thanked = true;
    addStewardshipHistory(personGuid, "Marked thank-you complete");
    renderModalStewardship(donorByGuid.get(personGuid));
  });
  modalSaveNote.addEventListener("click", () => {
    const personGuid = modalSaveNote.dataset.personGuid;

    if (!personGuid) {
      return;
    }

    const state = getStewardshipState(personGuid);
    state.notes = modalNotesTextarea.value.trim();
    addStewardshipHistory(personGuid, state.notes ? "Saved stewardship note" : "Cleared stewardship note");
    renderModalStewardship(donorByGuid.get(personGuid));
  });
  modalActionList.addEventListener("click", (event) => {
    const taskButton = event.target.closest("[data-modal-task]");
    const personGuid = modalActionList.dataset.personGuid;

    if (!taskButton || !personGuid) {
      return;
    }

    addStewardshipHistory(personGuid, `Queued action: ${taskButton.querySelector("span")?.textContent || taskButton.textContent}`);
    renderModalStewardship(donorByGuid.get(personGuid));
  });

  designationModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-designation-modal-close='true']")) {
      closeDesignationModal();
    }
  });

  document.addEventListener("keydown", (event) => {
    if (activeModal) {
      trapFocus(event, activeModal);
    }

    if (event.key === "Escape" && donorModal.classList.contains("is-open")) {
      closeDonorModal();
    }

    if (event.key === "Escape" && designationModal.classList.contains("is-open")) {
      closeDesignationModal();
    }
  });

  window.addEventListener("resize", rerenderAlertsForLayout);
}

function renderLoadingState() {
  dashboardRoot.classList.add("dashboard-loading");
  designationSummary.textContent = "Loading designation activity...";
  designationList.innerHTML = `
    <article class="designation-item">
      <div class="designation-item-copy">
        <span class="skeleton-line skeleton-line-title"></span>
        <span class="skeleton-line skeleton-line-text"></span>
      </div>
      <span class="skeleton-line skeleton-line-badge"></span>
    </article>
    <article class="designation-item">
      <div class="designation-item-copy">
        <span class="skeleton-line skeleton-line-title"></span>
        <span class="skeleton-line skeleton-line-text"></span>
      </div>
      <span class="skeleton-line skeleton-line-badge"></span>
    </article>
  `;

  alertsFeed.innerHTML = `
    <article class="skeleton-row">
      <div class="skeleton-row-top">
        <span class="skeleton-line skeleton-line-title"></span>
        <span class="skeleton-line skeleton-line-badge"></span>
      </div>
      <span class="skeleton-line skeleton-line-text"></span>
    </article>
    <article class="skeleton-row">
      <div class="skeleton-row-top">
        <span class="skeleton-line skeleton-line-title"></span>
        <span class="skeleton-line skeleton-line-badge"></span>
      </div>
      <span class="skeleton-line skeleton-line-text"></span>
    </article>
    <article class="skeleton-row">
      <div class="skeleton-row-top">
        <span class="skeleton-line skeleton-line-title"></span>
        <span class="skeleton-line skeleton-line-badge"></span>
      </div>
      <span class="skeleton-line skeleton-line-text"></span>
    </article>
  `;

  donorResults.innerHTML = `
    <tr>
      <td class="data-td">
        <span class="skeleton-table-line skeleton-table-line-name"></span>
        <span class="skeleton-table-line skeleton-table-line-subtext"></span>
      </td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
    </tr>
    <tr>
      <td class="data-td">
        <span class="skeleton-table-line skeleton-table-line-name"></span>
        <span class="skeleton-table-line skeleton-table-line-subtext"></span>
      </td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
    </tr>
    <tr>
      <td class="data-td">
        <span class="skeleton-table-line skeleton-table-line-name"></span>
        <span class="skeleton-table-line skeleton-table-line-subtext"></span>
      </td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
      <td class="data-td"><span class="skeleton-table-line skeleton-table-line-cell"></span></td>
    </tr>
  `;
}

function renderErrorState() {
  dashboardRoot.classList.remove("dashboard-loading");
  const errorMessage = lastLoadErrorMessage || "The donor feed could not be loaded.";
  designationSummary.textContent = "Unable to load designation activity.";
  designationList.innerHTML = `
    <article class="designation-item designation-item-empty">
      <p class="designation-item-name">Designation data unavailable</p>
      <p class="designation-item-meta">${escapeHtml(errorMessage)}</p>
    </article>
  `;

  donorResults.innerHTML = `
    <tr>
      <td class="data-td" colspan="4">${escapeHtml(errorMessage)}</td>
    </tr>
  `;

  alertsFeed.innerHTML = `
    <article class="alert-item">
      <div class="alert-item-row">
        <span class="alert-item-title">Data unavailable</span>
        <span class="alert-badge alert-badge-lapsed">Error</span>
      </div>
      <p class="alert-item-reason">${escapeHtml(errorMessage)}</p>
    </article>
  `;
}

function applyDonorPayload(payload, anonymousPayload = null) {
  donorByGuid.clear();
  const rows = Array.isArray(payload.row) ? payload.row : [];
  const anonymousRows = Array.isArray(anonymousPayload?.row) ? anonymousPayload.row : [];
  donors = [
    ...rows.map(normalizeDonor),
    ...anonymousRows.map((row, index) => normalizeAnonymousDonor(row, index))
  ].filter((donor) => donor && donor.latestGift);

  const latestGiftDate = donors
    .flatMap((donor) => donor.gifts)
    .filter(isQualifyingGift)
    .reduce((latest, gift) => {
      if (!gift.date) {
        return latest;
      }

      if (!latest || gift.date > latest) {
        return gift.date;
      }

      return latest;
    }, null);

  if (shouldUseLocalSampleData() && latestGiftDate) {
    const latestGiftYear = getYearKey(latestGiftDate);
    const allowedYears = new Set([
      String(new Date().getFullYear()),
      String(new Date().getFullYear() - 1),
      String(new Date().getFullYear() - 2)
    ]);

    selectedOverviewPeriod = allowedYears.has(latestGiftYear)
      ? getOverviewMonthValue(getMonthKey(latestGiftDate))
      : getDefaultOverviewPeriod();

    if (dateRangeFilter) {
      dateRangeFilter.value = "all-time";
    }
  } else {
    selectedOverviewPeriod = getDefaultOverviewPeriod();
  }
}

function renderLoadedDashboard() {
  renderOverviewMonthSelector();
  renderCurrentMonthLabels();
  renderCurrentMonthStats();
  renderSummaryBar();
  renderInsightCards();
  renderDesignationOverview();
  renderAlerts();
  renderDashboard();
  dashboardRoot.classList.remove("dashboard-loading");
}

async function loadDonors() {
  renderLoadingState();
  donorByGuid.clear();
  lastLoadErrorMessage = "";
  const fetchUrl = getFetchUrl();
  const anonymousFetchUrl = getAnonymousFetchUrl();
  const requests = [fetchJsonWithRetries(fetchUrl, "Donor", 3)];

  if (anonymousFetchUrl) {
    requests.push(fetchJsonWithRetries(anonymousFetchUrl, "Anonymous", 3));
  }

  const [payload, anonymousPayload = null] = await Promise.all(requests);

  applyDonorPayload(payload, anonymousPayload);
}

async function init() {
  bindFilters();
  renderLoadingState();

  try {
    await loadDonors();
    renderLoadedDashboard();
  } catch (error) {
    lastLoadErrorMessage = error instanceof Error ? error.message : String(error);
    console.error(error);
    renderErrorState();
  }
}

init();
