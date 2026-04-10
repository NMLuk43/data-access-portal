/* Built from src/app.js */
const LEVELS = [
  { label: "Gold", minAmount: 5000, className: "pill-level-gold", rangeLabel: "$5,000 and above" },
  { label: "Platinum", minAmount: 1000, className: "pill-level-platinum", rangeLabel: "$1,000 to $4,999" },
  { label: "Silver", minAmount: 500, className: "pill-level-silver", rangeLabel: "$500 to $999" },
  { label: "Bronze", minAmount: 100, className: "pill-level-bronze", rangeLabel: "$100 to $499" },
  { label: "Foundational", minAmount: 0, className: "pill-level-foundational", rangeLabel: "Under $100" }
];

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
const monthOverviewTitle = document.querySelector("#month-overview-title");
const monthTotalLabel = document.querySelector("#month-total-label");
const monthSelector = document.querySelector("#month-selector");

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
const levelGuideOpen = document.querySelector("#level-guide-open");
const levelGuideModal = document.querySelector("#level-guide-modal");
const levelGuideClose = document.querySelector("#level-guide-close");
const levelGuideList = document.querySelector("#level-guide-list");

let donors = [];
const donorByGuid = new Map();
let selectedOverviewPeriod = "";

function getFetchUrl() {
  const url = new URL(window.location.href);
  url.search = "";
  url.searchParams.set("cmd", "fetch-donors");
  return url.toString();
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

function getGivingLevel(amount) {
  return LEVELS.find((level) => amount >= level.minAmount) || LEVELS[LEVELS.length - 1];
}

function getGivingLevelBadgeMarkup(level) {
  return `<span class="pill-level ${level.className}" title="${escapeHtml(level.rangeLabel)}">${level.label}</span>`;
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
  const latestAmount = latestGift ? latestGift.amount : 0;
  const personGuid = row.person_guid || personInfo.person_guid || "";
  const preferredName = getRecordValue(personInfo, ["Person Preferred", "person_preferred"]);
  const lastName = getRecordValue(personInfo, ["Person Last", "person_last"]);
  const status = getRecordValue(personInfo, ["Person Status", "person_status"]) || row.person_status || "Unknown";
  const donorStatus =
    getRecordValue(personInfo, ["Person Donor Status", "donor_status", "person_donor_status"]) ||
    row.donor_status ||
    "Unknown";
  const contactRestriction =
    getRecordValue(personInfo, ["Person Contact Restriction", "contact_restriction", "person_contact_restriction"]) ||
    row.contact_restriction;

  const donor = {
    personGuid,
    preferredName: preferredName || row.person_preferred || "",
    lastName: lastName || row.person_last || "",
    fullName: [preferredName || row.person_preferred, lastName || row.person_last].filter(Boolean).join(" ").trim() || "Unknown Donor",
    status,
    donorStatus,
    address: formatAddress(getPrimaryAddress(row, personInfo)),
    contactRestrictions: normalizeContactRestrictions(contactRestriction),
    gifts,
    latestGift,
    level: getGivingLevel(latestAmount)
  };

  if (personGuid) {
    donorByGuid.set(personGuid, donor);
  }

  return donor;
}

function getAvailableOverviewPeriods() {
  const monthKeys = new Set();
  const yearKeys = new Set();

  for (const donor of donors) {
    for (const gift of donor.gifts) {
      if (isQualifyingGift(gift) && gift.date) {
        monthKeys.add(getMonthKey(gift.date));
        yearKeys.add(getYearKey(gift.date));
      }
    }
  }

  monthKeys.add(getMonthKey(new Date()));
  yearKeys.add(getYearKey(new Date()));

  return {
    months: [...monthKeys].sort((left, right) => right.localeCompare(left)),
    years: [...yearKeys].sort((left, right) => right.localeCompare(left))
  };
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

  const haystack = [donor.fullName, donor.status, donor.donorStatus, donor.address, donor.level.label]
    .join(" ")
    .toLowerCase();

  return haystack.includes(searchTerm);
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
        <td class="data-td" colspan="7">No donors match the current filters.</td>
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
      const donorStatusBadge = isAnonymousDisplay ? "--" : getDonorStatusBadgeMarkup(donor.donorStatus);
      const fundName = donor.latestMatchingGift.fundName || "--";
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
          <td class="data-td">${escapeHtml(donorSegment)}</td>
          <td class="data-td">${donorStatusBadge}</td>
          <td class="data-td">${escapeHtml(fundName)}</td>
          <td class="data-td">${formatDate(donor.latestMatchingGift.date)}</td>
          <td class="data-td">${giftTypeLabel}</td>
          <td class="data-td">${getGivingLevelBadgeMarkup(donor.level)}</td>
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
    selectedOverviewPeriod = getOverviewMonthValue(getMonthKey(new Date()));
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
      const level = getGivingLevel(gift.amount);
      const fundName = gift.fundName || "--";

      return `
        <tr>
          <td class="data-td">${formatDate(gift.date)}</td>
          <td class="data-td">${escapeHtml(fundName)}</td>
          <td class="data-td">${getGivingLevelBadgeMarkup(level)}</td>
          <td class="data-td">${escapeHtml(giftType)}</td>
          <td class="data-td">${escapeHtml(method)}</td>
        </tr>
      `;
    })
    .join("");
}

function openDonorModal(personGuid) {
  const donor = donorByGuid.get(personGuid);

  if (!donor) {
    return;
  }

  const givingSummary = getGivingSummary(donor);
  const latestGiftType = donor.latestGift.recurring ? "Recurring" : donor.latestGift.type || "One-Time";

  donorModalTitle.textContent = donor.fullName;
  modalDonorBadges.innerHTML = donor.contactRestrictions.length > 0 ? getRestrictionBadgeMarkup() : "";
  modalDonorStatus.textContent = `${donor.status} · ${donor.donorStatus}`;
  modalDonorAddress.textContent = donor.address;
  modalLatestGift.textContent = `${donor.level.label} on ${formatDate(donor.latestGift.date)}`;
  modalLatestGiftMeta.textContent = `${latestGiftType} gift via ${donor.latestGift.method || "Unknown method"}`;
  modalLatestGiftFund.textContent = donor.latestGift.fundName || "Fund unavailable";
  modalGivingSummary.textContent = `${donor.level.label} donor`;
  modalGivingMeta.textContent = `${givingSummary.count} gifts on record · ${givingSummary.recurringCount} recurring gifts`;
  renderModalHistoryRows(donor);

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

  donorModal.classList.add("is-open");
  donorModal.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";
}

function closeDonorModal() {
  donorModal.classList.remove("is-open");
  donorModal.setAttribute("aria-hidden", "true");
  document.body.style.overflow =
    levelGuideModal.classList.contains("is-open") || designationModal.classList.contains("is-open") ? "hidden" : "";
}

function renderLevelGuide() {
  levelGuideList.innerHTML = LEVELS.map((level) => `
    <div class="level-guide-item">
      <div class="level-guide-item-copy">
        <div class="level-guide-item-label">${escapeHtml(level.label)}</div>
        <div class="level-guide-item-range">${escapeHtml(level.rangeLabel)}</div>
      </div>
      ${getGivingLevelBadgeMarkup(level)}
    </div>
  `).join("");
}

function openLevelGuideModal() {
  levelGuideModal.classList.add("is-open");
  levelGuideModal.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";
}

function closeLevelGuideModal() {
  levelGuideModal.classList.remove("is-open");
  levelGuideModal.setAttribute("aria-hidden", "true");
  document.body.style.overflow =
    donorModal.classList.contains("is-open") || designationModal.classList.contains("is-open") ? "hidden" : "";
}

function openDesignationModal() {
  renderDesignationModal();
  designationModal.classList.add("is-open");
  designationModal.setAttribute("aria-hidden", "false");
  document.body.style.overflow = "hidden";
}

function closeDesignationModal() {
  designationModal.classList.remove("is-open");
  designationModal.setAttribute("aria-hidden", "true");
  document.body.style.overflow =
    donorModal.classList.contains("is-open") || levelGuideModal.classList.contains("is-open") ? "hidden" : "";
}

function bindFilters() {
  const rerender = () => {
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
  levelGuideOpen.addEventListener("click", openLevelGuideModal);
  levelGuideClose.addEventListener("click", closeLevelGuideModal);
  designationModalOpen.addEventListener("click", openDesignationModal);
  designationModalClose.addEventListener("click", closeDesignationModal);

  levelGuideModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-level-guide-close='true']")) {
      closeLevelGuideModal();
    }
  });

  designationModal.addEventListener("click", (event) => {
    if (event.target.closest("[data-designation-modal-close='true']")) {
      closeDesignationModal();
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && donorModal.classList.contains("is-open")) {
      closeDonorModal();
    }

    if (event.key === "Escape" && levelGuideModal.classList.contains("is-open")) {
      closeLevelGuideModal();
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
  designationSummary.textContent = "Unable to load designation activity.";
  designationList.innerHTML = `
    <article class="designation-item designation-item-empty">
      <p class="designation-item-name">Designation data unavailable</p>
      <p class="designation-item-meta">The monthly designation summary could not be loaded.</p>
    </article>
  `;

  donorResults.innerHTML = `
    <tr>
      <td class="data-td" colspan="7">Unable to load donors right now.</td>
    </tr>
  `;

  alertsFeed.innerHTML = `
    <article class="alert-item">
      <div class="alert-item-row">
        <span class="alert-item-title">Data unavailable</span>
        <span class="alert-badge alert-badge-lapsed">Error</span>
      </div>
      <p class="alert-item-reason">The donor feed could not be loaded.</p>
    </article>
  `;
}

async function loadDonors() {
  renderLoadingState();
  donorByGuid.clear();

  const response = await fetch(getFetchUrl(), {
    credentials: "same-origin"
  });

  if (!response.ok) {
    throw new Error(`Request failed with status ${response.status}`);
  }

  const payload = await response.json();
  const rows = Array.isArray(payload.row) ? payload.row : [];
  donors = rows.map(normalizeDonor).filter((donor) => donor.latestGift);
  selectedOverviewPeriod = getOverviewMonthValue(getMonthKey(new Date()));
}

async function init() {
  bindFilters();
  renderLoadingState();
  renderLevelGuide();

  try {
    await loadDonors();
    renderOverviewMonthSelector();
    renderCurrentMonthLabels();
    renderCurrentMonthStats();
    renderDesignationOverview();
    renderAlerts();
    renderDashboard();
    dashboardRoot.classList.remove("dashboard-loading");
  } catch (error) {
    console.error(error);
    renderErrorState();
  }
}

init();
