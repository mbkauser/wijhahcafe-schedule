/**
 * WIJHAH CAFÉ — AUTOMATIC SCHEDULE GENERATOR
 * ============================================
 * Paste this whole file into Extensions > Apps Script (attached to your
 * Form's response Google Sheet). See SETUP INSTRUCTIONS at the bottom
 * of the chat message for how to connect and deploy this.
 *
 * WHAT IT DOES:
 * 1. Reads every availability response from your Form Responses sheet.
 * 2. For a given month, figures out every calendar date that needs a shift
 *    (Sundays, Wednesdays, Fridays, and the last Thursday only).
 * 3. For each date, builds the pool of people who are available that day
 *    and NOT blocked out on that specific date.
 * 4. Splits that pool by gender (see CONFIG.MALES) — every shift is staffed
 *    entirely by one group, never mixed.
 * 5. Randomly assigns "active" baristas + "backup" baristas + a "lead"
 *    from within the chosen group, weighted toward whoever has worked the
 *    fewest shifts / led the fewest times so far, so it stays fair AND
 *    isn't the exact same lineup every time you regenerate.
 * 6. Publishes the result as JSON your website can fetch directly.
 */

// ────────────────────────────────────────────────────────────────
// CONFIG — edit these to match your real staffing needs
// ────────────────────────────────────────────────────────────────
const CONFIG = {
  // How many people needed per shift type, and what time each one is.
  // dow = JavaScript day-of-week number (0=Sun, 1=Mon ... 6=Sat)
  SHIFTS: {
    Sunday:    { dow: 0, time: '10:30 AM – 1:00 PM',  active: 3, backup: 2, firstWeekOnly: false },
    Wednesday: { dow: 3, time: '7:30 PM – 10:00 PM',  active: 3, backup: 1, firstWeekOnly: false },
    Friday:    { dow: 5, time: '7:00 PM – 10:00 PM',  active: 5, backup: 3, firstWeekOnly: false },
    Thursday:  { dow: 4, time: '6:30 PM – 9:00 PM',   active: 2, backup: 0, lastWeekOnly: true }, // last Thursday of month only
  },

  // Column header keywords used to find your Form Responses columns.
  // If your form's exact wording differs, tweak these keyword fragments —
  // the script searches each header for these substrings (case-insensitive).
  COLUMNS: {
    name: 'name',
    month: 'month',
    // Your form uses ONE combined checkbox question ("Event Availability")
    // rather than a separate column per day. This keyword just needs to
    // match part of that question's header text in your Sheet.
    availability: 'event availability',
    // Separate "Current Volunteering Commitments" checkbox question —
    // used to auto-exclude someone from barista-ing a day they're
    // volunteering at.
    volunteering: 'volunteering',
    blockedDates: 'not available',
  },

  // Names in this list are treated as "male." Everyone else on the team
  // (as they appear in form responses) is treated as "female." Shifts are
  // never mixed — every shift is assigned entirely from one group.
  MALES: ['Umar', 'Taha Butta', 'Abdullah S.'],

  // Name of the sheet tab that holds raw Form Responses
  RESPONSES_SHEET_NAME: 'Form Responses 1',

  // Name of a (new, empty) sheet tab this script will use to store the
  // published JSON output. The script creates this automatically if missing.
  OUTPUT_SHEET_NAME: 'Published Schedule JSON',
};

const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];

// ────────────────────────────────────────────────────────────────
// MENU — adds a button inside the Google Sheet UI
// ────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 Schedule Generator')
    .addItem('Generate & Publish Schedule (next 3 months)', 'generateAndPublish')
    .addToUi();
}

// ────────────────────────────────────────────────────────────────
// MAIN ENTRY POINT
// ────────────────────────────────────────────────────────────────
function generateAndPublish() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responsesSheet = ss.getSheetByName(CONFIG.RESPONSES_SHEET_NAME);
  if (!responsesSheet) {
    throw new Error('Could not find sheet tab named "' + CONFIG.RESPONSES_SHEET_NAME + '". Check CONFIG.RESPONSES_SHEET_NAME matches your actual tab name.');
  }

  const data = responsesSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const col = mapColumns(headers);
  const responses = parseResponses(rows, col);

  // Generate schedule for current month + next 2 months
  const today = new Date();
  const months = [];
  for (let i = 0; i < 3; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() + i, 1);
    months.push({ year: d.getFullYear(), monthIndex: d.getMonth() });
  }

  const output = months.map(m => buildMonthSchedule(m.year, m.monthIndex, responses));

  // Write JSON to an output sheet (acts as your "database" for the web app)
  let outSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);
  if (!outSheet) outSheet = ss.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
  outSheet.clear();
  outSheet.getRange(1, 1).setValue('Generated: ' + new Date().toISOString());
  outSheet.getRange(2, 1).setValue(JSON.stringify(output));

  SpreadsheetApp.getUi().alert('Schedule generated for ' + output.map(o => o.monthName).join(', ') + '. Now deploy/redeploy the Web App if you haven\'t already (see setup instructions).');
}

// ────────────────────────────────────────────────────────────────
// WEB APP ENDPOINT — this is what your website fetches
// ────────────────────────────────────────────────────────────────
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);
  let json = '[]';
  if (outSheet) {
    const val = outSheet.getRange(2, 1).getValue();
    if (val) json = val;
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────────
// HELPERS
// ────────────────────────────────────────────────────────────────
function mapColumns(headers) {
  const find = (kw) => headers.findIndex(h => String(h).toLowerCase().includes(kw));
  return {
    name: find(CONFIG.COLUMNS.name),
    month: find(CONFIG.COLUMNS.month),
    availability: find(CONFIG.COLUMNS.availability),
    volunteering: find(CONFIG.COLUMNS.volunteering),
    blockedDates: find(CONFIG.COLUMNS.blockedDates),
  };
}

// Turns raw sheet rows into: { name, month, availableDays:Set, blockedDates:Set('2026-07-05') }
function parseResponses(rows, col) {
  const responses = [];
  rows.forEach(row => {
    const name = col.name >= 0 ? String(row[col.name]).trim() : '';
    if (!name) return;
    // Month dropdown now includes the year, e.g. "July 2026" — parse both
    // pieces out so scheduling is unambiguous across year boundaries.
    const monthRaw = col.month >= 0 ? String(row[col.month]).trim() : '';
    const monthMatch = monthRaw.match(/([A-Za-z]+)\s+(\d{4})/);
    const month = monthMatch ? monthMatch[1] : monthRaw; // e.g. "July"
    const monthYear = monthMatch ? parseInt(monthMatch[2], 10) : null; // e.g. 2026

    // Your form's "Event Availability" checkbox answer looks like:
    // "Friday Faith Circles: 6:30PM - 9:30PM, Sunday Suhbah: 10:30AM - 1PM"
    // We just check whether each day name appears anywhere in that text.
    const availableDays = new Set();
    const availabilityText = col.availability >= 0 ? String(row[col.availability]).toLowerCase() : '';
    ['Friday','Sunday','Wednesday','Thursday'].forEach(day => {
      if (availabilityText.includes(day.toLowerCase())) {
        availableDays.add(day);
      }
    });

    // "Current Volunteering Commitments" — if they're volunteering at an
    // event, they should NOT be scheduled as barista for that same day
    // this month, even if they also checked it above (form doesn't block
    // that conflict itself, so we resolve it here).
    const volunteeringText = col.volunteering >= 0 ? String(row[col.volunteering]).toLowerCase() : '';
    ['Friday','Sunday','Wednesday','Thursday'].forEach(day => {
      if (volunteeringText.includes(day.toLowerCase())) {
        availableDays.delete(day);
      }
    });

    const blockedDates = new Set();
    if (col.blockedDates >= 0 && row[col.blockedDates]) {
      String(row[col.blockedDates]).split(',').forEach(chunk => {
        const parsed = parseLooseDate(chunk.trim(), monthYear);
        if (parsed) blockedDates.add(parsed);
      });
    }

    responses.push({ name, month, monthYear, availableDays, blockedDates });
  });
  return responses;
}

// Parses strings like "July 5" into "2026-07-05", using the year from
// the response's own Month dropdown answer (falls back to current year
// if that's somehow missing).
function parseLooseDate(str, yearHint) {
  const m = str.match(/([A-Za-z]+)\s+(\d{1,2})/);
  if (!m) return null;
  const monthIdx = MONTH_NAMES.findIndex(mn => mn.toLowerCase().startsWith(m[1].toLowerCase()));
  if (monthIdx === -1) return null;
  const day = parseInt(m[2], 10);
  const year = yearHint || new Date().getFullYear();
  return year + '-' + String(monthIdx + 1).padStart(2, '0') + '-' + String(day).padStart(2, '0');
}

function buildMonthSchedule(year, monthIndex, responses) {
  const monthName = MONTH_NAMES[monthIndex];
  const relevant = responses.filter(r =>
    (r.month === monthName && (r.monthYear === null || r.monthYear === year)) || r.month === ''
  );

  const shiftCounts = {}; // name -> total shifts assigned this run
  const leadCounts = {};  // name -> total times leading this run

  const dates = getDatesForMonth(year, monthIndex);
  const schedule = dates.map(d => assignShift(d, relevant, shiftCounts, leadCounts));

  return { year, monthName, schedule };
}

// Returns every date in the month matching each configured day-of-week,
// respecting firstWeekOnly (e.g. 1st Thursday only) or lastWeekOnly (last Thursday only)
function getDatesForMonth(year, monthIndex) {
  const results = [];
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

  // First pass: collect every matching date per day-name, in order
  const byDayName = {};
  for (let day = 1; day <= daysInMonth; day++) {
    const d = new Date(year, monthIndex, day);
    const dow = d.getDay();
    for (const [dayName, cfg] of Object.entries(CONFIG.SHIFTS)) {
      if (cfg.dow !== dow) continue;
      if (!byDayName[dayName]) byDayName[dayName] = [];
      byDayName[dayName].push({
        date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        day: dayName,
      });
    }
  }

  // Second pass: apply firstWeekOnly / lastWeekOnly filtering per day-name
  for (const [dayName, cfg] of Object.entries(CONFIG.SHIFTS)) {
    const occurrences = byDayName[dayName] || [];
    if (cfg.firstWeekOnly) {
      if (occurrences.length) results.push(occurrences[0]);
    } else if (cfg.lastWeekOnly) {
      if (occurrences.length) results.push(occurrences[occurrences.length - 1]);
    } else {
      results.push(...occurrences);
    }
  }

  results.sort((a, b) => a.date.localeCompare(b.date));
  return results;
}

// Randomly shuffles an array in place (Fisher–Yates) — used so that among
// equally-fair candidates, who actually gets picked isn't always the same
// person / alphabetically first.
function shuffle(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// Sorts a pool of names by fairness (fewest shifts so far, then fewest
// leads so far), with a RANDOM tiebreak instead of alphabetical — so ties
// don't always resolve the same way every time you regenerate.
function fairSort(names, shiftCounts, leadCounts) {
  const shuffled = shuffle([...names]); // randomize first so ties are random
  return shuffled.sort((a, b) => {
    const sc = (shiftCounts[a] || 0) - (shiftCounts[b] || 0);
    if (sc !== 0) return sc;
    return (leadCounts[a] || 0) - (leadCounts[b] || 0);
  });
}

function assignShift(dateInfo, relevant, shiftCounts, leadCounts) {
  const cfg = CONFIG.SHIFTS[dateInfo.day];
  const isMale = (name) => CONFIG.MALES.includes(name);

  // Eligible pool: available that day-of-week, not blocked on this exact date
  let eligible = relevant.filter(r =>
    r.availableDays.has(dateInfo.day) && !r.blockedDates.has(dateInfo.date)
  ).map(r => r.name);
  eligible = [...new Set(eligible)]; // de-dupe (in case someone submitted twice)

  // Split into single-gender pools — shifts are NEVER mixed.
  const malePool = fairSort(eligible.filter(isMale), shiftCounts, leadCounts);
  const femalePool = fairSort(eligible.filter(n => !isMale(n)), shiftCounts, leadCounts);

  // A pool is "viable" if it can fully staff the required number of active spots.
  const maleViable = malePool.length >= cfg.active;
  const femaleViable = femalePool.length >= cfg.active;

  let chosenPool;
  if (maleViable && femaleViable) {
    // Both groups could staff this shift — pick whichever group is
    // currently BEHIND on shifts (fairness across the whole team, not
    // just within a gender), and randomize when it's a close call.
    const avgCount = (pool) => {
      const top = pool.slice(0, cfg.active);
      return top.reduce((sum, n) => sum + (shiftCounts[n] || 0), 0) / top.length;
    };
    const maleAvg = avgCount(malePool);
    const femaleAvg = femalePool.length ? avgCount(femalePool) : Infinity;
    if (maleAvg === femaleAvg) {
      chosenPool = Math.random() < 0.5 ? malePool : femalePool;
    } else {
      chosenPool = maleAvg < femaleAvg ? malePool : femalePool;
    }
  } else if (maleViable) {
    chosenPool = malePool;
  } else if (femaleViable) {
    chosenPool = femalePool;
  } else {
    // Neither group has enough people — do the best we can with
    // whichever group is larger, rather than leaving the shift empty.
    chosenPool = malePool.length >= femalePool.length ? malePool : femalePool;
  }

  const active = chosenPool.slice(0, cfg.active);
  const backup = chosenPool.slice(cfg.active, cfg.active + cfg.backup);

  // Pick lead = person in "active" with fewest lead turns so far (random tiebreak)
  let lead = null;
  if (active.length) {
    lead = fairSort(active, {}, leadCounts)[0];
    leadCounts[lead] = (leadCounts[lead] || 0) + 1;
  }

  active.forEach(n => { shiftCounts[n] = (shiftCounts[n] || 0) + 1; });

  return {
    date: dateInfo.date,
    day: dateInfo.day,
    time: cfg.time,
    active,
    backup,
    lead,
    runner: null,
  };
}
