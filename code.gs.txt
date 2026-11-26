/**
 * Receipt Tracker v2 — single-file Apps Script app.
 * doGet serves the UI; google.script.run handles reads/writes.
 *
 * Script Properties required:
 *   GEMINI_KEY    - Google AI Studio key
 *   SHEET_ID      - spreadsheet id
 *   GEMINI_MODEL  - optional, defaults below
 */

// 'Groceries & Household' groups supermarket and department-store purchases:
// a receipt from a sundry shop is rarely purely food. Kept as one category on
// purpose - splitting it would need per-line-item reading, which the free model
// does not do reliably.
var CATEGORIES = ['Groceries & Household','Eating Out','Transport','Health','Fun','Kids','Gifts','Sats','One-off','Reimbursed'];
// Old installs recorded 'Groceries'. Mapped on read so history never orphans.
var CATEGORY_ALIASES = { 'groceries': 'Groceries & Household' };

// Settings live in a 'Settings' sheet tab (key | value | note). These are the
// fallbacks used only when that tab is missing or a key is blank.
// These are FALLBACKS only - used when the Settings tab is missing, or a key
// is blank/non-numeric. A fresh install starts at ZERO: a stranger's numbers
// are theirs to type, and the template author's income is worse than none.
// An existing install keeps whatever real values its Settings tab holds.
var DEFAULT_SETTINGS = {
  income_total: 0,
  income_salary: 0,
  income_wife: 0,
  fixed_total: 0,
  variable_budget: 0
};

// Non-numeric settings live here. Kept separate because readSettings() only
// accepts numbers, and a currency code must never fall back to 0.
var DEFAULT_TEXT_SETTINGS = {
  currency: 'USD'
};

// Suffix used for every amount. Widening this list is the whole change needed
// to support a new currency: no maths anywhere depends on it.
var CURRENCY_SYMBOLS = {
  BND: 'BND', USD: 'USD', SGD: 'SGD', MYR: 'MYR', EUR: 'EUR', GBP: 'GBP',
  AUD: 'AUD', CAD: 'CAD', JPY: 'JPY', INR: 'INR', IDR: 'IDR', PHP: 'PHP',
  THB: 'THB', CNY: 'CNY', HKD: 'HKD', KRW: 'KRW', TWD: 'TWD', NZD: 'NZD',
  CHF: 'CHF', AED: 'AED', SAR: 'SAR', ZAR: 'ZAR', NGN: 'NGN', KES: 'KES'
};
// BTC is deliberately NOT a spending currency: sats are the denomination, not
// a ledger currency. Pricing BTC in BTC divides a number by itself.
var HEADERS = ['date','vendor','amount','currency','category','payment','photo','confidence','logged_at','id','updated_at'];
// Receipt extraction is exactly what Google's docs name as Flash-Lite's job:
// "thinking_level.MINIMAL for latency-sensitive or simpler classification and
// extraction tasks". 3.5 Flash is deprecated and auto-routes to the 3.6/3.8
// thinking models, which is why scans went slow. Flash-Lite defaults to minimal
// thinking and is the current recommendation for JSON extraction.
var DEFAULT_GEMINI_MODEL = 'gemini-3.5-flash-lite';

var PROMPT = 'You are reading a photo of a printed retail receipt.\n'
 + 'Return ONLY valid JSON. No markdown fences, no commentary.\n\n'
 + 'Schema:\n'
 + '{"date":"YYYY-MM-DD","vendor":string,"amount":number,'
 + '"currency":ISO 4217 code,"category":string,'
 + '"payment":"Cash"|"Card"|"Transfer","confidence":number}\n\n'
 + 'RULES:\n'
 + '1. DATE: the PURCHASE date, converted to YYYY-MM-DD. If unreadable use null.\n'
 + '   TODAY IS <TODAY>. It is today.\n'
 + '   If the year is not printed, use the current year. NEVER guess a year in the past or future.\n'
 + '   If the date you read is more than 30 days before today, or in the future, use today instead.\n'
 + '2. AMOUNT: the FINAL TOTAL actually paid, AFTER discounts. '
 + 'Never use sub-total, gross total, amount tendered, change, or item counts. '
 + 'If SUB-TOTAL, DISC, then TOTAL appear, use TOTAL.\n'
 + '3. CURRENCY: RM -> MYR. S\$ -> SGD. USD sign or bare \$ -> <HOME_CURRENCY>.\n'
 + '4. CATEGORY: exactly one of: ' + CATEGORIES.join(' | ') + '\n'
 + '   Choose by VENDOR TYPE, not by individual line items:\n'
 + '   - Supermarket, hypermarket, grocery, sundry, department store -> Groceries & Household\n'
 + '   - Restaurant, cafe, takeaway, food court, delivery -> Eating Out\n'
 + '   - Petrol station -> Transport\n'
 + '   - Parking, bus, taxi, ride-hailing, car wash, workshop -> Transport\n'
 + ''
 + '   - Pharmacy, clinic, gym, dental -> Health\n'
 + '   - Cinema, subscriptions, games, hobbies -> Fun\n'
 + '   - Toys, books, clothes, school supplies for your own children -> Kids\n'
 + '   - Anything bought to give away -> Gifts\n'
 + '   - Bitcoin purchase -> Sats\n'
 + '5. CONFIDENCE: 0.0-1.0. Lower it if the total is faded, obscured, or several candidate totals exist.\n'
 + '6. Never invent values. Unreadable -> null.';

/* ---------------- web entry ---------------- */

// No query-parameter API. The UI talks only through google.script.run, so an
// open doPost/doGet?action= endpoint was a data-exfiltration hole for anyone
// deploying with "Who has access: Anyone". Removed deliberately: do not add it
// back, the app does not use it.
function doGet(e) {
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('Stackable')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1,viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------- API (google.script.run) ---------------- */

function getSummary(p) {
  try {
    var body = p || {};
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    var month = body.month || thisMonth();
    var cfg = readSettings();
    var homeCur = usernameCurrency();
    var total = 0, spent = 0, reimbursed = 0, oneoff = 0, oneoffCount = 0;
    // Rows in another currency are folded into the home total at the live
    // rate, never 1:1. Summing raw amounts put a USD 40 dinner into a BND
    // budget as 40, and every downstream figure (pace, stackable, sats)
    // inherited the error. Home-currency rows are untouched.
    var homePerUsd = (homeCur === 'USD') ? 1 : usdPerUnit(homeCur);
    var fxCounts = {};
    // regular = everything except One-off. Reimbursed stays IN because a
    // same-month offset (a friend paying back their half of a dinner) means
    // the gross receipt was never the user's spend. One-off is excluded
    // because a repair is an event, not a daily rate.
    var regular = 0;
    var stacked = 0, stackedCount = 0, stackedRows = [];
    var byCategory = {}, count = 0, oneoffRows = [];

    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r || !r.length) continue;
      if (ymOf(r[0]) !== month) continue;
      var rawAmt = Number(r[2]);
      if (isNaN(rawAmt)) continue;
      // Missing code means the user's own ledger currency, the same rule the
      // write path uses. A foreign row is converted here, exactly once.
      var rowCur = String(r[3] || homeCur).toUpperCase();
      var amt = rawAmt;
      if (rowCur !== homeCur) {
        amt = toHomeAmount(rawAmt, rowCur, homeCur, homePerUsd);
        fxCounts[rowCur] = (fxCounts[rowCur] || 0) + 1;
      }
      // Normalise on READ too. Rows written before the Groceries rename still
      // hold the old string in the sheet, and folding them here is what stops
      // the By Category grid splitting one bucket into two.
      var cat = ensureCategory(String(r[4] || 'Uncategorised').trim());
      total += amt;
      if (cat === 'Sats') {
        // Bitcoin is the DESTINATION, not a cost. Treating it like Eating Out
        // meant a stack on day 5 projected as month-end spending, drove the
        // pace to nonsense and taught the auto-budget that stacking is a
        // living expense. It is excluded from spent, regular, byCategory and
        // the budget average; it gets its own figure that the hero reads.
        stacked += amt;
        stackedCount++;
        if (stackedRows.length < 50) {
          stackedRows.push({ date: isoOf(r[0]), vendor: String(r[1] || '').trim() || 'Unnamed',
                             amount: round2(rawAmt), currency: rowCur });
        }
      } else if (cat === 'Reimbursed') {
        // Money came back. It must lower what this month cost, or the hero,
        // the pace projection and "stackable" all overstate the outflow.
        // Stored negative; netting it into spent is the whole point.
        reimbursed += amt; // negative
        spent += amt;
        regular += amt;
      } else if (cat === 'One-off') {
        // One-off money is still money spent, so it counts in the month total.
        // It is tracked separately only so the pace/bar can compare like with
        // like; the summary and the stackable maths both include it.
        oneoff += amt;
        oneoffCount++;
        // Carry the actual entries so the client can expand the one-off row
        // without a second round-trip. Capped: this is a drill-down, not a ledger.
        if (oneoffRows.length < 50) {
          // Stored amount and its own code, NOT the converted figure: this is
          // a receipt drill-down, and a receipt shows what the shop charged.
          oneoffRows.push({ date: isoOf(r[0]), vendor: String(r[1] || '').trim() || 'Unnamed',
                            amount: round2(rawAmt), currency: rowCur });
        }
        spent += amt;
        byCategory[cat] = round2((byCategory[cat] || 0) + amt);
        count++;
      } else {
        spent += amt;
        regular += amt;
        byCategory[cat] = round2((byCategory[cat] || 0) + amt);
        count++;
      }
    }

    // Budget comes from the Settings tab unless the caller overrides it.
    var budget = (typeof body.budget === 'number' && body.budget > 0)
      ? Number(body.budget) : cfg.variable_budget;
    // Spend is real money out the door: One-off included, reimbursements netted.
    var variableSpent = round2(spent);
    return {
      ok: true, month: month, total: round2(total), spent: round2(spent),
      // The rate-bearing spend. The client projects THIS to month end, then
      // adds one-off once. Never re-derive it client-side: reimbursed is
      // stored negative and the sign has bitten twice already.
      regular: round2(regular),
      reimbursed: round2(reimbursed), budget: budget,
      remaining: round2(budget - variableSpent),
      oneoff: round2(oneoff), oneoffCount: oneoffCount, oneoffRows: oneoffRows,
      // What has actually been bought this month. The hero shows this beside
      // the remaining capacity; it is NOT part of spent, so it never inflates
      // the pace projection or the budget suggestion.
      stacked: round2(stacked), stackedCount: stackedCount, stackedRows: stackedRows,
      settings: cfg,
      // The client's first-run check reads this from the summary response.
      // Without it the setup screen reopened on every single page load.
      configured: prop('SETTINGS_CONFIGURED') === '1',
      // Settings reads these to show "Key saved" instead of the placeholder.
      // A boolean only - the key never leaves Script Properties.
      hasKey: !!prop('GEMINI_KEY'),
      geminiModel: prop('GEMINI_MODEL') || '',
      debtPayment: totalDebtPayment(),
      // Which currencies were converted this month, so the client can say so
      // instead of silently changing a total the user checks on receipts.
      homeCurrency: homeCur, fx: fxCounts,
      byCategory: byCategory, count: count
    };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// Rolling average of REAL variable spend over the last N complete months.
// This is the 'find your number' feature: the budget stops being a guess and
// becomes a measurement. Only completed months count - a partial month always
// looks low and would drag the suggestion down.
function getBudgetSuggestion(p) {
  try {
    var body = p || {};
    var months = Number(body.months) || 3;
    if (months < 1) months = 1;
    if (months > 12) months = 12;

    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    var nowYm = thisMonth();

    // Walk back from last month; skip any that predate the first receipt.
    var homeCur = usernameCurrency();
    var homePerUsd = (homeCur === 'USD') ? 1 : usdPerUnit(homeCur);
    var buckets = {};
    var earliest = '';
    var counts = {};
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r || !r.length) continue;
      var ym = ymOf(r[0]);
      if (!ym) continue;
      if (!earliest || ym < earliest) earliest = ym;
      var rawAmt = Number(r[2]);
      if (isNaN(rawAmt)) continue;
      // Same conversion as getSummary: a foreign row is folded into the home
      // currency at the live rate before it enters a monthly bucket. An average
      // that mixes raw USD and BND amounts suggests a target nobody can hit.
      var rowCur = String(r[3] || homeCur).toUpperCase();
      var amt = toHomeAmount(rawAmt, rowCur, homeCur, homePerUsd);
      // ensureCategory, not the raw string: a row typed as 'one-off' or
      // ' One-off ' used to slip past the One-off filter below and pollute the
      // average with an event-shaped expense. getSummary already normalises
      // this way, so the suggestion now uses the same category identity.
      var cat = ensureCategory(String(r[4] || '').trim());
      // One-off only. Reimbursed is included deliberately so this average uses
      // the same definition of 'regular spend' as getSummary().regular and the
      // pace projection -- a same-month offset was never real spend, and three
      // numbers in one app disagreeing about that is how targets rot.
      // Sats is stacking, not consumption. Counting it here taught the app
      // that buying bitcoin is a monthly living cost, which inflated the
      // suggested target by whatever the user was stacking.
      if (cat === 'One-off' || cat === 'Sats') continue;
      buckets[ym] = round2((buckets[ym] || 0) + amt);
      counts[ym] = (counts[ym] || 0) + 1;
    }

    // Complete months only: strictly before the current month.
    var used = [];
    var ym = nowYm;
    for (var m = 0; m < months; m++) {
      ym = prevMonth(ym);
      if (!ym) break;
      used.push(ym);
    }
    used = used.filter(function(y){ return y >= earliest; }).reverse();

    // Months that simply predate the first receipt hold no information, so
    // drop them and average only months that were actually tracked. Averaging
    // a blank August into a real October produces a target nobody should use.
    used = used.filter(function(y){
      return (counts[y] || 0) > 0 || (buckets[y] || 0) > 0;
    });

    if (!used.length) {
      return { ok: true, months: 0, average: 0, used: [],
               ready: false, reason: 'No complete month yet' };
    }

    var sum = 0;
    var detail = [];
    for (var j = 0; j < used.length; j++) {
      // Never negative. A month whose only rows were reimbursements nets below
      // zero, and averaging that in pulls the suggested target down to a number
      // nobody could live on.
      var v = Math.max(0, buckets[used[j]] || 0);
      sum += v;
      detail.push({ month: used[j], spent: round2(v), count: counts[used[j]] || 0 });
    }
    var avg = round2(sum / used.length);
    return {
      ok: true,
      months: used.length,
      average: avg,
      used: detail,
      ready: used.length >= 2,
      reason: used.length >= 2 ? ''
        : 'Only ' + used.length + ' tracked month' + (used.length === 1 ? '' : 's') +
          ' so far. One more and the average means something.'
    };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// ---- FX table ---------------------------------------------------------
// One fetch of /v6/latest/USD returns every supported rate at once. Cached
// 12h: rates do not move enough intraday to justify a request per reload.
var FX_MAP_KEY = 'FX_USD_MAP_CACHE';
var FX_TTL_MS = 12 * 60 * 60 * 1000;

function loadFxMap(cached) {
  // 1. New shape: a full table.
  try {
    var raw = prop(FX_MAP_KEY);
    if (raw) {
      var m = JSON.parse(raw);
      if (m && m.rates && m.ts && (Date.now() - m.ts) < FX_TTL_MS) {
        // A table written by an older build may lack currencies the dropdown
        // now offers. Treat it as expired so we refetch rather than return 0.
        var gap = false;
        for (var c in CURRENCY_SYMBOLS) {
          if (!(Number(m.rates[c]) > 0)) { gap = true; break; }
        }
        if (!gap) return m;
      }
    }
  } catch (e) { /* corrupt: fall through */ }

  // 2. Fetch once, keep whatever we recognise.
  try {
    var res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD',
      { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() === 200) {
      var j = JSON.parse(res.getContentText());
      var src = (j && j.rates) || null;
      if (src) {
        var keep = {};
        for (var code in CURRENCY_SYMBOLS) {
          var v = Number(src[code]);
          if (v && isFinite(v) && v > 0) keep[code] = v;
        }
        // A table missing any supported code is a broken table: it would
        // silently return 0 for that currency and blank the badge. Only cache
        // a table that covers every code the dropdown offers. USD is the base
        // so it is 1 by definition and may be absent from the feed.
        keep.USD = 1;
        var missing = [];
        for (var want in CURRENCY_SYMBOLS) {
          if (!keep[want]) missing.push(want);
        }
        if (!missing.length) {
          var out = { ts: Date.now(), rates: keep };
          try { PropertiesService.getScriptProperties().setProperty(FX_MAP_KEY, JSON.stringify(out)); } catch (e) {}
          return out;
        }
        // Partial table: refuse to cache it, but keep what we got so the
        // caller can still price the currencies it does cover.
        return { ts: 0, rates: keep, partial: missing };
      }
    }
  } catch (e) { /* fall through to stale copies */ }

  // 3. Expired but real table beats nothing.
  try {
    var raw2 = prop(FX_MAP_KEY);
    if (raw2) {
      var m2 = JSON.parse(raw2);
      if (m2 && m2.rates) return m2;
    }
  } catch (e) { /* ignore */ }

  // 4. Legacy single-pair cache (pre-multi-currency builds).
  try {
    var legacy = prop('FX_USD_BND_CACHE');
    if (legacy) {
      var lc = JSON.parse(legacy);
      if (lc && lc.rate > 0) return { ts: lc.ts || 0, rates: { BND: Number(lc.rate), USD: 1 } };
    }
  } catch (e) { /* ignore */ }
  if (cached && cached.fx_bnd > 0) return { ts: 0, rates: { BND: Number(cached.fx_bnd), USD: 1 } };

  return { ts: 0, rates: { USD: 1 } };
}

// UNITS PER USD for a currency code (BND: 1.34, MYR: 4.40, JPY: 150), the same
// shape as loadFxMap().rates. This is NOT "how many USD is one unit".
// Converting A -> B therefore needs BOTH rates:
//     amountB = amountA / usdPerUnit(A) * usdPerUnit(B)
// Crossing to USD and stopping there was the bug: a foreign row in a non-USD
// ledger was left denominated in USD and under-counted everywhere downstream.
// Returns 1 for an unknown code so a missing rate degrades to "no conversion"
// rather than zeroing a real expense.
function usdPerUnit(code) {
  var c = String(code || '').toUpperCase();
  if (!c || c === 'USD') return 1;
  var map = loadFxMap(null);
  var r = Number(map && map.rates ? map.rates[c] : 0);
  return (r && isFinite(r) && r > 0) ? r : 1;
}

// One row amount expressed in the home currency. Both callers must pass
// homePerUsd from OUTSIDE their row loop: fetching the FX table per row would
// do a cache read (or worse, a fetch) for every expense in the sheet.
function toHomeAmount(amt, rowCur, homeCur, homePerUsd) {
  if (!rowCur || rowCur === homeCur) return amt;
  var hp = (homePerUsd > 0) ? homePerUsd : usdPerUnit(homeCur);
  return amt / usdPerUnit(rowCur) * hp;
}

// The currency the ledger is kept in. Sats maths must use THIS rate, not BND.
function usernameCurrency() {
  try {
    var cur = String(readSettings().currency || 'USD').toUpperCase();
    return CURRENCY_SYMBOLS[cur] ? cur : 'USD';
  } catch (e) { return 'USD'; }
}

// Live BTC price, cached 60 min in Script Properties so a phone reloading the
// summary a dozen times doesn't hammer the API (and get rate-limited).
// CoinGecko has no BND pair, so we chain USD -> BND via a free FX feed.
function getBtcPrice(p) {
  try {
    var body = p || {};
    var force = !!body.force;
    var cacheKey = 'BTC_PRICE_CACHE';
    var freshMs = 60 * 60 * 1000;      // serve from cache, no network
    var staleMs = 24 * 60 * 60 * 1000; // last resort, clearly flagged

    // Read whatever we have once: it is both the fresh path and the fallback.
    var cached = null;
    var rawCache = prop(cacheKey);
    if (rawCache) {
      try {
        var parsed = JSON.parse(rawCache);
        // Keyed on price_usd: that is the value we quote and the one that
        // cannot silently age. Legacy entries (price_bnd only) still load so
        // an upgrade does not blank the badge for an hour.
        if (parsed && parsed.ts && parsed.price_usd > 0) cached = parsed;
      } catch (e) { /* corrupt cache: ignore */ }
    }

    if (!force && cached && (Date.now() - cached.ts) < freshMs) {
      var cMap = loadFxMap(cached);
      var cCur = usernameCurrency();
      var cFx = (cCur === 'USD') ? 1 : Number(cMap.rates[cCur] || 0);
      return { ok: true, price_bnd: round2(cached.price_usd * cFx),
               price_usd: cached.price_usd,
               fx_bnd: Number(cMap.rates.BND) || 0, fx_rate: cFx, fx_currency: cCur,
               fetched_at: cached.ts, cached: true, stale: false };
    }

    // Spot price: three independent no-key sources, first success wins.
    // CoinGecko rate-limits anonymous callers with HTTP 429, so it must not
    // be the only door.
    var usd = 0;
    var lastErr = '';
    var sources = [
      { name: 'coingecko', url: 'https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=usd',
        pick: function(j){ return j && j.bitcoin && Number(j.bitcoin.usd); } },
      { name: 'coinbase', url: 'https://api.coinbase.com/v2/prices/BTC-USD/spot',
        pick: function(j){ return j && j.data && Number(j.data.amount); } },
      { name: 'kraken', url: 'https://api.kraken.com/0/public/Ticker?pair=XBTUSD',
        pick: function(j){ return j && j.result && j.result.XXBTZUSD && Number(j.result.XXBTZUSD.c[0]); } }
    ];

    for (var i = 0; i < sources.length; i++) {
      var src = sources[i];
      var tries = 1;
      // A 429 is usually momentary; one retry clears most of them.
      for (var attempt = 0; attempt < tries + 1; attempt++) {
        try {
          var res = UrlFetchApp.fetch(src.url, { muteHttpExceptions: true, followRedirects: true });
          var code = res.getResponseCode();
          if (code === 200) {
            var v = src.pick(JSON.parse(res.getContentText()));
            if (v && isFinite(v) && v > 0) { usd = v; break; }
            lastErr = src.name + ': no price in payload';
          } else {
            lastErr = src.name + ' HTTP ' + code;
            if (code === 429 && attempt < tries) { Utilities.sleep(1200); continue; }
          }
        } catch (err) {
          lastErr = src.name + ': ' + String(err);
        }
        break; // non-429 failures go to the next source immediately
      }
      if (usd > 0) break;
    }

    // FX: ONE fetch returns every rate we might need, cached 12h. This used to
    // be hardcoded to BND, which meant a user on SGD divided SGD by a BND price
    // and got a wrong sats figure that still looked plausible. Now every code in
    // CURRENCY_SYMBOLS works, and a currency with no rate reports null instead
    // of silently inventing a price.
    var fxMap = loadFxMap(cached);
    var wantCur = usernameCurrency();
    var fx = (wantCur === 'USD') ? 1 : Number(fxMap.rates[wantCur] || 0);

    // Sanity gate: a source can return a garbled number that would be cached
    // for an hour and silently distort every sats figure. Compare against the
    // last known good price; anything beyond +/-35% in one fetch is a bad
    // payload, not a market move.
    if (usd > 0 && cached && cached.price_usd > 0) {
      var drift = Math.abs(usd - cached.price_usd) / cached.price_usd;
      if (drift > 0.35) {
        lastErr = 'rejected outlier: ' + round2(usd) + ' vs cached ' +
                  round2(cached.price_usd) + ' (' + Math.round(drift * 100) + '% drift)';
        usd = 0;
      }
    }
    // A BTC price outside 1k-10M USD is corrupt, not news.
    if (usd > 0 && (usd < 1000 || usd > 10000000)) {
      lastErr = 'rejected implausible price: ' + usd;
      usd = 0;
    }

    if (usd > 0 && fx > 0) {
      var priceBnd = round2(usd * fx);
      var payload = { ts: Date.now(), price_bnd: priceBnd, price_usd: round2(usd),
                      fx_bnd: Number(fxMap.rates.BND) || 0, fx_rate: fx, fx_currency: wantCur };
      try {
        PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(payload));
      } catch (e) { /* best effort */ }
      return { ok: true, price_bnd: priceBnd, price_usd: round2(usd),
               fx_bnd: payload.fx_bnd, fx_rate: fx, fx_currency: wantCur,
               fetched_at: payload.ts, cached: false, stale: false };
    }

    // Live fetch failed. An old real price beats a blank, as long as it says
    // how old it is - the client renders that age next to the figure.
    if (cached && (Date.now() - cached.ts) < staleMs) {
      var sMap = loadFxMap(cached);
      var sCur = usernameCurrency();
      var sFx = (sCur === 'USD') ? 1 : Number(sMap.rates[sCur] || 0);
      return { ok: true, price_bnd: round2(cached.price_usd * sFx),
               price_usd: cached.price_usd,
               fx_bnd: Number(sMap.rates.BND) || 0, fx_rate: sFx, fx_currency: sCur,
               fetched_at: cached.ts, cached: true, stale: true,
               note: lastErr || 'live fetch failed' };
    }

    return { ok: false, error: lastErr || 'Price feed unavailable' };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function prevMonth(ym) {
  var p = String(ym || '').split('-');
  var y = parseInt(p[0], 10), m = parseInt(p[1], 10);
  if (!y || !m) return '';
  m -= 1;
  if (m === 0) { m = 12; y -= 1; }
  return y + '-' + (m < 10 ? '0' : '') + m;
}

function getList(p) {
  try {
    var body = p || {};
    // Was a hard 25 with no way to ask for more, so a busy month silently
    // hid older rows and the user concluded they had not saved. Cap kept as a
    // safety valve; the client asks for more when it wants them.
    var limit = Math.min(Number(body.limit) || 25, 500);
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    var rows = [];
    // Filter to the selected month FIRST. Without this the list showed the
    // newest rows of any month, and the header total summed every month ever
    // recorded, which is why Expenses disagreed with Summary.
    var wantMonth = String(body.month || thisMonth());
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r || !r.length) continue;
      if (ymOf(r[0]) !== wantMonth) continue;
      rows.push({
        id: String(r[9] || ''),
        rowNumber: i + 1,
        date: isoOf(r[0]),
        vendor: String(r[1] || ''),
        amount: Number(r[2]) || 0,
        // Same rule as everywhere else: a blank code means the ledger's own
        // currency, never a hardcoded BND.
        currency: String(r[3] || usernameCurrency()),
        category: String(r[4] || ''),
        payment: String(r[5] || ''),
        ts: tsOf(r[0])
      });
    }
    // Latest date on top. Tie-break: newest row first.
    rows.sort(function(a, b) {
      if (a.ts !== b.ts) return b.ts - a.ts;
      return b.rowNumber - a.rowNumber;
    });
    // Must match getSummary's rules exactly, or Expenses and Summary report
    // different totals for the same month. Summary excludes Reimbursed rows
    // from the count and subtracts them from spend; the list simply does not
    // count them, which is the same headline figure a reader expects.
    // Net Reimbursed rows IN, not out: getSummary subtracts them from spend, and
    // a list header that skipped them disagreed with the Summary by exactly the
    // reimbursed amount. Still not counted as an entry of its own.
    // The header total must equal getSummary().spent for the same month, or the
    // two tabs report different numbers for identical data. That means the
    // same conversion: a foreign receipt is worth its home-currency value, not
    // its raw number. Home rows take the no-op path.
    var listHomeCur = usernameCurrency();
    var listHomePerUsd = (listHomeCur === 'USD') ? 1 : usdPerUnit(listHomeCur);
    var monthTotal = 0, monthCount = 0;
    for (var m = 0; m < rows.length; m++) {
      var catM = ensureCategory(String(rows[m].category || '').trim());
      var rowMCur = String(rows[m].currency || listHomeCur).toUpperCase();
      // Sats rows are shown (so they can be edited or deleted) but are not
      // spending, so they must not enter the header total. Same rule as
      // getSummary: the two tabs have to agree on what "total" means.
      if (catM === 'Sats') continue;
      monthTotal += (rowMCur === listHomeCur)
        ? (Number(rows[m].amount) || 0)
        : toHomeAmount(Number(rows[m].amount) || 0, rowMCur, listHomeCur, listHomePerUsd);
      if (catM !== 'Reimbursed') monthCount++;
    }
    monthTotal = round2(monthTotal);

    // Truncation must be measured BEFORE the slice. Computing it after (which
    // an earlier build did) compares an already-capped array against the cap,
    // so it is always false and the button never renders.
    var hasMore = rows.length > limit;
    if (hasMore) rows = rows.slice(0, limit);
    for (var j = 0; j < rows.length; j++) delete rows[j].ts;
     return { ok: true, rows: rows, total: monthTotal, count: monthCount,
             hasMore: hasMore };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function addManual(p) {
  try {
    var b = p || {};
    var vendor = String(b.vendor || '').trim();
    var amount = round2(b.amount);
    if (!vendor) return { ok: false, error: 'Vendor is required' };
    if (typeof amount !== 'number' || isNaN(amount) || amount === 0) {
      return { ok: false, error: 'Amount must be a non-zero number' };
    }
    // A 'Reimbursed' entry stores a negative amount (money coming back in).
    var cat = ensureCategory(b.category);
    if (cat === 'Reimbursed') amount = -Math.abs(amount);

    var row = {
      date: b.date || todayISO(),
      vendor: vendor,
      amount: amount,
      // Missing currency means the user's own ledger currency. 'BND' here put a
      // US or UK user's typed expense in a currency they never chose.
      currency: String(b.currency || usernameCurrency()).toUpperCase(),
      category: cat,
      payment: b.payment || 'Cash',
      photo: '',
      confidence: 1,
      logged_at: nowISO()
    };
    appendRow(row);
    return { ok: true, row: row };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// Reads the receipt and returns what it found. Does NOT write to the sheet -
// the user reviews amount/category/vendor first, then calls addReceipt.
function extractReceipt(p) {
  try {
    var b = p || {};
    if (!b.image) return { ok: false, error: 'No image received' };

    var rec = extractFromImage(b.image, b.mimeType || 'image/jpeg');
    if (rec.error) {
      return { ok: false, error: rec.error, throttled: !!rec.throttled };
    }

    // Plausibility guard: a receipt photographed today is not from 2023.
    // Rather than trust an invented year, flag it in the draft so the user
    // sees it before it is saved. Silent wrong data is the worst outcome.
    var dIso = rec.date || todayISO();
    var today = todayISO();
    var dateWarn = '';
    if (dIso && dIso < today) {
      var daysBack = Math.round((new Date(today + 'T00:00:00') - new Date(dIso + 'T00:00:00')) / 86400000);
      if (daysBack > 30) dateWarn = 'Read as ' + dIso + ', which is ' + daysBack + ' days ago. Is that right?';
    }
    if (dIso && dIso > today) {
      var daysFwd = Math.round((new Date(dIso + 'T00:00:00') - new Date(today + 'T00:00:00')) / 86400000);
      if (daysFwd > 0) dateWarn = 'Read as ' + dIso + ', which is in the future. Fix the date.';
    }

    return {
      ok: true,
      draft: {
        date: dIso,
        dateWarn: dateWarn,
        vendor: String(rec.vendor || '').trim(),
        amount: round2(rec.amount),
        // A receipt with no printed currency belongs to the user's ledger, not
        // to Brunei. Hardcoded 'BND' put a stranger's receipt in a currency
        // their settings never mention.
        currency: String(rec.currency || usernameCurrency()).toUpperCase(),
        category: ensureCategory(rec.category),
        payment: rec.payment || 'Cash',
        confidence: (typeof rec.confidence === 'number') ? rec.confidence : null
      }
    };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function addReceipt(p) {
  try {
    var b = p || {};
    var vendor = String(b.vendor || '').trim();
    var amount = round2(b.amount);
    if (!vendor) return { ok: false, error: 'Vendor is required' };
    if (typeof amount !== 'number' || isNaN(amount) || amount === 0) {
      return { ok: false, error: 'Amount must be a non-zero number' };
    }
    // Same rule as addManual: a Reimbursed photo must store a negative amount,
    // or getSummary counts money coming back as money spent.
    var rcat = ensureCategory(b.category);
    if (rcat === 'Reimbursed') amount = -Math.abs(amount);

    var row = {
      date: b.date || todayISO(),
      vendor: vendor,
      amount: amount,
      currency: String(b.currency || usernameCurrency()).toUpperCase(),
      category: rcat,
      payment: b.payment || 'Cash',
      photo: b.photoName || '',
      confidence: (typeof b.confidence === 'number') ? b.confidence : '',
      logged_at: nowISO()
    };
    appendRow(row);
    return { ok: true, row: row };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}


// Locate a row by its stable id (column J). Row numbers are a snapshot: sort
// the Sheet, or delete a row on another device, and the number points at a
// different expense. The id is the only safe key.
function rowNumberById_(sheet, id) {
  var want = String(id || '').trim();
  if (!want) return 0;
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var ids = sheet.getRange(2, 10, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === want) return i + 2;
  }
  return 0;
}

function updateExpense(p) {
  try {
    var b = p || {};
    var sheet = getSheet();
    var rowNumber = b.id ? rowNumberById_(sheet, b.id) : Number(b.rowNumber);
    if (!rowNumber || rowNumber < 2) {
      return { ok: false, error: b.id ? 'That expense is no longer in the sheet' : 'Missing row' };
    }
    var last = sheet.getLastRow();
    if (rowNumber > last) return { ok: false, error: 'That expense no longer exists' };

    var vendor = String(b.vendor || '').trim();
    var amount = round2(b.amount);
    if (!vendor) return { ok: false, error: 'Vendor is required' };
    if (typeof amount !== 'number' || isNaN(amount) || amount === 0) {
      return { ok: false, error: 'Amount must be a non-zero number' };
    }

    var range = sheet.getRange(rowNumber, 1, 1, 11);
    var cur = range.getValues()[0];
    var cat = ensureCategory(b.category || cur[4]);
    // Reimbursed rows are stored negative. The edit form pre-fills the stored
    // -50.00, so only force the sign when the row is NOT already reimbursed;
    // a re-save used to double-negate it back to +50 and flip the meaning.
    var wasReimbursed = ensureCategory(cur[4]) === 'Reimbursed';
    if (cat === 'Reimbursed' && !wasReimbursed) amount = -Math.abs(amount);
    // ...and the reverse: moving a -50.00 reimbursement to Eating Out must not
    // leave Eating Out holding a negative expense. Only flip when the sign
    // still comes from the old category, i.e. the user did not retype it.
    else if (cat !== 'Reimbursed' && wasReimbursed) amount = Math.abs(amount);

    range.setValues([[
      b.date || isoOf(cur[0]) || todayISO(),
      vendor,
      amount,
      String(b.currency || cur[3] || 'USD').toUpperCase(),
      cat,
      b.payment || cur[5] || 'Cash',
      cur[6] || '',
      cur[7] === '' ? 1 : cur[7],
      cur[8] || nowISO(),
      cur[9] || newId(),
      nowISO()
    ]]);

    return { ok: true, row: {
      id: String(cur[9] || ''), rowNumber: rowNumber, date: b.date || isoOf(cur[0]),
      vendor: vendor, amount: amount, currency: String(b.currency || cur[3] || 'USD').toUpperCase(),
      category: cat, payment: b.payment || cur[5]
    }};
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function deleteExpense(p) {
  try {
    var b = p || {};
    var sheet = getSheet();
    var rowNumber = b.id ? rowNumberById_(sheet, b.id) : Number(b.rowNumber);
    if (!rowNumber || rowNumber < 2) {
      return { ok: false, error: b.id ? 'That expense is no longer in the sheet' : 'Missing row' };
    }
    if (rowNumber > sheet.getLastRow()) return { ok: false, error: 'That expense no longer exists' };
    sheet.deleteRow(rowNumber);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/* ---------------- gemini ---------------- */

// The prompt needs today's date so the model cannot invent a year. A receipt
// with no printed year was read as 2023-11-22 on 2026-11-22: with no anchor
// the model guessed. Substituted at call time, not baked into PROMPT.
function promptWithToday() {
  var today = Utilities.formatDate(new Date(), TZ(), 'yyyy-MM-dd');
  // A bare "$" is the user's own currency, not necessarily Brunei's. Hardcoding
  // BND meant a Singapore, US or Australian user's receipt was filed in a
  // currency their ledger never uses.
  return PROMPT.replace('<TODAY>', today).replace(/<HOME_CURRENCY>/g, usernameCurrency());
}

function extractFromImage(b64, mime) {
  var key = prop('GEMINI_KEY');
  // User-facing: an internal property name means nothing to someone who
  // tapped "Skip for now", and this is the first thing they hit on a scan.
  if (!key) return { error: 'Add your Gemini key in Settings to scan receipts. Typing them in still works.' };

  var model = prop('GEMINI_MODEL') || DEFAULT_GEMINI_MODEL;
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + model + ':generateContent' + '?' + 'ke' + 'y' + '=' + key;

  var payload = {
    contents: [{
      parts: [
        { text: promptWithToday() },
        { inline_data: { mime_type: mime, data: b64 } }
      ]
    }],
    generationConfig: {
      // Flash-Lite ignores custom temperature (a documented 3.5 breaking change),
      // so it is dropped rather than sent and silently discarded.
      response_mime_type: 'application/json',
      maxOutputTokens: 512,
      // thinkingBudget:0 was ignored by 3.x models - they thought anyway. The
      // new control is thinkingLevel; MINIMAL is the documented setting for
      // extraction and is the model's own default, made explicit here.
      thinkingConfig: { thinkingLevel: 'MINIMAL' }
    }
  };

  // ONE attempt only. Retrying burned Apps Script's 30s execution budget and
  // killed the call silently, which is what made scans hang then snap back.
  // A hard 15s cap means we fail *before* Apps Script does, so the UI gets a
  // real message instead of dying on the vine.
  var res, code, txt;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
      // No deadline param: UrlFetchApp silently ignores it. The ~60s ceiling is
      // fixed by Apps Script, which is why this is ONE attempt only above.
    });
    code = res.getResponseCode();
    txt = res.getContentText();
  } catch (netErr) {
    return { error: 'Gemini did not answer in time. Try again, or type it in.', throttled: true };
  }

  if (code !== 200) {
    var short = txt;
    try {
      var ej = JSON.parse(txt);
      if (ej.error && ej.error.message) short = ej.error.message;
    } catch (ignore) {}
    if (code === 403) {
      return { error: 'Gemini 403 (model ' + model + ') - the model name or key is wrong. Check GEMINI_MODEL / GEMINI_KEY.', throttled: false };
    }
    if (code === 429) {
      return { error: 'Gemini is rate-limited right now. Type it in, or retry in a minute.', throttled: true };
    }
    return { error: 'Gemini ' + code + ': ' + String(short).substring(0, 220) };
  }

  try {
    var data = JSON.parse(txt);
    if (!data.candidates || !data.candidates.length) {
      var fb = (data.promptFeedback && data.promptFeedback.blockReason)
        ? 'blocked: ' + data.promptFeedback.blockReason : 'no candidates';
      return { error: 'Gemini returned nothing (' + fb + ')' };
    }
    var cand = data.candidates[0];
    if (cand.finishReason && cand.finishReason !== 'STOP') {
      if (cand.finishReason === 'MAX_TOKENS') {
        return { error: 'Gemini ran out of output tokens - try a closer photo' };
      }
      if (cand.finishReason === 'SAFETY' || cand.finishReason === 'RECITATION') {
        return { error: 'Gemini refused this image (' + cand.finishReason + ')' };
      }
    }
    var parts = (cand.content && cand.content.parts) ? cand.content.parts : [];
    var raw = '';
    for (var pi = 0; pi < parts.length; pi++) {
      var pt = parts[pi];
      // Thinking models return reasoning parts alongside the answer. Concatenating
      // them prepends prose to the JSON and breaks the parse. Skip thought parts.
      if (pt && pt.thought === true) continue;
      if (pt && typeof pt.text === 'string') raw += pt.text;
    }
    if (!raw) return { error: 'Gemini returned an empty response' };
    return parseJSONLoose(raw);
  } catch (err) {
    return { error: 'Could not parse Gemini response: ' + String(err).substring(0, 120) };
  }
}

function parseJSONLoose(raw) {
  var t = String(raw).trim();
  var fence = String.fromCharCode(96, 96, 96);
  if (t.indexOf(fence) === 0) {
    t = t.replace(new RegExp('^' + fence + '(?:json)?', 'i'), '');
  }
  t = t.split(fence).join('').trim();
  try { return JSON.parse(t); } catch (e1) {}
  var m = t.match(/[{][\s\S]*[}]/);
  if (m) {
    try { return JSON.parse(m[0]); } catch (e2) {}
  }
  return { error: 'Model did not return valid JSON' };
}

/* ---------------- bound-sheet support ---------------- */

// When this script is bound to a spreadsheet (Extensions > Apps Script from
// inside the sheet), SpreadsheetApp.getActive() is the user's own copy. That is
// what makes a copy-link install work without a setup screen.
// EVERY server-side write must work from a /exec web app, where there is no
// "active spreadsheet" and getActiveSpreadsheet() returns null. Order:
//   1. SHEET_ID in Script Properties (set once by setupSheetId)
//   2. The bound spreadsheet, when this runs from inside the Sheet
// Without step 1, a copy-link install silently wrote nothing: the functions
// returned {ok:false} and Apps Script still logged them as 'Completed'.
function resolveSheet_() {
  var id = '';
  try { id = prop('SHEET_ID') || ''; } catch (e) { id = ''; }
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* fall through */ }
  }
  try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { return null; }
}

// One-tap setup from inside the Sheet (menu item or editor Run). A bound
// script can see its own spreadsheet; a web app cannot. This is the only place
// SHEET_ID is written, so it must be run once per install.
function setupSheetId() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { ok: false, error: 'Run this from inside the Sheet (Extensions > Apps Script > Run), not from the web app URL.' };
    PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
    return { ok: true, sheetId: ss.getId() };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

// A bound script IS the spreadsheet. Reporting configured:false there opened a
// connect screen asking for a link the user had already opened.
function getInstallState() {
  try {
    var id = prop('SHEET_ID') || '';
    var bound = false;
    try { bound = !!SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { bound = false; }
    return {
      ok: true,
      configured: !!(id || bound),
      bound: bound,
      hasKey: !!prop('GEMINI_KEY'),
      model: prop('GEMINI_MODEL') || DEFAULT_GEMINI_MODEL,
      sheetTail: id ? id.slice(-6) : ''
    };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// Menu for a bound copy. Without this the user has the code but no way to
// reach it except a deployment URL, which is the step they cannot do.
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Stackable')
      .addItem('Open expense tracker', 'showTracker')
      .addSeparator()
      .addItem('Set up this sheet (web app access)', 'setupSheetId')
      .addToUi();
  } catch (e) { /* not bound, or no UI */ }
}

function showTracker() {
  // Single-file project: the UI is the APP_HTML literal in this file. Never
  // call createTemplateFromFile - a copy-link install has no Index file.
  var html = HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('Stackable')
    .setWidth(1100)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'Stackable');
}


/* ---------------- standalone install ---------------- */

// Writes what a user must supply into Script Properties, then proves the sheet
// is reachable before reporting success. Only used by an UNBOUND copy (the
// /exec web-app route). A bound copy already knows its spreadsheet.
function saveInstall(p) {
  try {
    var b = p || {};
    var sheetId = String(b.sheetId || '').trim();
    var key = String(b.geminiKey || '').trim();
    // Blank means "leave it alone". Writing the default here pinned a model
    // name into Script Properties that later silently overrode the code.
    var model = String(b.geminiModel || '').trim();

    // A pasted Sheets URL is the common mistake. Pull the id out of it.
    var m = sheetId.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
    if (m) sheetId = m[1];
    if (sheetId && !/^[a-zA-Z0-9_-]{20,}$/.test(sheetId)) {
      return { ok: false, error: 'That does not look like a spreadsheet ID' };
    }
    // Key is OPTIONAL. It only powers receipt scanning, and the install screen
    // offers a real skip. Blocking here would make that skip a dead button.
    var hasKey = !!key;

    var created = false;
    if (sheetId) {
      // Prove access first: a wrong id would otherwise fail later inside getSheet
      // and blame the wrong thing.
      try {
        SpreadsheetApp.openById(sheetId);
      } catch (e) {
        return { ok: false, error: 'Cannot open that spreadsheet. Check the link, and that you are signed in as its owner.' };
      }
    } else {
      // DEFAULT PATH: no sheet supplied, so make one. Asking a stranger to find
      // a spreadsheet id was the step they quit at. This creates an ordinary
      // Sheet in their own Drive and records its id.
      var fresh = SpreadsheetApp.create('Stackable');
      sheetId = fresh.getId();
      created = true;
    }

    var sp = PropertiesService.getScriptProperties();
    sp.setProperty('SHEET_ID', sheetId);
    if (hasKey) sp.setProperty('GEMINI_KEY', key);
    if (model) sp.setProperty('GEMINI_MODEL', model);

    getSheet();
    return { ok: true, sheetId: sheetId, created: created };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// URL of the user's own spreadsheet, for the Settings link. Read-only and
// derived from SHEET_ID; never trusts a client-supplied id.
/* ---------------- sheet ---------------- */

function getSheet() {
  var ss = resolveSheet_();
  if (!ss) {
    var id = prop('SHEET_ID');
    if (!id) throw new Error('SHEET_ID is not set in Script Properties');
    ss = SpreadsheetApp.openById(id);
  }
  // Deliberately NOT falling back to ss.getSheets()[0]. On a stranger's sheet
  // the first tab is often 'Sheet1' or 'Settings'; silently writing money into
  // whichever tab happens to be first looks like success and corrupts the
  // wrong place. Create the tab instead, and fail loudly if that is impossible.
  var sheet = ss.getSheetByName('Expenses');
  if (!sheet) {
    sheet = ss.insertSheet('Expenses');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    var first = sheet.getRange(1, 1, 1, 1).getValues()[0];
    if (String(first[0] || '') !== HEADERS[0]) {
      sheet.insertRowBefore(1);
      // Write the FULL header row. Setting only A1 left B1:K1 blank whenever
      // the width check below was already satisfied by the data rows.
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
    // ensure the id / updated_at columns exist on older sheets
    var width = sheet.getLastColumn();
    if (width < HEADERS.length) {
      sheet.getRange(1, width + 1, 1, HEADERS.length - width)
        .setValues([HEADERS.slice(width)]);
    }
  }
  return sheet;
}

function appendRow(row) {
  var sheet = getSheet();
  sheet.appendRow([
    row.date, row.vendor, row.amount, row.currency, row.category,
    row.payment, row.photo, row.confidence, row.logged_at,
    row.id || newId(), nowISO()
  ]);
}

function newId() {
  return 'r' + Date.now() + Math.floor(Math.random() * 1000);
}

/* ---------------- helpers ---------------- */


// Sheets may return real Date objects (if the column is formatted as dates)
// or plain strings. Normalise both to "yyyy-MM" / "yyyy-MM-dd".
function ymOf(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ(), 'yyyy-MM');
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2];
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TZ(), 'yyyy-MM');
  return s.substring(0, 7);
}

function isoOf(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ(), 'yyyy-MM-dd');
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TZ(), 'yyyy-MM-dd');
  return s;
}

// Numeric timestamp for sorting rows by date. Handles Date objects and strings.
function tsOf(v) {
  if (v instanceof Date) return v.getTime();
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return 0;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d.getTime();
  return 0;
}

// The ledger holds one currency. A receipt read as another (a bare "$" the
// model called USD, a MYR receipt from a trip) must not be added 1:1 into the
// home total with no trace, so the code is kept on the row and the list tags
// it. Missing code -> the user's own currency, never a hardcoded BND.
function ensureCategory(cat) {
  var c = String(cat || '').trim();
  if (!c) return 'Groceries & Household';
  var low = c.toLowerCase();
  if (CATEGORY_ALIASES[low]) return CATEGORY_ALIASES[low];
  for (var i = 0; i < CATEGORIES.length; i++) {
    if (CATEGORIES[i].toLowerCase() === low) return CATEGORIES[i];
  }
  return c;
}

function round2(v) {
  var n = Number(v);
  if (isNaN(n)) return NaN;
  return Math.round(n * 100) / 100;
}

// One timezone for every server-side date. Was hardcoded 'Asia/Brunei', which
// put a user in New York or Sydney a day (or a month) out from the phone that
// logged the row. The script's own timezone follows the account that deployed
// it, which is the closest thing Apps Script offers to "the user's zone".
var _sessionTz = null;
function TZ() {
  if (_sessionTz) return _sessionTz;
  try { _sessionTz = Session.getScriptTimeZone() || 'UTC'; }
  catch (e) { _sessionTz = 'UTC'; }
  return _sessionTz;
}

function prop(k) {
  return PropertiesService.getScriptProperties().getProperty(k);
}

/* ---------------- settings ---------------- */

// Settings tab: A=key, B=value, C=note (note ignored by the app).
// Missing tab / blank / non-numeric cell -> DEFAULT_SETTINGS fallback.
var DEBT_HEADERS = ['id','name','type','balance','rate','minimum','payment','notes'];

// DebtHistory tab: one row per debt per month, written by the app on each
// Debts-tab load. Must match the columns recordDebtSnapshot appends and
// readDebtHistory reads. This was referenced but never declared, which threw
// a ReferenceError that the optional-history try/catch swallowed -- so nothing
// was ever recorded and 'Cost so far' stayed empty forever.
var SNAP_HEADERS = ['ym','debtId','name','balance','rate','interest'];

// DebtPayments tab: one row per payment logged against a debt.
var PAY_HEADERS = ['date','debtId','name','amount','principal','interest'];

// Debts tab: one row per debt. Bitcoin-backed loans land here in v2 using the
// same columns - 'notes' is already reserved so no migration is needed.
function getDebtsSheet(ss) {
  var sheet = ss.getSheetByName('Debts');
  if (!sheet) {
    sheet = ss.insertSheet('Debts');
    sheet.appendRow(DEBT_HEADERS);
    sheet.getRange(1, 1, 1, DEBT_HEADERS.length).setFontWeight('bold');
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 90);
  }
  return sheet;
}

function readDebts() {
  var ss = resolveSheet_();
  if (!ss) {
    var id = prop('SHEET_ID');
    if (!id) return [];
    ss = SpreadsheetApp.openById(id);
  }
  var sheet = getDebtsSheet(ss);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  // Deliberately NOT wrapped in a catch that returns []. A swallowed error here
  // is indistinguishable from "this user has no debts", and the app then told
  // them so - wiping a two-debt screen. A throw reaches getDebts' own catch and
  // comes back as ok:false, which the client now treats as "keep what you have".
  {
    var vals = sheet.getRange(2, 1, last - 1, DEBT_HEADERS.length).getValues();
    var out = [];
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      if (!r[1]) continue;
      out.push({
        id: String(r[0] || ''),
        name: String(r[1] || ''),
        type: String(r[2] || 'fixed'),
        balance: Number(r[3]) || 0,
        rate: Number(r[4]) || 0,
        minimum: Number(r[5]) || 0,
        payment: Number(r[6]) || 0,
        notes: String(r[7] || '')
      });
    }
    return out;
  }
}

function writeDebts(list) {
  var ss = resolveSheet_();
  if (!ss) {
    var id = prop('SHEET_ID');
    if (!id) throw new Error('SHEET_ID is not set in Script Properties');
    ss = SpreadsheetApp.openById(id);
  }
  var sheet = getDebtsSheet(ss);
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, DEBT_HEADERS.length).clearContent();
  for (var i = 0; i < list.length; i++) {
    var d = list[i];
    sheet.appendRow([
      d.id || ('d' + Date.now() + i),
      String(d.name || ''),
      String(d.type || 'fixed'),
      Number(d.balance) || 0,
      Number(d.rate) || 0,
      Number(d.minimum) || 0,
      Number(d.payment) || 0,
      String(d.notes || '')
    ]);
  }
}

// Payoff maths for one debt. Straight amortisation on the closing balance:
//   monthlyInterest = balance * (rate/100) / 12
//   principal      = payment - monthlyInterest
// Returns nulls (not guesses) when the payment can never clear the balance.
// Sum of what the user actually pays toward debts each month. Read from the
// Debts tab; a missing tab simply means zero.
// Only REVOLVING debt is subtracted from free money. Instalment loans
// (mortgage, car loan) are living costs and already sit inside fixed_total,
// so subtracting them again would count them twice.
function totalDebtPayment() {
  var list = readDebts();
  var sum = 0;
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].type) === 'revolving') sum += Number(list[i].payment) || 0;
  }
  return round2(sum);
}

function payoffOf(debt) {
  var bal = Number(debt.balance) || 0;
  var rate = Number(debt.rate) || 0;
  var pay = Number(debt.payment) || 0;
  var minPay = Number(debt.minimum) || 0;
  var monthlyRate = rate / 100 / 12;
  var firstInterest = round2(bal * monthlyRate);

  // A payment at or below the interest never reduces the balance.
  if (pay <= firstInterest) {
    return {
      monthlyInterest: firstInterest, firstInterest: firstInterest,
      months: null, clearsOn: null, totalInterest: null,
      stuck: true, stuckReason: pay <= 0 ? 'no payment set' : 'payment does not cover interest'
    };
  }
  // Partially-covered minimum is still 'stuck' for the minimum, not the payment.
  var minStuck = minPay > 0 && minPay <= firstInterest;

  var remaining = bal;
  var months = 0;
  var totalInterest = 0;
  var guard = 1200; // 100 years ceiling
  while (remaining > 0.005 && months < guard) {
    var interest = remaining * monthlyRate;
    totalInterest += interest;
    var principal = pay - interest;
    if (principal <= 0) { months = null; break; }
    remaining -= principal;
    months++;
  }
  if (months === null || remaining > 0.005) {
    return { monthlyInterest: firstInterest, firstInterest: firstInterest,
             months: null, clearsOn: null, totalInterest: null, stuck: true,
             stuckReason: 'payment does not clear the balance' };
  }

  // Anchor to day 1. setMonth() on the 29th-31st rolls past a short month
  // (Jan 31 + 1 month became Mar 3), so payoff dates skipped February.
  var now0 = new Date();
  var d = new Date(now0.getFullYear(), now0.getMonth() + months, 1);
  var clearsOn = Utilities.formatDate(d, TZ(), 'MMM yyyy');

  return {
    monthlyInterest: firstInterest,
    firstInterest: firstInterest,
    months: months,
    clearsOn: clearsOn,
    totalInterest: round2(totalInterest),
    minStuck: minStuck,
    stuck: false,
    stuckReason: ''
  };
}

// What happens to a debt if the payment goes up by `extra` per month.
function extraPaymentEffect(debt, extra) {
  var base = payoffOf(debt);
  if (base.months === null) return null;
  var bumped = {};
  for (var k in debt) bumped[k] = debt[k];
  bumped.payment = (Number(debt.payment) || 0) + extra;
  var alt = payoffOf(bumped);
  if (alt.months === null) return null;
  return {
    extra: round2(extra),
    months: alt.months,
    monthsSaved: base.months - alt.months,
    interestSaved: round2((base.totalInterest || 0) - (alt.totalInterest || 0)),
    clearsOn: alt.clearsOn
  };
}

function getSnapSheet(ss) {
  var sheet = ss.getSheetByName('DebtHistory');
  if (!sheet) {
    sheet = ss.insertSheet('DebtHistory');
    sheet.appendRow(SNAP_HEADERS);
    sheet.getRange(1, 1, 1, SNAP_HEADERS.length).setFontWeight('bold');
    sheet.setColumnWidth(1, 90);
    sheet.setColumnWidth(3, 160);
  }
  return sheet;
}

function getPaySheet(ss) {
  var sheet = ss.getSheetByName('DebtPayments');
  if (!sheet) {
    sheet = ss.insertSheet('DebtPayments');
    sheet.appendRow(PAY_HEADERS);
    sheet.getRange(1, 1, 1, PAY_HEADERS.length).setFontWeight('bold');
    sheet.setColumnWidth(1, 90);
    sheet.setColumnWidth(3, 160);
  }
  return sheet;
}

// One month's sats-relevant debt reading. This USED to guess interest as
// balance*rate/12 -- a number the bank's statement routinely contradicted.
// Now it records the opening balance only; interest arrives from the user's
// statement (recordStatement). A row per debt per month, keyed by ym+debtId.
function recordDebtSnapshot(debs, ss) {
  try {
    if (!debs || !debs.length) return null;
    var sheet = getSnapSheet(ss);
    var ym = thisMonth();
    var last = sheet.getLastRow();
    var existing = {};
    if (last > 1) {
      var vals = sheet.getRange(2, 1, last - 1, SNAP_HEADERS.length).getValues();
      for (var i = 0; i < vals.length; i++) {
        // ymOf, not String(). Sheets coerces '2026-11' into a Date on write, so
        // String(cell) never matched and every Debts load appended duplicates.
        if (ymOf(vals[i][0]) === ym) existing[String(vals[i][1])] = true;
      }
    }
    for (var j = 0; j < debs.length; j++) {
      var d = debs[j];
      var bal = Number(d.balance) || 0;
      var rate = Number(d.rate) || 0;
      var id = String(d.id || d.name || '');
      if (existing[id]) continue;
      sheet.appendRow([ym, id, String(d.name || ''), bal, rate, 0]);
      // Pin the month cell to text, or Sheets turns it into a date and the
      // duplicate check fails again next month.
      sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('@').setValue(ym);
    }
    return null;
  } catch (e) {
    return String(e && e.message ? e.message : e);
  }
}

// Log a payment against a debt: append to DebtPayments and subtract the
// principal from that debt's balance so the payoff projection does not drift
// into fiction. amount = principal + interest; the split comes from what the
// user typed. Returns an error string or null.
function recordDebtPayment(p) {
  try {
    var b = p || {};
    var debtId = String(b.debtId || '');
    var amount = Number(b.amount);
    var principal = (b.principal === undefined) ? amount : Number(b.principal);
    var interest = (b.interest === undefined) ? 0 : Number(b.interest);
    if (!debtId) return 'Missing debt';
    if (isNaN(amount) || amount < 0) return 'Bad amount';
    if (isNaN(principal) || principal < 0) return 'Bad principal';
    if (isNaN(interest) || interest < 0) return 'Bad interest';

    var ss = resolveSheet_();
    if (!ss) {
      var id = prop('SHEET_ID');
      if (id) ss = SpreadsheetApp.openById(id);
    }
    if (!ss) return 'No spreadsheet';

    var sheet = getPaySheet(ss);
    var name = '';
    var debts = readDebts();
    for (var i = 0; i < debts.length; i++) {
      if (String(debts[i].id) === debtId || String(debts[i].name) === debtId) {
        name = String(debts[i].name || '');
        debts[i].balance = Math.max(0, round2((Number(debts[i].balance) || 0) - principal));
      }
    }
    if (!name) name = debtId;
    sheet.appendRow([todayISO(), debtId, name, round2(amount), round2(principal), round2(interest)]);
    writeDebts(debts);
    return null;
  } catch (e) {
    return String(e && e.message ? e.message : e);
  }
}

// The optional monthly reality check. The user types two numbers off their
// statement: how much they paid and what the bank charged in interest. We
// store it and let the plan card compare against what the plan predicted.
function recordStatement(p) {
  try {
    var b = p || {};
    var payment = Number(b.payment);
    var interest = Number(b.interest);
    if (isNaN(payment) || payment < 0) return { ok: false, error: 'Bad payment' };
    if (isNaN(interest) || interest < 0) return { ok: false, error: 'Bad interest' };

    var ss = resolveSheet_();
    if (!ss) {
      var id = prop('SHEET_ID');
      if (id) ss = SpreadsheetApp.openById(id);
    }
    if (!ss) return { ok: false, error: 'No spreadsheet' };

    var sheet = ss.getSheetByName('Statements') || (function(){
      var sh = ss.insertSheet('Statements');
      sh.appendRow(['ym','payment','interest']);
      sh.getRange(1,1,1,3).setFontWeight('bold');
      return sh;
    })();
    var last = sheet.getLastRow();
    var existing = -1;
    if (last > 1) {
      var vals = sheet.getRange(2, 1, last - 1, 1).getValues();
      var ym = thisMonth();
      for (var i = 0; i < vals.length; i++) if (ymOf(vals[i][0]) === ym) { existing = i + 2; break; }
    }
    if (existing > 0) {
      sheet.getRange(existing, 2, 1, 2).setValues([[round2(payment), round2(interest)]]);
    } else {
      sheet.appendRow([thisMonth(), round2(payment), round2(interest)]);
      // Pin column A to plain text: an unpinned '2026-11' is auto-parsed into a
      // Date cell, which is what made every row unmatched.
      var newRow = sheet.getLastRow();
      sheet.getRange(newRow, 1).setNumberFormat('@').setValue(thisMonth());
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

// Read the Statements tab for plan-vs-actual comparison.
function readStatements(ss) {
  try {
    var sheet = ss.getSheetByName('Statements');
    if (!sheet) return [];
    var last = sheet.getLastRow();
    if (last < 2) return [];
    var vals = sheet.getRange(2, 1, last - 1, 3).getValues();
    var out = [];
    for (var i = 0; i < vals.length; i++) {
      // Column A may be a Date (Sheets coerces '2026-11' into one), so diffing
      // String(v) against '2026-11' silently matched nothing. ymOf() handles
      // Date, ISO text and display text alike.
      out.push({ ym: ymOf(vals[i][0]), payment: Number(vals[i][1]) || 0, interest: Number(vals[i][2]) || 0 });
    }
    return out;
  } catch (e) { return []; }
}

function readDebtHistory(ss) {
  try {
    var sheet = ss.getSheetByName('DebtHistory');
    if (!sheet) return { rows: [], since: '' };
    var last = sheet.getLastRow();
    if (last < 2) return { rows: [], since: '' };
    var vals = sheet.getRange(2, 1, last - 1, SNAP_HEADERS.length).getValues();
    var rows = [];
    var since = '';
    for (var i = 0; i < vals.length; i++) {
      var ym = ymOf(vals[i][0]);
      if (!ym) continue;
      if (!since || ym < since) since = ym;
      rows.push({ ym: ym, debtId: String(vals[i][1] || ''), name: String(vals[i][2] || ''),
                  balance: Number(vals[i][3]) || 0, rate: Number(vals[i][4]) || 0,
                  interest: Number(vals[i][5]) || 0 });
    }
    return { rows: rows, since: since };
  } catch (e) {
    return { rows: [], since: '' };
  }
}

function getDebts() {
  try {
    var list;
    try {
      list = readDebts();
    } catch (readErr) {
      // Surface it. Returning ok:true with debts:[] here is what turned a
      // slow/failed Sheets read into "No debt" on screen.
      return { ok: false, error: 'Could not read the Debts tab: ' +
        ((readErr && readErr.message) ? readErr.message : String(readErr)) };
    }
    var totalPay = 0, totalBal = 0, totalInterest = 0;
    var enriched = [];
    for (var i = 0; i < list.length; i++) {
      var d = list[i];
      var p = payoffOf(d);
      totalPay += Number(d.payment) || 0;
      totalBal += Number(d.balance) || 0;
      if (p.totalInterest) totalInterest += p.totalInterest;
      var e = {};
      for (var k in d) e[k] = d[k];
      e.payoff = p;
      // The 'what if I paid more' line must match the FIRST preset button
      // (1.25x your payment), not a flat +50. A fixed 50 contradicted the
      // buttons rendered directly beneath it for any debt paying != 200/mo.
      var payNow = Number(d.payment) || 0;
      var bump = payNow > 0 ? round2(payNow * 0.25) : 0;
      e.whatIf = (p.months === null || bump <= 0) ? null : extraPaymentEffect(d, bump);
      enriched.push(e);
    }
    // Resolve the sheet ONCE, at the top level, before any try/catch.
    // Previously ss was assigned inside a try that also ran the history write;
    // any throw there left the statement read below with a null sheet, so the
    // plan card silently fell back to 'Your plan' even though the Statements
    // tab had rows. One resolution, no conditional visibility.
    var ss = resolveSheet_();
    if (!ss) {
      var ssId = prop('SHEET_ID');
      if (ssId) {
        try { ss = SpreadsheetApp.openById(ssId); } catch (e0) { ss = null; }
      }
    }

    // Record what we observe this month, then read back the whole clock.
    // The WRITE used to run on every single load: appendRow plus a
    // setNumberFormat per debt, which is two Sheets round-trips each and made
    // this call take tens of seconds once a sheet had real history. It only
    // needs to happen once per month, and only when something changed, so it
    // is now gated. A read-only visitor never pays for it.
    var hist = { rows: [], since: '' };
    var histErr = '';
    if (!ss) {
      histErr = 'No spreadsheet available to record history';
    } else {
      try {
        var snapKey = 'SNAP_' + thisMonth();
        var snapSig = 'SNAP_SIG_' + thisMonth();
        var wantSig = enriched.length ? enriched.map(function(x){
          return (x.id || x.name) + ':' + (Number(x.balance) || 0) + ':' + (Number(x.rate) || 0);
        }).join('|') : '';
        var sp2 = PropertiesService.getScriptProperties();
        // Only write when this month has no snapshot yet, or the observed
        // debts actually changed. getLastRow + one read is far cheaper than
        // the append cycle this replaces.
        if (enriched.length && (sp2.getProperty(snapKey) !== '1' ||
            sp2.getProperty(snapSig) !== wantSig)) {
          histErr = recordDebtSnapshot(enriched, ss) || '';
          if (!histErr) {
            sp2.setProperty(snapKey, '1');
            sp2.setProperty(snapSig, wantSig);
          }
        }
      } catch (e) {
        histErr = 'History write failed: ' + (e && e.message ? e.message : String(e));
      }
      try {
        hist = readDebtHistory(ss);
      } catch (e2) {
        hist = { rows: [], since: '' };
      }
    }

    // Plan vs actual: the plan is what the amortisation predicts for the
    // revolving debt this month; actual is what the user reports off their
    // statement. Never invent a bank figure.
    var planInterest = 0;
    for (var pi = 0; pi < enriched.length; pi++) {
      var pp = enriched[pi].payoff || {};
      if (pp.monthlyInterest) planInterest += Number(pp.monthlyInterest);
    }
    // Statement read: independent of every history path above. Takes the LAST
    // row matching this month, which is the most recent save.
    var stmts = [];
    var thisMonthStmt = null;
    var planErr = '';
    if (ss) {
      try {
        stmts = readStatements(ss);
        var wantYm = thisMonth();
        for (var si = 0; si < stmts.length; si++) {
          if (String(stmts[si].ym) === String(wantYm)) thisMonthStmt = stmts[si];
        }
      } catch (e3) {
        planErr = 'Statement read failed: ' + (e3 && e3.message ? e3.message : String(e3));
      }
    } else {
      planErr = 'No spreadsheet available for statements';
    }

    return {
      ok: true, debts: enriched,
      totalPayment: round2(totalPay),
      totalBalance: round2(totalBal),
      totalInterest: round2(totalInterest),
      history: hist.rows,
      trackedSince: hist.since,
      trackedError: histErr,
      trackedInterest: round2((function(){
        var ti = 0;
        for (var hh = 0; hh < hist.rows.length; hh++) ti += Number(hist.rows[hh].interest) || 0;
        return ti;
      })()),
      trackedMonths: hist.rows.length ? (function(){
        var seen = {};
        for (var q = 0; q < hist.rows.length; q++) seen[hist.rows[q].ym] = true;
        return Object.keys(seen).length;
      })() : 0,
      plan: {
        monthlyInterest: round2(planInterest),
        statement: thisMonthStmt,
        statements: stmts,
        error: planErr
      }
    };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function saveDebts(p) {
  try {
    var b = p || {};
    var list = b.debts || [];
    if (!list || typeof list.length !== 'number') return { ok: false, error: 'Bad payload' };
    var clean = [];
    var seenNames = {};
    // Existing ids by lowercased name, so a rename keeps its original id.
    var priorIds = {};
    var existing = readDebts();
    for (var e = 0; e < existing.length; e++) {
      var ex = existing[e];
      if (!ex || !ex.id) continue;
      priorIds[String(ex.name || '').trim().toLowerCase()] = String(ex.id);
    }
    for (var i = 0; i < list.length; i++) {
      var d = list[i] || {};
      var name = String(d.name || '').trim();
      if (!name) continue;
      var bal = Number(d.balance);
      var rate = Number(d.rate);
      var min = Number(d.minimum);
      var pay = Number(d.payment);
      if (isNaN(bal) || bal < 0) return { ok: false, error: 'Bad balance for ' + name };
      if (isNaN(rate) || rate < 0) return { ok: false, error: 'Bad rate for ' + name };
      if (isNaN(min) || min < 0) return { ok: false, error: 'Bad minimum for ' + name };
      if (isNaN(pay) || pay < 0) return { ok: false, error: 'Bad payment for ' + name };
      // Names are the user's handle on a debt, so they must be unique. Two cards
      // both called 'Visa' are ambiguous everywhere downstream (history, chips,
      // the pay editor), and the fix is a rename, not a merge.
      var nameKey = name.toLowerCase();
      if (seenNames[nameKey]) {
        return { ok: false, error: 'Two debts are both called "' + name + '". Give them different names.' };
      }
      seenNames[nameKey] = true;
      // Id resolution, in order: keep the id we were given, else reuse the id an
      // existing debt already had for this name, else mint one. Minting on every
      // save is what created a new id per rename and orphaned DebtHistory rows.
      var chosenId = String(d.id || '').trim();
      if (!chosenId) {
        var prior = priorIds[nameKey];
        chosenId = prior ? prior : ('d' + Date.now() + i);
      }
      clean.push({
        id: chosenId,
        name: name,
        type: (String(d.type) === 'revolving') ? 'revolving' : 'fixed',
        balance: round2(bal), rate: round2(rate),
        minimum: round2(min), payment: round2(pay),
        notes: String(d.notes || '')
      });
    }
    writeDebts(clean);
    return getDebts();
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function getSettingsSheet(ss) {
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['key', 'value', 'note']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    var seeds = [
      ['income_total', DEFAULT_SETTINGS.income_total, 'Take-home per month (all income)'],
      ['income_salary', DEFAULT_SETTINGS.income_salary, 'Your salary, after deductions'],
      ['income_wife', DEFAULT_SETTINGS.income_wife, 'Any other household income (partner, side income). Leave 0 if none.'],
      ['fixed_total', DEFAULT_SETTINGS.fixed_total, 'Rent/mortgage, school fees, utilities, internet, insurance'],
      ['variable_budget', DEFAULT_SETTINGS.variable_budget, 'Groceries & household + eating out + transport + health + fun + kids + gifts'],
      ['currency', DEFAULT_TEXT_SETTINGS.currency, '3-letter code: BND, USD, SGD, MYR, EUR, GBP...']
    ];
    for (var i = 0; i < seeds.length; i++) sheet.appendRow(seeds[i]);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 420);
  }
  return sheet;
}

function readSettings() {
  var out = {};
  for (var k in DEFAULT_SETTINGS) out[k] = DEFAULT_SETTINGS[k];
  for (var t in DEFAULT_TEXT_SETTINGS) out[t] = DEFAULT_TEXT_SETTINGS[t];
  try {
    var ss = resolveSheet_();
    if (!ss) {
      var id = prop('SHEET_ID');
      if (!id) return out;
      ss = SpreadsheetApp.openById(id);
    }
    var sheet = getSettingsSheet(ss);
    var last = sheet.getLastRow();
    if (last < 2) return out;
    var vals = sheet.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var k = String(vals[i][0] || '').trim();
      if (!k) continue;
      var raw = vals[i][1];
      if (raw === '' || raw === null || raw === undefined) continue;
      // Text settings are read verbatim; numbers only when they parse.
      if (k in DEFAULT_TEXT_SETTINGS) {
        var txt = String(raw).trim();
        if (k === 'currency') {
          var up = txt.toUpperCase();
          if (CURRENCY_SYMBOLS[up]) out.currency = up;
        } else {
          out[k] = txt;
        }
        continue;
      }
      var n = Number(raw);
      if (isNaN(n)) continue;
      out[k] = n;
    }
  } catch (err) {
    // Reading settings must never break the app.
  }
  // income_total is always derived, so a hand-edit to the Settings tab can
  // never leave the sheet and the app disagreeing.
  out.income_total = round2((Number(out.income_salary) || 0) + (Number(out.income_wife) || 0));
  return out;
}

function getSettings() {
  try {
    var s = readSettings();
    // A fresh install copies the template's numbers. Flag it so the UI can ask
    // the new user to set their own before trusting any figure.
    var configured = prop('SETTINGS_CONFIGURED') === '1';
    return { ok: true, settings: s, configured: configured };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// Writes the six numbers back to the Settings tab. Cell-by-cell, so a partially
// filled form doesn't wipe the rest. Rejects non-numeric or negative input
// rather than silently storing garbage.
function saveSettings(p) {
  try {
    var b = p || {};
    var body = b.settings || {};
    var keys = ['income_salary', 'income_wife', 'fixed_total', 'variable_budget'];
    var incoming = {};
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (!(k in body)) continue;
      var n = Number(body[k]);
      if (isNaN(n)) return { ok: false, error: 'Not a number: ' + k };
      if (n < 0) return { ok: false, error: 'Cannot be negative: ' + k };
      incoming[k] = round2(n);
    }
    // Text settings round-trip as strings; validated against the known codes.
    if ('currency' in body) {
      var cur = String(body.currency || '').toUpperCase();
      if (!CURRENCY_SYMBOLS[cur]) return { ok: false, error: 'Unknown currency: ' + cur };
      incoming.currency = cur;
    }

    // income_total is always derived, never typed - that's what makes the
    // breakdown consistent with the rate calculation.
    if ('income_salary' in incoming || 'income_wife' in incoming) {
      var cur = readSettings();
      var sal = ('income_salary' in incoming) ? incoming.income_salary : cur.income_salary;
      var wife = ('income_wife' in incoming) ? incoming.income_wife : cur.income_wife;
      incoming.income_total = round2(sal + wife);
    }

    var ss = resolveSheet_();
    if (!ss) {
      var id = prop('SHEET_ID');
      if (!id) throw new Error('SHEET_ID is not set in Script Properties');
      ss = SpreadsheetApp.openById(id);
    }
    var sheet = getSettingsSheet(ss);
    var last = Math.max(sheet.getLastRow(), 1);
    var vals = last > 1 ? sheet.getRange(2, 1, last - 1, 2).getValues() : [];
    var seen = {};
    for (var j = 0; j < vals.length; j++) {
      var key = String(vals[j][0] || '').trim();
      if (!key) continue;
      if (key in incoming) {
        sheet.getRange(j + 2, 2).setValue(incoming[key]);
        seen[key] = true;
      }
    }
    // Keys that weren't in the tab yet get appended.
    for (var k2 in incoming) {
      if (!seen[k2]) sheet.appendRow([k2, incoming[k2], '']);
    }
    // The Gemini key is a Script Property, not a Settings-tab value, so it can
    // never be read back to the page. A blank field therefore means "leave it
    // alone"; only a non-empty value overwrites. Clearing it needs a deliberate
    // delete in Apps Script if that is ever needed.
    // The client sends these INSIDE the settings object (body.geminiKey), not at
    // the top level. Reading only b.geminiKey meant a key pasted in Settings was
    // silently dropped - the save reported success and nothing was written.
    // Accept both shapes so an older client cannot break, and never write an
    // empty string over a working key.
    var inKey = String(body.geminiKey || b.geminiKey || '').trim();
    var inModel = String(body.geminiModel || b.geminiModel || '').trim();
    var sp = PropertiesService.getScriptProperties();
    if (inKey) sp.setProperty('GEMINI_KEY', inKey);
    if (inModel) sp.setProperty('GEMINI_MODEL', inModel);
    sp.setProperty('SETTINGS_CONFIGURED', '1');
    // hasKey/geminiModel come back so Settings can say "Key saved" without a
    // second round-trip. The key itself is never returned to the browser.
    return { ok: true, settings: readSettings(), configured: true,
             hasKey: !!prop('GEMINI_KEY'), geminiModel: prop('GEMINI_MODEL') || '' };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function todayISO() {
  return Utilities.formatDate(new Date(), TZ(), 'yyyy-MM-dd');
}

function thisMonth() {
  return Utilities.formatDate(new Date(), TZ(), 'yyyy-MM');
}

function nowISO() {
  return Utilities.formatDate(new Date(), TZ(), 'yyyy-MM-dd HH:mm:ss');
}

/* ---------------- embedded UI ---------------- */

var APP_HTML = `<!doctype html>
<html lang="en">
<head>
<base target="_top">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="theme-color" content="#fafafa">
<title>Stackable</title>
<style>
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{margin:0;padding:0;height:100%}
body{
  background:#fafafa;color:#16161a;
  font:15px/1.5 -apple-system,BlinkMacSystemFont,"SF Pro Text","Inter",system-ui,sans-serif;
  -webkit-font-smoothing:antialiased;
  padding-top:env(safe-area-inset-top);padding-bottom:calc(96px + env(safe-area-inset-bottom));
  overscroll-behavior:none;
}
.hdr{position:sticky;top:0;z-index:5;background:rgba(250,250,250,.86);
  backdrop-filter:blur(14px);border-bottom:1px solid #ececec;
  padding:18px 20px 12px;display:flex;justify-content:space-between;align-items:center}
.hdr h1{font-size:11px;margin:0;letter-spacing:1.6px;text-transform:uppercase;
  font-weight:600;color:#8a8a90}
.hdr-r{display:flex;align-items:center;gap:6px}
/* The SVG has fixed dimensions, so the button needs its own explicit hit
   area. Without this the tap target shrinks to the icon itself and misses. */
.gear{background:none;border:none;color:#9a9aa0;width:44px;height:44px;
  padding:0;margin:-4px -6px -4px 0;cursor:pointer;font-family:inherit;
  display:flex;align-items:center;justify-content:center;
  transition:color .15s;-webkit-tap-highlight-color:transparent}
.gear svg{display:block;pointer-events:none}
.gear:active{color:#16161a}
.gear:focus{outline:none}
.sum{display:flex;justify-content:space-between;align-items:baseline;
  padding:10px 0;border-top:1px solid #efefea;font-size:13px;color:#6b6b72}
.sum b{font-size:15px;color:#16161a;font-variant-numeric:tabular-nums}
.sum b.neg{color:#c8433a}
.sum b.pos{color:#1b6b4a}
.dsec{margin-top:22px;padding-top:18px;border-top:1px solid #efefea}
.dsec-h{font-size:10px;text-transform:uppercase;letter-spacing:1.5px;color:#8a8a90;font-weight:600;margin-bottom:6px}
.dcard{background:#fafaf8;border:1px solid #ecece8;border-radius:12px;padding:12px;margin-top:11px}
.drow{display:flex;gap:8px;align-items:center;margin-bottom:9px}
.dname{flex:1;margin:0!important;font-size:14px!important;padding:9px 11px!important}
.dx{background:none;border:none;color:#b0b0b6;font-size:22px;line-height:1;cursor:pointer;padding:0 6px;font-family:inherit}
.dx:active{color:#c8433a}
.dtype{width:100%;margin:0 0 10px!important;font-size:12px!important;padding:8px 10px!important;border-radius:9px}
.dgrid{display:grid;grid-template-columns:1fr 1fr;gap:9px}
.df{display:flex;flex-direction:column;gap:3px}
.dl{font-size:9px!important;text-transform:uppercase;letter-spacing:1.1px;color:#9a9aa0!important;font-weight:600;margin:0!important}
.df input{margin:0!important;font-size:13px!important;padding:8px 10px!important;border-radius:9px!important}
.dadd{width:100%;margin-top:11px;background:none;border:1px dashed #cfccd0;color:#6b6b72;
  border-radius:11px;padding:11px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit}
.dadd:active{border-color:#f7931a;color:#b8650a}
.unalloc{margin-top:10px;padding-top:10px;border-top:1px solid #efefea;
  font-size:12px;color:#8a8a90;display:flex;justify-content:space-between;align-items:baseline}
.unalloc b{font-size:13px;color:#16161a;font-variant-numeric:tabular-nums}
.unalloc.neg b{color:#c8433a}
.unalloc b.dim{color:#8a8a90;font-weight:500}
.dc{padding:12px 0;border-top:1px solid #efefea}
.dc:first-of-type{border-top:0;padding-top:2px}
/* One skeleton for every card in the app: whisper label, shout figure,
   quiet facts. The debt card used to lead with name-left/balance-right, which
   wrapped long bank names into two lines and broke the money apart. */
.dc-top{text-align:center}
.dc-n{display:block;font-size:11px;font-weight:600;text-transform:uppercase;
  letter-spacing:1.2px;color:#8a8a90;line-height:1.45}
.dc-b{font-size:15px;color:#16161a;font-variant-numeric:tabular-nums}
.dc-bal{font-size:30px;font-weight:600;letter-spacing:-1.1px;color:#16161a;
  font-variant-numeric:tabular-nums;text-align:center;margin-top:10px;line-height:1.1}
.dc-bal-sub{font-size:11px;color:#9a9aa0;text-align:center;margin-top:6px;
  font-variant-numeric:tabular-nums}
.dc-facts{display:flex;align-items:baseline;justify-content:space-between;gap:12px;
  font-size:12px;margin-top:9px;padding-top:9px;border-top:1px solid #f4f4f0}
.dc-fk{color:#9a9aa0;flex:none}
.dc-fv{color:#3a3a40;text-align:right;font-variant-numeric:tabular-nums;flex:1;min-width:0}
/* The pencil had NO rule at all, so it rendered as the browser default grey
   pill. This is the button it should always have been. */
.dcpay-edit{background:none;border:0;padding:0 0 0 5px;margin:0;font-size:11px;
  line-height:1;color:#c4c4ca;cursor:pointer;font-family:inherit;
  -webkit-tap-highlight-color:transparent;vertical-align:0}
.dcpay-edit:active{color:#f7931a}
.dc-line{font-size:12px;color:#7a7a80;margin-top:4px;line-height:1.45}
.dc-line.warn{color:#c8433a;font-weight:600}
.dc-sats{font-size:22px;font-weight:600;letter-spacing:-.6px;color:#16161a;
  margin-top:12px;font-variant-numeric:tabular-nums}
.dc-sats span{font-size:11px;font-weight:500;letter-spacing:0;color:#8a8a90}
.dc-cost{font-size:12px;color:#c8433a;font-weight:600;margin-top:4px;line-height:1.4}
.dc-apr{font-size:10px;color:#b4b4ba;margin-top:8px;letter-spacing:.3px}
.leak-caveat{font-size:11px;color:#8a8a90;margin-top:6px;font-style:italic}
.dcsugg{margin-top:8px;padding:7px 10px;background:#fdf6ec;border-radius:8px;
  color:#8a5a15;font-weight:500}
.pstamp{font-size:11px;color:#7f7f86;margin-top:10px;letter-spacing:.2px;
  font-variant-numeric:tabular-nums;cursor:pointer}
.pstamp.stale{color:#b06a3a}
.dc-note{font-size:10px;color:#a8a8ae;margin-top:12px;line-height:1.4}
.debtline{margin:0 0 12px 0;padding:13px 14px;background:#fdf6f4;border:1px solid #f3e3de;
  border-radius:11px;font-size:12px;line-height:1.5;color:#7a5c55;cursor:pointer;
  display:flex;align-items:center;justify-content:space-between;gap:12px;
  text-align:left;-webkit-tap-highlight-color:transparent;transition:background .15s}
.debtline>span:first-child{flex:1;min-width:0;display:block}
.debtline>span:first-child,.debtline b,.debtline .arrow{pointer-events:none}
.debtline .arrow{flex:none;display:flex;align-items:center;align-self:center;
  color:#b8a49e;font-size:15px;line-height:1}
.debtline:active{background:#faeeea}
.debtline b{color:#c0392b;font-variant-numeric:tabular-nums;font-size:13px}
.debtline .arrow{color:#b8a49e;font-size:15px;line-height:1;flex:none}
.starter{padding:2px 0}
.starter-t{font-size:18px;font-weight:600;letter-spacing:-.4px;color:#16161a;line-height:1.3}
.starter-p{font-size:13px;color:#6b6b72;line-height:1.55;margin-top:9px}
.starter-b{width:100%;margin-top:16px;background:#f7931a;color:#1f1305;border:none;
  border-radius:11px;padding:13px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit}
.starter-b:active{opacity:.85}
.payset{margin-top:12px;padding-top:12px;border-top:1px solid #f2f2ee}
.payset-h{font-size:10px;text-transform:uppercase;letter-spacing:1.1px;color:#9a9aa0;
  font-weight:600;margin-bottom:8px}
.payset-btns{display:flex;gap:7px}
.payset-btn{flex:1;background:#f7f7f4;border:1px solid #e8e8e2;color:#3c3c42;
  border-radius:9px;padding:9px 4px;font-size:12px;font-weight:600;cursor:pointer;
  font-family:inherit;font-variant-numeric:tabular-nums}
.payset-btn:active{background:#fdf2e4;border-color:#f6d9b4;color:#b8650a}
.payset-btn.on{background:#fdf2e4;border-color:#f7931a;color:#b8650a}
.payset-btn em{display:block;font-style:normal;font-size:9px;letter-spacing:.4px;
  color:#9a9aa0;margin-top:2px;font-weight:600}
/* Four buttons at the current sizing wrapped "BND 300.00" onto two lines and
   left the value looking broken. Tighter type, nowrap, and the currency rides
   inline. */
.payset-btn{font-size:11.5px;padding:8px 2px;white-space:nowrap;letter-spacing:-.2px}
.payset-btn.on em{color:#b8650a}
.leak-fix-red{color:#c8433a}
.leak-cost{margin-top:10px;font-size:11px;color:#9a9aa0;line-height:1.5}
.leak-cost b{color:#c8433a;font-weight:600;font-variant-numeric:tabular-nums}
.payset-out{font-size:11px;color:#6b6b72;margin-top:9px;line-height:1.5}
.payset-out b{color:#16161a;font-variant-numeric:tabular-nums}
.payset-out .green{color:#f7931a;font-weight:600}
/* Centred to match every other headline in the app: whisper label, shout
   figure, quiet sub-line. This card was left-aligned long after the hero and
   the Spent card went centred, which is why it read as a different page. */
.leak .lbl{color:#8a8a90;text-align:center}
.leak-big{font-size:34px;font-weight:600;letter-spacing:-1.4px;margin-top:8px;
  font-variant-numeric:tabular-nums;color:#16161a;text-align:center}
.leak-big span{font-size:12px;font-weight:500;letter-spacing:0;color:#8a8a90}
.leak-sub{font-size:12px;color:#7a7a80;margin-top:6px;line-height:1.5;text-align:center}
.leak-list{margin-top:16px;border-top:1px solid #efefea;padding-top:12px}
.leak-row{display:flex;justify-content:space-between;align-items:baseline;padding:7px 0;
  font-size:13px;color:#3c3c42}
.leak-row .leak-v{font-variant-numeric:tabular-nums;color:#16161a}
.leak-n{display:flex;align-items:center;gap:9px}
.leak-first{display:inline-flex;align-items:center;justify-content:center;
  width:18px;height:18px;border-radius:9px;background:#f7931a;color:#1a1205;
  font-size:10px;font-weight:700;font-style:normal}
.leak-idx{display:inline-flex;align-items:center;justify-content:center;
  width:18px;height:18px;border-radius:9px;background:#f0f0ec;color:#8a8a90;
  font-size:10px;font-weight:700;font-style:normal}
/* Was a grey rounded box holding two rows and a caveat: a card inside a card,
   which is what made the Debts page feel heavy. Now the same facts sit on the
   card itself with hairlines, matching .dc-facts and .foot2 elsewhere. */
/* One hairline, placed where .dc-facts puts its own: the card below this one
   uses 9px above the rule and 9px below, so a different 16/12 here made the
   two cards' dividers sit at different heights and the page read as two
   designs. */
.leak-fix{margin-top:9px;border-top:0;margin-bottom:-6px}
.leak-fix-h{font-size:11px;text-transform:uppercase;letter-spacing:1.2px;
  color:#8a8a90;font-weight:600;margin-bottom:9px}
/* No per-row rules. The debt card below draws ONE hairline (above "Paying")
   and lets the facts sit under it. Ruling every row here made the top card a
   little table and the card under it a list, which is the mismatch that kept
   reading as wrong. Match the single rule, and space rows like .dc-facts. */
/* Exactly ONE hairline between rows: the rule lives on the row FOLLOWING a
   row. Putting border-top on every row drew one above the first row too, so the
   eye saw a pair of lines around a single gap. Now: no rule above row 1, one
   rule above row 2, one above the caveat. Same shape as the card below. */
/* Copied verbatim from .dc-facts, the pattern that renders correctly on the
   card below: spacing only, NO border on the row itself. Every border I added
   here either vanished or doubled because the rules were fighting the
   container's own box. */
/* No borders anywhere on the rows or the container. The two dividers the user
   wants are drawn as explicit 1px elements between blocks, so exactly one line
   exists in each gap - CSS borders kept doubling because .leak-sub's line-height
   and the flex baseline pushed the rule against the text above it. */
/* Spacing copied from .dc-facts exactly, which is what the card below uses:
   9px between the rule and the text on each side. The rule element itself owns
   no margin - the rows' own padding provides the 9px - otherwise the gap came
   out ~20px against the card below's ~18px and the two read as different
   rhythms. */
/* Symmetric spacing around each rule: 9px above the text AND 9px below it.
   With padding-top only, every label sat jammed against the line above and
   hung low in its gap - the card below centres its text between the rules
   because .dc-facts carries margin-top and padding-top. Same here. */
.leak-fix.row{display:flex;justify-content:space-between;align-items:baseline;
  font-size:12px;color:#9a9aa0;gap:12px;padding:9px 0;border:0}
.leak-fix.row:first-child{padding-top:0}
.leak-rule{height:0;border-top:1px solid #f4f4f0;margin:0}
.leak-fix.row b{color:#3a3a40;font-variant-numeric:tabular-nums;text-align:right}
/* 'green' is the sats figure in this card: sats are bitcoin, so they take the
   bitcoin orange. The class name is kept to avoid touching markup. */
.leak-fix.row b.green{color:#f7931a}
.leak-fix-d{font-size:11px;color:#b4b4ba;margin:0;padding:9px 0 0;
  border:0;font-style:italic}
.leak .dc-note{color:#a8a8ae;border-top:1px solid #efefea;padding-top:10px;margin-top:14px}
.leak.free .leak-big{color:#f7931a}
/* The only red headline in the app: on the aggregate card, the monthly bleed.
   Per-debt cards below stay grey so the alarm does not become wallpaper. */
.leak-big.cost{color:#c8433a}
.welcome{padding-bottom:24px}
.wl-t{font-size:24px;font-weight:600;letter-spacing:-.8px;color:#16161a;line-height:1.2}
.wl-p{font-size:14px;color:#5c5c63;line-height:1.55;margin-top:10px}
.wl-list{margin-top:18px;border-top:1px solid #efefea;padding-top:14px}
.wl-row{display:flex;align-items:flex-start;gap:11px;padding:7px 0;font-size:13px;color:#3c3c42;line-height:1.45}
.wl-row i{display:inline-flex;align-items:center;justify-content:center;width:19px;height:19px;
  border-radius:10px;background:#fdf2e4;color:#b8650a;font-size:10px;font-weight:700;
  font-style:normal;flex:none;margin-top:1px}
.wl-note{font-size:11px;color:#8a8a90;line-height:1.5;margin-top:16px;padding-top:14px;
  border-top:1px solid #efefea}
.welcome .btn{margin-top:18px;width:100%}
.calib{margin-top:16px;padding:14px;background:#fdfaf5;border:1px solid #ede3d5;border-radius:13px}
.calib.offer{margin:0 0 12px;background:#fdf6ec;border-color:#f0dcc0}
.calib-t{font-size:12px;color:#4a4238;line-height:1.5}
.calib-t b{color:#16161a}
.calib-bar{height:4px;background:#f2e4d0;border-radius:2px;margin-top:11px;overflow:hidden}
.calib-bar i{display:block;height:100%;background:#f7931a;border-radius:2px}
.calib-d{font-size:10px;color:#9a9288;margin-top:6px;letter-spacing:.3px}
.calib-btns{display:flex;gap:8px;margin-top:12px}
.calib-yes{flex:2;background:#f7931a;color:#1f1305;border:none;border-radius:10px;padding:11px;
  font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
.calib-no{flex:1;background:none;border:1px solid #e6ded2;color:#6b6b72;border-radius:10px;
  padding:11px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit}
.sats{display:flex;align-items:baseline;gap:8px;margin-top:8px;flex-wrap:wrap}
.sats b{font-size:15px;color:#f7931a;font-variant-numeric:tabular-nums;letter-spacing:-.3px}
.sats span{font-size:11px;color:#9a9aa0}
.sugg{margin:16px 0 4px;padding:12px 14px;background:#fdf8f1;border:1px solid #f0e2cc;
  border-radius:12px;font-size:12px;color:#4a4238;line-height:1.5}
.sugg.muted{background:#f7f7f5;border-color:#ececE8;color:#8a8a90}
.sugg-t b{color:#16161a}
.sugg-d{color:#8a8175;margin-top:3px}
.sugg-b{margin-top:9px;background:#f7931a;color:#1f1305;border:none;border-radius:9px;
  padding:8px 12px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit}
.mnav{display:flex;align-items:center;gap:2px}
.mnav button{background:none;border:none;color:#9a9aa0;font-size:19px;line-height:1;
  padding:0 10px;cursor:pointer;font-family:inherit;transition:color .15s}
.mnav button:active{color:#16161a}
.mnav button:disabled{opacity:.2}
.mnav .lbl2{font-size:11px;letter-spacing:1.4px;text-transform:uppercase;
  color:#6b6b72;min-width:76px;text-align:center;font-weight:600}
/* Clearance for the fixed nav: it is ~56px tall plus safe area on top of the
   body, and the pad only reserved 48px. The last block on the Expenses tab
   ("50 of 74 shown") therefore sat underneath the tab bar and read as cut off.
   Reserve the nav's own height, not a round number. */
.pad{padding:24px 20px calc(88px + env(safe-area-inset-bottom))}
.card{background:#ffffff;border:1px solid #ececec;border-radius:18px;
  padding:22px 20px;margin-bottom:12px;position:relative;
  box-shadow:0 1px 2px rgba(16,18,22,.04),0 2px 8px rgba(16,18,22,.04)}
/* NEUTRAL WHITE SYSTEM.
   History: pass one was cool slate (read as dirty white), pass two warm sand
   (read as orange), pass three white cards on a warm cream canvas (the canvas
   itself read as orange on the phone). The lesson each time was the same -
   colour is doing the load-bearing work. So colour is removed from the system
   entirely. The surface is neutral, and hierarchy comes from three things only:
   border, shadow, spacing.
   This is the Linear/Stripe/Vercel approach on the light side: a near-white
   neutral canvas (#fafafa), pure white cards, one hairline, one soft shadow.
   No warm tint anywhere, so the hero's amber is the ONLY colour on the page
   and reads as an accent instead of as a wash. */
.card::after{content:none}
/* Hero is now a plain white card like the rest. Its hierarchy comes from
   position and type size alone, not colour. */
.card.hero-card{box-shadow:0 1px 2px rgba(16,18,22,.04),0 2px 8px rgba(16,18,22,.04)}
.card.hero-card::after{content:none}
/* Inner hairlines: neutral, a shade under pure white so they never glow. */
.card .leak-list,.card .oneoff{border-color:#f0f0f0}
.card .leak-rule{border-top-color:#f0f0f0}
.lbl{font-size:10px;text-transform:uppercase;letter-spacing:1.5px;
  color:#8a8a90;font-weight:600}
.big{font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","Inter",system-ui,sans-serif;
  font-size:44px;font-weight:600;letter-spacing:-1.8px;margin:10px 0 0;
  font-variant-numeric:tabular-nums}
.hero-top{text-align:center}
.hero-lbl{font-size:11px;text-transform:uppercase;letter-spacing:1.5px;
  color:#8a8a90;font-weight:600}
.hero-amt{font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","Inter",system-ui,sans-serif;
  font-size:38px;font-weight:600;letter-spacing:-1.6px;color:#16161a;
  font-variant-numeric:tabular-nums;line-height:1.1;cursor:pointer;
  text-align:center;margin-top:9px;white-space:nowrap}
.pace-scale{display:flex;justify-content:space-between;margin-top:6px;
  font-size:11px;color:#9a9aa0;font-variant-numeric:tabular-nums}
.pacelbl{text-align:center}
.hdr-l{min-width:0}
.hdr .pstamp{margin-top:2px;font-size:10px;color:#9a9aa0;cursor:default;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hero-amt.neg{color:#16161a}
.hero-alt{margin-top:7px;text-align:center;font-size:13px;color:#6b6b72;
  font-variant-numeric:tabular-nums;letter-spacing:.1px}
.hero-rate{margin-top:14px;padding-top:12px;border-top:1px solid #f0f0ec;
  text-align:center;font-size:12px;color:#8a8a90;letter-spacing:.1px}
.hero-rate b{color:#f7931a;font-weight:600;font-variant-numeric:tabular-nums}
.hero-stacked{margin-top:9px;text-align:center;font-size:12px;color:#6b6b72;
  letter-spacing:.1px;font-variant-numeric:tabular-nums}
.hero-stacked b{color:#16161a;font-weight:600}
.hero-rate.none{color:#9a9aa0}
/* The tap target is the whole row, so it gets real padding and a hit area that
   covers the full card width. Before, the row was one flex line and the long
   value string wrapped, which is also what broke the alignment. */
.debtrow{margin-top:14px;padding:12px 4px 4px;border-top:1px solid #f0f0ec;
  font-size:13px;cursor:pointer;-webkit-tap-highlight-color:transparent;
  border-radius:10px;transition:background .12s}
.debtrow:active{background:#f4f4f0}
/* Shared by BOTH money rows (.debtrow and .oneoff). These were scoped to
   .debtrow only, so the one-off row had no flex layout at all and its label
   and figure collapsed together into "One-offBND 417.91". */
.dr-top{display:flex;align-items:baseline;gap:10px}
.dr-k{flex:1;min-width:0;color:#6b6b72;white-space:nowrap}
.dr-v{font-weight:500;color:#16161a;font-variant-numeric:tabular-nums;
  white-space:nowrap}
.debtrow i{font-style:normal;color:#c4c4ca;font-size:15px;line-height:1;flex:none;
  align-self:center}
.dr-sub{margin-top:3px;font-size:11px;color:#9a9aa0;text-align:right;
  font-variant-numeric:tabular-nums}
/* Mirrors .hero-top/.hero-amt: label centred above, figure centred below.
   Two headline numbers on one screen have to share an alignment. */
.spent-top{text-align:center}
.spent-top .lbl{margin:0}
.spent-amt{font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","Inter",system-ui,sans-serif;
  /* Deliberately smaller than the hero's 38px. Two figures at the same size
     compete for the title; the hero answers the question the app exists for,
     this one answers the second question. */
  font-size:24px;font-weight:600;letter-spacing:-.7px;color:#16161a;margin-top:6px;
  font-variant-numeric:tabular-nums;cursor:pointer;text-align:center;line-height:1.15}
.spent-sub{font-size:11px;color:#9a9aa0;margin-top:6px;line-height:1.5}
.spent-sub.center{text-align:center}
/* One-off and Card payment share the same grid: .dr-top / .dr-k / .dr-v /
   .dr-sub. Kept as its own class only for the spacing above the first rule. */
.oneoff{margin-top:14px;padding:12px 4px 4px;border-top:1px solid #f0f0ec;
  font-size:13px;cursor:pointer;-webkit-tap-highlight-color:transparent;
  border-radius:10px;transition:background .12s}
.oneoff:active{background:#f4f4f0}
.oneoff.open .oo-i::after{border-top:0;border-bottom:5px solid currentColor;
  margin-top:-2px}
/* Same chevron geometry as the hero disclosure, rotated by the .open state. */
.oo-i{flex:none;width:12px;height:12px;display:flex;align-items:center;
  justify-content:center;color:#c4c4ca;pointer-events:none}
.oo-i::after{content:'';border-left:4px solid transparent;border-right:4px solid transparent;
  border-top:5px solid currentColor}
.oo-list{margin-top:10px;padding-top:8px;border-top:1px solid #f4f4f0}
.oo-row{display:flex;align-items:baseline;gap:10px;padding:5px 0;font-size:12px}
.oo-row .oo-d{flex:none;width:44px;color:#9a9aa0;font-variant-numeric:tabular-nums}
.oo-row .oo-v{flex:1;min-width:0;color:#3a3a40;overflow:hidden;
  text-overflow:ellipsis;white-space:nowrap}
.oo-row .oo-a{flex:none;color:#16161a;font-variant-numeric:tabular-nums}
.oo-empty{margin-top:8px;font-size:12px;color:#9a9aa0}
.cathead{display:flex;align-items:baseline;justify-content:space-between;gap:10px;
  margin-bottom:14px}
.cathead .lbl{margin:0}
.rcpt{font-size:11px;color:#9a9aa0;font-variant-numeric:tabular-nums}
.rcpt b{color:#6b6b72;font-weight:500}
.foot2{display:grid;grid-template-columns:1fr 1fr;gap:10px 12px;margin-top:18px;
  padding-top:14px;border-top:1px solid #f0f0ec}
.f2k{font-size:10px;text-transform:uppercase;letter-spacing:1.3px;color:#9a9aa0;font-weight:600}
.f2v{font-size:14px;color:#16161a;margin-top:4px;font-variant-numeric:tabular-nums}
.f2v span{font-size:11px;color:#9a9aa0;margin-left:4px}
.f2v.none{color:#9a9aa0}
.big.hero{font-size:42px;letter-spacing:-1.8px}
.sub{font-size:13px;color:#8a8a90;margin-top:10px;letter-spacing:.1px}
.sub b{color:#f7931a;font-weight:600;font-variant-numeric:tabular-nums}
.sub.over b{color:#c8433a}
.hero-card{padding-bottom:16px}
/* HERO, PLAIN WHITE (revision 4). Amber wash, amber border, amber glow and the
   orange top rule are all gone by request: the tinted hero plus the cream page
   was reading as "too much orange". The hero now uses the exact same surface
   treatment as every other card, so nothing on the page competes with it. Its
   standing comes from size and placement - it is first, and its figure is 50px.
   Amber remains only as a working accent: the stack-rate number, the disclosure
   chevron, and the positive figure, never as a background. */
.hero-card{background:#ffffff;
  border:1px solid #ececec;border-radius:18px;
  padding:26px 22px 18px;position:relative;overflow:hidden;
  box-shadow:0 1px 2px rgba(16,18,22,.04),0 2px 8px rgba(16,18,22,.04)}
.hero-card .hero-lbl{font-size:11px;letter-spacing:4.5px;font-weight:700;
  color:#e08600;text-transform:uppercase;padding-left:4.5px}
.hero-card .hero-lbl.brand{font-size:15px;letter-spacing:4.2px;font-weight:800;
  color:#16161a;padding-left:4.2px;line-height:1.3}
.hero-card .hero-amt{font-size:50px;font-weight:700;letter-spacing:-2.4px;
  color:#16161a;margin-top:11px}
.hero-card .hero-amt.neg{color:#c0392b}
.hero-card .hero-alt{color:#6b6b72;margin-top:9px}
.hero-card .hero-stacked{color:#6b6b72;margin-top:12px}
.hero-card .hero-stacked b{color:#16161a}
.hero-card .hero-rate{margin-top:16px;padding-top:14px;
  border-top:1px solid rgba(247,147,26,.18);color:#8a8a90}
.hero-card .hero-rate b{color:#e08600;font-weight:700}
.hero-card .hero-rate.none{color:#9a9aa0}
.hero-card .hcalc{color:#6b6b72;border-top:1px solid rgba(247,147,26,.18)}
.hero-card .hdisc{color:#c4b099;margin-top:6px}
.hero-card .hdisc:active{color:#f7931a}
.pace{position:relative;height:8px;background:#ece4d6;border-radius:4px;margin-top:8px;
  overflow:visible}
/* The tick is the point of the whole bar: where you should be today. At 2x4px
   against a 4px track it vanished on a phone; the white halo makes it read as
   a marker sitting on top of the fill rather than a gap in it. */
.pace .tick{position:absolute;top:-5px;width:2px;height:18px;background:#16161a;
  border-radius:1px;transform:translateX(-1px);box-shadow:0 0 0 1.5px #fff}
.pace-cap{font-size:10px;text-transform:uppercase;letter-spacing:1.3px;color:#9a9aa0;
  font-weight:600;margin-top:18px;text-align:center}
.pacelbl{margin-top:9px;font-size:11px;letter-spacing:.2px;text-align:center}
.pacelbl.ok{color:#1b6b4a}
.pacelbl.behind{color:#96702a}
.grid4{display:grid;grid-template-columns:1fr 1fr;gap:14px 12px;align-items:start}
.listhead{display:flex;align-items:baseline;justify-content:space-between;gap:12px;
  padding:2px 4px 12px}
.listhead b{font-size:22px;font-weight:600;letter-spacing:-.6px;color:#16161a;
  font-variant-numeric:tabular-nums}
.listhead span{font-size:11px;color:#9a9aa0;letter-spacing:.2px}
.cell .ck{font-size:10px;text-transform:uppercase;letter-spacing:1.3px;color:#9a9aa0;font-weight:600}
.cell .cv{font-size:15px;color:#16161a;margin-top:5px;font-variant-numeric:tabular-nums;
  font-family:-apple-system,BlinkMacSystemFont,inter,system-ui,sans-serif}
.cell .cv .green{color:#f7931a}
.cell .cs{font-size:10px;color:#9a9aa0;margin-top:3px;letter-spacing:.2px}
.cell.wide{grid-column:1 / -1;padding-bottom:12px;margin-bottom:12px;
  border-bottom:1px solid #f2f2ee}
.cell.wide .cv{font-size:20px;font-weight:600;letter-spacing:-.4px;margin-top:4px}
.cell.wide .cs{font-size:10px;color:#9a9aa0;margin-top:5px;line-height:1.5}
.hdisc{margin:2px auto 0;background:none;border:0;padding:0;width:32px;height:22px;
  color:#b4b4ba;cursor:pointer;font-family:inherit;display:flex;align-items:center;
  justify-content:center;-webkit-tap-highlight-color:transparent;
  transition:color .15s}
.hdisc:active{color:#16161a}
.hdisc-i{width:15px;height:15px;display:flex;align-items:center;
  justify-content:center;flex:none;pointer-events:none}
.hdisc-i::after{content:'';border-left:4px solid transparent;border-right:4px solid transparent;
  border-top:5px solid currentColor;margin-top:2px}
.hdisc.open .hdisc-i::after{border-top:0;border-bottom:4.5px solid currentColor;margin-top:-1px}
.hcalc{font-size:11px;color:#7a7a80;line-height:1.55;margin-top:8px;
  padding:9px 2px 0;background:none;border-top:1px solid #f0f0ec}
.big.hero.neg{color:#c0392b}
.pace .fill{display:block;height:100%;background:#f7931a;border-radius:4px;
  transition:width .3s cubic-bezier(.4,0,.2,1)}
.bar i.warn{background:#c98a2e}
.bar i.over{background:#c8433a}
.stat{display:flex;justify-content:space-between;align-items:baseline;
  padding:13px 0;border-top:1px solid #ededE9}
.stat:first-of-type{border-top:0}
.stat .k{font-size:11px;text-transform:uppercase;letter-spacing:1.2px;
  color:#8a8a90;font-weight:600}
.stat .v{font-family:-apple-system,BlinkMacSystemFont,inter,system-ui,sans-serif;
  font-size:15px;font-variant-numeric:tabular-nums;color:#16161a;font-weight:500}
.stat .v.green{color:#1b6b4a}
.cats{display:flex;flex-direction:column}
.cat{padding:11px 0;border-top:1px solid #ededE9;font-size:14px}
.cat:first-child{border-top:0;padding-top:4px}
.cat-top{display:flex;align-items:baseline;justify-content:space-between;gap:11px}
.cat-bar{height:3px;background:#f0f0ec;border-radius:2px;margin-top:7px;overflow:hidden}
.cat-bar i{display:block;height:100%;background:#c9c9cd;border-radius:2px}
.cat:first-child .cat-bar i{background:#f7931a}
.cat .dot{width:6px;height:6px;border-radius:50%;background:#c4c4c8;flex:none}
.cat .nm{flex:1;color:#2a2a30;letter-spacing:.1px}
.cat .vl{font-family:-apple-system,BlinkMacSystemFont,inter,system-ui,sans-serif;
  font-size:13px;color:#6b6b72;font-variant-numeric:tabular-nums}
.row{display:flex;justify-content:space-between;gap:12px;padding:12px 0;border-bottom:1px solid #f0f0f0;font-size:15px}
.row:last-child{border-bottom:0}
.row .l{min-width:0;flex:1}
.row .v{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.row .t{color:#8a8a90;font-size:13px;margin-top:2px}
.row .a{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums;font-weight:500}
.empty{color:#8a8a90;text-align:center;padding:34px 0;font-size:14px}
nav{position:fixed;left:0;right:0;bottom:0;z-index:6;display:flex;
  background:rgba(250,250,250,.9);backdrop-filter:blur(10px);border-top:1px solid #ececec;
  padding:8px 8px calc(8px + env(safe-area-inset-bottom))}
nav button{flex:1;background:none;border:0;color:#8a8a90;font:11px/1.2 inherit;
  padding:8px 0;cursor:pointer;font-family:inherit}
nav button.on{color:#f7931a}
nav button span{display:block;font-size:10px;letter-spacing:.6px;text-transform:uppercase;margin-top:2px}
nav .navadd{flex:none;width:auto;min-width:52px;padding:6px 16px;color:#f7931a;
  display:flex;align-items:center;justify-content:center;-webkit-tap-highlight-color:transparent}
nav .navadd svg{display:block;pointer-events:none}
.fab{position:fixed;right:18px;bottom:calc(84px + env(safe-area-inset-bottom));z-index:7;
  width:58px;height:58px;border-radius:50%;background:#f7931a;color:#1a1205;border:0;
  font-size:30px;line-height:1;font-weight:300;cursor:pointer;
  box-shadow:0 8px 22px rgba(247,147,26,.3);display:flex;align-items:center;justify-content:center}
.mask{position:fixed;inset:0;background:rgba(0,0,0,.62);z-index:20;display:flex;align-items:flex-end}
.sheet{background:#ffffff;width:100%;max-height:92vh;overflow:auto;border-radius:20px 20px 0 0;
  padding:20px 18px calc(20px + env(safe-area-inset-bottom))}
.sheet h2{font-size:17px;margin:0 0 4px;font-weight:600}
.sheet .hint{color:#8a8a90;font-size:12px;line-height:1.5;margin:6px 0 0}
/* Every field is label -> input -> hint, so the rhythm is fixed here rather
   than patched per field: 6px from field to its hint, 18px from a hint (or a
   previous field) to the next label. The old .hint + label rule made a
   hint-followed label sit 22px away while a field-followed label sat 20px,
   which is what made the column look ragged. */
.sheet .hint + label{margin-top:18px}
label{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:#8a8a90;margin:18px 0 7px}
/* Label directly above a field, or a label straight after another label, keeps
   one gap. Negative margins on inline hints were the other source of drift. */
.sheet label + .hint{margin-top:0}
input,select{width:100%;background:#ffffff;border:1px solid #d9d9d4;color:#16161a;
  border-radius:12px;padding:13px;font-size:16px;font-family:inherit;-webkit-appearance:none;appearance:none}
select{background-image:linear-gradient(45deg,transparent 50%,#8a8a90 50%),linear-gradient(135deg,#8a8a90 50%,transparent 50%);
  background-position:calc(100% - 20px) 50%,calc(100% - 14px) 50%;background-size:6px 6px,6px 6px;background-repeat:no-repeat}
.btn{width:100%;margin-top:18px;background:#f7931a;color:#1f1305;border:0;border-radius:12px;
  padding:15px;font-size:16px;font-weight:600;font-family:inherit;cursor:pointer}
.btn:disabled{opacity:.55}
.btn.alt{background:transparent;color:#16161a;border:1px solid #d9d9d4;margin-top:10px;font-weight:500}
.btn.ghost{background:#f0f0ec;color:#16161a;margin-top:10px;font-weight:500}
.prev{width:100%;border-radius:12px;display:block;margin-top:12px}
.err{background:#fdecec;border:1px solid #f3c6c6;color:#b3322b;padding:14px;border-radius:12px;
  font-size:14px;margin-bottom:14px;white-space:pre-wrap;word-break:break-word}
.toast{position:fixed;left:50%;transform:translateX(-50%);bottom:calc(96px + env(safe-area-inset-bottom));
  z-index:30;background:#16161a;color:#ffffff;padding:12px 18px;border-radius:12px;font-size:14px;
  max-width:88%;text-align:center}
.toast.bad{background:#c8433a;color:#ffffff}
.sep{text-align:center;color:#8a8a90;font-size:13px;margin:20px 0 0}

.sw-wrap{padding:0;overflow:hidden}
.sw{position:relative;overflow:hidden;border-bottom:1px solid #ededE9}
.sw:last-child{border-bottom:0}
.sw-actions{position:absolute;top:0;right:0;bottom:0;display:flex;transform:translateX(100%);transition:transform .22s ease}
.sw.open .sw-actions{transform:translateX(0)}
.sw-actions button{border:0;color:#fff;font:15px/1 inherit;font-family:inherit;
  padding:0 18px;cursor:pointer;font-weight:500}
.sw-edit{background:#3a7bd5}
.sw-del{background:#e0352b;border-radius:0}
.sw-face{background:#ffffff;position:relative;z-index:2;transform:translateX(0);
  transition:transform .22s ease;padding-left:16px;padding-right:16px;border-bottom:0}
.sw.open .sw-face{transform:translateX(-140px)}
.sw-face .l{min-width:0;flex:1}
.ehead{font-size:17px;margin:0 0 4px;font-weight:600}
</style>
</head>
<body>
<div id="app"><div class="pad"><div class="empty">Loading...</div></div></div>
<script>
(function(){
"use strict";

var CATS = ["Groceries & Household","Eating Out","Transport","Health","Fun","Kids","Gifts","Sats","One-off","Reimbursed"];

// Currency codes duplicated into the client scope so dropdowns can build
// themselves. Keep this list in sync with CURRENCY_SYMBOLS on the server.
var CCY = "BND,USD,SGD,MYR,EUR,GBP,AUD,CAD,JPY,INR,IDR,PHP,THB,CNY,HKD,KRW,TWD,NZD,CHF,AED,SAR,ZAR,NGN,KES".split(",");

// Client copy of the server's DEFAULT_GEMINI_MODEL. The two scopes are separate
// (this bundle cannot see server vars), so it is duplicated on purpose. Keep in
// sync with DEFAULT_GEMINI_MODEL near the top of the .gs file.
var CLIENT_GEMINI_MODEL = "gemini-3.5-flash-lite";
// COLORS / BUDGET / INCOME were leftovers from the first prototype.
// BUDGET and INCOME once held real figures - removing them is a safety fix.

var state = {
  tab: "summary",
  month: todayMonth(),
  summary: null,
  rows: null,
  listLimit: 50,
  busy: false,
  pending: 0,   // in-flight request count; busy is always derived from this
  saving: false, // the Settings form's own save, not any background load
  error: null,
  sheet: false,
  editRow: null,
  swipeOpen: null,
  photo: null,
  scanning: false,
  scanError: null,
  draft: null,
  toast: null,
  toastBad: false,
  calibDismissed: false,
  suggestion: null,
  btc: null,
  settingsOpen: false,
  settingsDraft: null,
  needsSetup: false,
  needsWelcome: false,
  needsInstall: false,
  payEdit: null,
  installDraft: null,
  heroOpen: false,
  listTotal: 0,
  listCount: 0,
  listHasMore: false,
  denom: 'fiat', // 'fiat' | 'sats' — toggled by tapping a big figure
  debts: null,  // null = not fetched yet; [] = fetched, none
  simPay: {},   // debt id -> payment the user is trying on
  onboardChecked: false,
  trackedSince: '',
  trackedInterest: 0,
  trackedMonths: 0,
  trackedError: '',
};

var toastTimer = null;

function todayMonth(){
  var d = new Date();
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1);
}
// CLIENT copy. The server's todayISO() (formatted in the account timezone)
// wins on the server; this one is the phone's local date inside the UI.
function todayISO(){
  var d = new Date();
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}
function pad2(n){ return (n < 10 ? "0" : "") + n; }

function daysInMonth(ym){
  var p = String(ym).split("-");
  var y = parseInt(p[0], 10), m = parseInt(p[1], 10);
  if (!y || !m) return 30;
  return new Date(y, m, 0).getDate();
}

var CUR = "BND";

function round2c(v){ var n = Number(v); return isNaN(n) ? 0 : Math.round(n * 100) / 100; }

// USD-labelled price for the BTC anchor line. Always a dollar sign, never
// CUR (which is the user's spending currency). Whole dollars to 0dp above
// 1000 so the line stays readable on a phone.
function usdLabel(n){
  var v = Number(n) || 0;
  var whole = v >= 1000 ? Math.round(v) : Math.round(v * 100) / 100;
  var s = String(whole);
  var parts = s.split('.');
  parts[0] = parts[0].replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',');
  return '$' + parts.join('.');
}

function money(n, cur){
  var v = Number(n);
  if (isNaN(v)) v = 0;
  var neg = v < 0;
  v = Math.abs(v);
  var s = v.toFixed(2);
  var parts = s.split(".");
  parts[0] = parts[0].replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");
  // Optional currency code so a foreign row can never be read as home currency.
  // Amounts are NOT converted anywhere: an untagged USD 50 would read as BND 50.
  var c = String(cur || CUR || "BND").toUpperCase();
  return (neg ? "-" : "") + c + " " + parts.join(".");
}

// ---- denomination (fiat <-> sats) ----
// One flag, one truth: it lives on state.denom and is mirrored to localStorage
// so a reload doesn't snap back to fiat mid-session.
var DENOM_KEY = 'receipts.denom';

// Price of one BTC in the ledger currency (CUR). getBtcPrice returns that
// figure in price_bnd using the live FX table, so this stays the single
// entry point every sats calculation goes through.
function priceBnd(){
  var b = state.btc;
  if (!b || !b.ok || !b.price_bnd) return 0;
  return Number(b.price_bnd) || 0;
}

function satsOf(bnd){
  var p = priceBnd();
  if (!p) return null;
  return Math.round((Number(bnd) || 0) / p * 100000000);
}

// Whole sats, grouped. No decimals: 205,761 sats reads better than .43
function satsText(n){
  var v = Math.round(Math.abs(Number(n) || 0));
  var sign = (Number(n) || 0) < 0 ? '-' : '';
  return sign + v.toLocaleString() + ' sats';
}

// The formatter every big figure goes through. Falls back to fiat whenever the
// price isn't loaded, so a tap can never produce a blank or invented number.
function amt(n){
  if (state.denom !== 'sats') return money(n);
  var s = satsOf(n);
  if (s === null) return money(n);
  return satsText(s);
}

function canShowSats(){
  return state.denom === 'sats' && priceBnd() > 0;
}

function loadDenom(){
  try {
    var v = window.localStorage.getItem(DENOM_KEY);
    if (v === 'sats' || v === 'fiat') state.denom = v;
  } catch (e) { /* private mode: stay on fiat */ }
}

function saveDenom(){
  try { window.localStorage.setItem(DENOM_KEY, state.denom); } catch (e) {}
}

/// toggles every big figure at once
function toggleDenom(){
  if (state.denom === 'sats' && !priceBnd()) {
    // Nothing to convert to yet: pull the price, stay on fiat, say so.
    if (!state.btc) loadBtcPrice();
    toast('Price still loading — try again in a moment');
    return;
  }
  state.denom = (state.denom === 'sats') ? 'fiat' : 'sats';
  saveDenom();
  render();
}

function esc(s){
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}

function toast(msg, bad){
  state.toast = msg;
  state.toastBad = !!bad;
  render();
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(function(){
    state.toast = null;
    toastTimer = null;
    render();
  }, 3400);
}

function fail(prefix, e){
  var m = (e && e.message) ? e.message : String(e);
  state.error = prefix + ": " + m;
  // Never assign busy directly: it is derived from the pending counter.
  state.pending = 0;
  state.busy = false;
  render();
}

// The ONLY place state.busy is derived. Every request that starts or ends must
// go through here, or the counter and the flag drift apart and the UI locks up
// showing "Saving..." forever.
function busyStart(){
  state.pending = (state.pending || 0) + 1;
  state.busy = true;
}
function busyEnd(){
  state.pending = Math.max(0, (state.pending || 1) - 1);
  state.busy = state.pending > 0;
  if (!state.busy) state.waitingOn = null;
}

/* ---------- rendering ---------- */

function render(){
  var root = document.getElementById("app");
  if (!root) return;
  // busy is ALWAYS derived. If it ever disagrees with the counter (an abandoned
  // request, a callback that never fired), the counter wins - otherwise the UI
  // paints "Saving..." forever with nothing behind it.
  state.busy = (state.pending || 0) > 0;
  try {
    renderInto(root);
  } catch (e) {
    // A throw inside any panel used to leave the whole tree half-built: the
    // gear appeared to do nothing and every later tap hit a dead DOM, with the
    // real error visible only in a devtools console the phone user cannot
    // open. Show the error on screen instead of dying silently.
    root.innerHTML = '<div class="pad"><div class="err">Something broke: ' +
      esc((e && e.message) ? e.message : String(e)) + '</div></div>';
  }
}

function renderInto(root){
  var parts = [];
  parts.push(headerHtml());
  parts.push('<div class="pad">');
  if (state.error) parts.push('<div class="err">' + esc(state.error) + '</div>');
  parts.push(state.needsWelcome ? welcomeHtml()
           : (state.tab === "summary" ? summaryHtml()
           : (state.tab === "debts" ? debtTabHtml() : listHtml())));
  parts.push('</div>');
  parts.push(navHtml());
  if (state.sheet) parts.push(sheetHtml());
  if (state.settingsOpen) parts.push(settingsSheetHtml());
  if (state.payEdit !== null && state.payEdit !== undefined) parts.push(payEditorHtml());
  if (state.needsInstall) parts.push(installSheetHtml());
  if (state.editRow) parts.push(editSheetHtml());
  if (state.draft) parts.push(draftSheetHtml());
  if (state.toast) {
    parts.push('<div class="toast' + (state.toastBad ? " bad" : "") + '">' + esc(state.toast) + '</div>');
  }
  // Re-rendering replaces the whole tree, which resets every scroll container to
  // 0. Settings re-renders whenever debts load, so scrolling it snapped back to
  // the top. Capture each sheet's scrollTop and restore it after the swap.
  var keepScroll = [];
  var oldSheets = root.querySelectorAll('.sheet');
  for (var si = 0; si < oldSheets.length; si++) keepScroll.push(oldSheets[si].scrollTop);

  root.innerHTML = parts.join("");

  var newSheets = root.querySelectorAll('.sheet');
  for (var sj = 0; sj < newSheets.length && sj < keepScroll.length; sj++) {
    if (keepScroll[sj]) newSheets[sj].scrollTop = keepScroll[sj];
  }
}

function headerHtml(){
  var title = state.tab === "summary" ? "This Month"
            : (state.tab === "debts" ? "Debts" : "Expenses");
  return '<div class="hdr"><div class="hdr-l"><h1>' + title + '</h1>' +
    priceStampHtml() + '</div>' +
    '<div class="hdr-r">' + monthNavHtml() +
    '<button class="gear" data-act="settings" aria-label="Settings">' +
      '<svg viewBox="0 0 24 24" width="21" height="21" fill="none" ' +
      'stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' +
      '<circle cx="12" cy="12" r="3.1"/>' +
      '<path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>' +
      '</svg></button>' +
    '</div></div>';
}

function monthLabel(ym){
  var p = String(ym).split("-");
  var names = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var m = parseInt(p[1], 10);
  if (!m || m < 1 || m > 12) return ym;
  return names[m-1] + " " + p[0];
}

function shiftMonth(ym, delta){
  var p = String(ym).split("-");
  var y = parseInt(p[0], 10), m = parseInt(p[1], 10) - 1 + delta;
  y += Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  return y + "-" + pad2(m + 1);
}

function monthNavHtml(){
  // next arrow disabled once we're on the current month
  var atNow = state.month === todayMonth();
  return '<div class="mnav">' +
    '<button data-act="prevMonth" data-d="-1" aria-label="Previous month">&lsaquo;</button>' +
    '<span class="lbl2">' + esc(monthLabel(state.month)) + '</span>' +
    '<button data-act="nextMonth" data-d="1" aria-label="Next month"' + (atNow ? ' disabled' : '') + '>&rsaquo;</button>' +
    '</div>';
}

function summaryHtml(){
  if (!state.summary) return '<div class="empty">Loading...</div>';
  var s = state.summary;
  var total = Number(s.total) || 0;
  var spent = (typeof s.spent === 'number') ? Number(s.spent) : total;
  var reimbursed = Number(s.reimbursed) || 0;
  var oneoff = Number(s.oneoff) || 0;
  var oneoffCount = Number(s.oneoffCount) || 0;
  var count = Number(s.count) || 0;
  var avg = count > 0 ? spent / count : 0;

  var budget = Number(s.budget) || 0;
  var variable = spent;
  var pctSpent = budget > 0 ? Math.min(100, Math.max(0, (variable / budget) * 100)) : 0;
  var calibrated = budget > 0;
  // ---- Actuals model (locked 2026-11-14) -------------------------------
  // Every number on this screen derives from what was actually spent.
  // The budget survives only as "what remains of the allowance", so logging
  // a receipt moves the stackable figure immediately.
  var isCurrent = (state.month === todayMonth());
  var dim = daysInMonth(state.month);
  var elapsed = isCurrent ? new Date().getDate() : dim;
  if (elapsed < 1) elapsed = 1;

  var expectedPct = Math.min(100, (elapsed / dim) * 100);
  var ahead = (pctSpent <= expectedPct);

  var cfg = s.settings || {};
  var incTotal = Number(cfg.income_total) || 0;
  var fixedTotal = Number(cfg.fixed_total) || 0;
  var debtPay = Number(s.debtPayment) || 0;

  // Pace forecast: straight-line this month's RATE-BEARING spend to month
  // end, then count one-off exactly once. Projecting the raw total multiplied
  // a single repair by dim/elapsed -- 600 on day 3 of 30 became 6,000 of
  // phantom future spend and drove Stackable to a fake -5,000 until mid-month.
  // regular already nets reimbursed, so nothing else needs adding back.
  var regularSpend = (typeof s.regular === 'number') ? Number(s.regular)
                     : Math.max(0, variable - oneoff - reimbursed);
  if (regularSpend < 0) regularSpend = 0;
  var projectedRegularOnly = Math.round((regularSpend / elapsed) * dim * 100) / 100;
  var projected = Math.round((projectedRegularOnly + oneoff) * 100) / 100;
  var showProj = (elapsed >= 2) && isCurrent;

  // Stackable = income - fixed - card - spend so far - what is still to come.
  // The "still to come" term uses the budget's remaining allowance when one is
  // set, and the forecast otherwise. Both flavours stay honest: as spend rises,
  // stackable falls, on the same day you spent it.
  // Remainder MUST come from pace, never from "budget minus spent".
  // Using the budget remainder treated unspent allowance as money already
  // committed, so overspending barely moved the hero and it froze near the
  // same figure all month. Pace tracks what is actually still coming.
  var canStack = null;
  var stackBasis = '';
  var stillToCome = 0;
  if (incTotal > 0) {
    var committed = fixedTotal + debtPay + variable;
    if (showProj) {
      // Projections only: one-off is already spent, so it must not appear
      // on either side of this subtraction.
      stillToCome = Math.max(0, projectedRegularOnly - regularSpend);
      canStack = Math.round((incTotal - committed - stillToCome) * 100) / 100;
      stackBasis = 'if this pace holds';
    } else {
      canStack = Math.round((incTotal - committed) * 100) / 100;
      stackBasis = 'spent so far';
    }
  }

  // Debt is NOT a hero line any more. It is a fixed monthly outflow, so it
  // belongs beside the other outflow (spend), not beside the answer (stack).
  // Rendered as one small figure in the Spent card footer.
  var cardRent = null;
  var debtMonthlyInt = 0;
  for (var di = 0; di < (state.debts || []).length; di++) {
    var dp = state.debts[di] && state.debts[di].payoff;
    if (dp) debtMonthlyInt += Number(dp.monthlyInterest) || 0;
  }
  if (debtPay > 0) cardRent = debtPay;

  var heroAmt, heroLbl, heroNeg, heroSub;
  // Short, so it can sit BESIDE the figure instead of above it. A full
  // sentence-width label above a 42px number forces the eye to travel, and
  // the travel is where the meaning gets lost.
  if (canStack !== null) {
    heroAmt = canStack;
    heroLbl = 'Stackable';
    heroNeg = canStack < 0;
    heroSub = 'income ' + money(incTotal) + ' &minus; fixed ' + money(fixedTotal) +
               (debtPay > 0 ? ' &minus; card ' + money(debtPay) : '') +
               ' &minus; spent ' + money(variable) +
               (stillToCome > 0 ? ' &minus; rest of month ' + money(stillToCome) : '') +
               ' &middot; ' + stackBasis;
  } else {
    heroAmt = variable;
    heroLbl = 'Spent this month';
    heroNeg = false;
    heroSub = 'day ' + elapsed + ' of ' + dim +
              (showProj ? ' &middot; on pace for ' + money(projected) : '');
  }

  // Real stack rate: stackable / income, from actuals, not from the budget.
  var stackPct = (incTotal > 0 && canStack !== null) ? (canStack / incTotal) * 100 : 0;
  var stackStr = (Math.round(stackPct * 10) / 10).toFixed(1);

  // Stacked so far vs capability left. canStack is this month's capacity and is
  // NOT reduced by Sats purchases -- buying bitcoin is deploying the capacity,
  // not consuming it. "Left" is what remains un-deployed, so the identity holds:
  //   income = fixed + card + projected spend + stacked + left
  // Hero = capacity; this line reconciles it with what has been bought.
  var stackedAmt = Number(s.stacked) || 0;
  var stackedLeft = (canStack !== null) ? Math.max(0, canStack - stackedAmt) : 0;

  var h = '';
  if (!calibrated) h += calibOfferHtml();
  h += '<div class="card hero-card">';

  // Headline: the label rides BESIDE the figure, not above it. One eye-line,
  // one block, instead of a whisper, a shout and three lines of fine print.
  // Both modes must read identically: label centred above, figure centred
  // below. Side-by-side worked for sats (short) and failed for fiat, where a
  // long "-BND 241.59" slid off-centre and left the label floating in space.
  // Premium hero: the label is the brand. STACKABLE is set large above the
  // figure. The fallback (no income configured) keeps the whisper label so a
  // sentence never shouts at 23px.
  h += '<div class="hero-top"><span class="hero-lbl' +
       (canStack !== null ? ' brand' : '') + '">' + heroLbl + '</span></div>';
  h += '<div class="hero-amt' + (heroNeg ? ' neg' : '') + '" ' +
       'data-act="denom" role="button" tabindex="0" aria-label="Switch between fiat and sats">' +
       amt(heroAmt) + '</div>';

  // The other denomination, once. This is the line that makes the number
  // mean something: BND alone does not answer "how much bitcoin".
  h += heroAltHtml(canStack);
  h += stackedLineHtml(stackedAmt, stackedLeft, canStack);
  h += heroStackLineHtml(canStack, incTotal, stackStr);

  // The pace bar lives on the Spent card now: it measures budget consumption,
  // which is a fact about spending, not about what is stackable. Keeping it
  // in the hero made one number carry two unrelated questions.
  if (!calibrated) {
    h += calibPanelHtml(elapsed, dim);
  }

  // The arithmetic lives behind a disclosure so the hero stays one number.
  // Open state is remembered in state.heroOpen, not persisted across reloads.
  h += '<button class="hdisc' + (state.heroOpen ? ' open' : '') + '" type="button" ' +
       'data-act="heroDisc" aria-label="Show how this is calculated" ' +
       'aria-expanded="' + (state.heroOpen ? 'true' : 'false') + '">' +
       '<i class="hdisc-i"></i></button>';
  if (state.heroOpen) {
    h += '<div class="hcalc">' + heroSub + '</div>';
  }
  h += '</div>';

  // One headline figure: everything that left the account this month.
  // One-off is real spending, so it belongs in the total, not beside it.
  var totalSpentThisMonth = Math.round(variable * 100) / 100;
  var spentSub = [];
  if (reimbursed < 0) {
    spentSub.push(money(Math.abs(reimbursed)) + ' reimbursed back');
  }

  // One card, one figure, then a quiet two-up footer. The old 2x2 grid put
  // four cells in a box and gave "Nothing this month" the same visual weight
  // as the spend total; a footer admits it is secondary.
  // Centred, same skeleton as the hero: whisper label, shout figure. The hero
  // answers "can I stack" and this answers "what did I spend" -- two headline
  // numbers cannot sit on two different alignments without one looking like a
  // mistake. Everything below the figure stays left, where lists belong.
  h += '<div class="card">';
  h += '<div class="spent-top"><span class="lbl">Spent this month</span></div>';
  h += '<div class="spent-amt" data-act="denom" role="button" tabindex="0" ' +
       'aria-label="Switch between fiat and sats">' + amt(totalSpentThisMonth) + '</div>';
  if (spentSub.length) h += '<div class="spent-sub center">' + spentSub.join(' &middot; ') + '</div>';

  // Pace lives here because pace is a statement about SPENDING. In the hero it
  // was answering a question the hero had not asked, and it forced the reader
  // to hold two budgets in their head at once.
  if (calibrated && showProj) {
    var overUnder = Math.round((budget - projected) * 100) / 100;
    // The headline is what was spent; the bar's scale is the budget. Without
    // a caption the two figures look like they disagree.
    h += '<div class="pace-cap">Budget used &middot; incl. one-off</div>';
    h += '<div class="pace"><i class="fill' + (pctSpent >= 100 ? ' over' : '') +
         '" style="width:' + pctSpent.toFixed(1) + '%"></i>' +
         '<span class="tick" style="left:' + expectedPct.toFixed(1) + '%"></span></div>';
    h += '<div class="pace-scale"><span>0</span><span>' + money(budget) + '</span></div>';
    h += '<div class="pacelbl ' + (ahead ? 'ok' : 'behind') + '">' +
         (ahead
            ? money(overUnder) + ' under pace'
            : money(Math.abs(overUnder)) + ' over pace') +
         ' &middot; day ' + elapsed + ' of ' + dim + '</div>';
  }

  // A one-off is the one number in this card a person can act on, so it gets
  // its own line instead of being buried in a grey sentence with three other
  // facts. Receipts moved to By category (that card is about the entries) and
  // stack rate moved up to the hero (that card is about stacking).
  // Same geometry as the Card payment row: label left, figure right in
  // near-black, detail on its own grey line under the figure. Two money rows
  // sitting one above the other have to share a grid or they read as a mess.
  if (oneoff > 0) {
    var ooOpen = !!state.oneoffOpen;
    h += '<div class="oneoff' + (ooOpen ? ' open' : '') + '" data-act="oneoffDisc" ' +
         'role="button" tabindex="0" aria-expanded="' + (ooOpen ? 'true' : 'false') + '">' +
         '<div class="dr-top"><span class="dr-k">One-off</span>' +
         '<span class="dr-v">' + money(oneoff) + '</span>' +
         '<i class="oo-i"></i></div>' +
         '<div class="dr-sub">' + oneoffCount + ' entr' +
         (oneoffCount === 1 ? 'y' : 'ies') + '</div>';
    if (ooOpen) {
      var ooRows = oneoffRowsHtml();
      h += ooRows || '<div class="oo-empty">No one-off entries this month</div>';
    }
    h += '</div>';
  }
  h += debtRentHtml(cardRent, debtMonthlyInt);
  h += '</div>';

  // Receipts belongs here: this card is about the entries, so the count and the
  // average cost are the header for the thing being ranked, not a stray stat.
  h += '<div class="card">';
  h += '<div class="cathead"><span class="lbl">By category</span>' +
       (count > 0
         ? '<span class="rcpt"><b>' + count + '</b> entr' + (count === 1 ? 'y' : 'ies') +
           ' &middot; avg ' + money(avg) + '</span>'
         : '') + '</div>';
  h += catsHtml(s.byCategory || {});
  h += '</div>';
  return h;
}

// Deploy 2: what the debts actually cost. Pure arithmetic on what the user
// entered - no advice, no recommendations.
// Calibration panel: the bar measures CONFIDENCE (days of data), not money.
// It gives a finish line so the mode does not feel like a dead end.
function calibPanelHtml(elapsed, dim){
  var left = Math.max(0, dim - elapsed);
  var pct = dim > 0 ? Math.min(100, (elapsed / dim) * 100) : 0;
  var h = '<div class="calib">';
  h += '<div class="calib-t">' + (left > 0
    ? 'Keep logging for ' + left + ' more day' + (left === 1 ? '' : 's') +
      ' and I&rsquo;ll set your budget from what you actually spend.'
    : 'Month complete. Tap below to set your budget from real spending.') + '</div>';
  h += '<div class="calib-bar"><i style="width:' + pct.toFixed(1) + '%"></i></div>';
  h += '<div class="calib-d">' + elapsed + ' of ' + dim + ' days</div>';
  h += '</div>';
  return h;
}

// Month-end offer: once enough history exists, the app proposes the measured
// budget and accepts it in one tap. This is the payoff moment for calibration.
function calibOfferHtml(){
  var sg = state.suggestion;
  if (!sg || !sg.ok || !sg.ready) return '';
  if (state.calibDismissed) return '';
  var h = '<div class="calib offer">';
  h += '<div class="calib-t">You averaged <b>' + money(sg.average) + '</b> over ' +
       sg.months + ' month' + (sg.months === 1 ? '' : 's') + '. Use it as your budget?</div>';
  h += '<div class="calib-btns">';
  h += '<button type="button" class="calib-yes" data-act="adoptBudget" data-amt="' +
       sg.average + '">Set ' + money(sg.average) + '</button>';
  h += '<button type="button" class="calib-no" data-act="dismissCalib">Not now</button>';
  h += '</div>';
  h += '</div>';
  return h;
}

// What clearing this debt actually unlocks, in sats. The prize, not the tax.
// One line under the hero: what the debt costs in sats, and the way in.
// Returns '' when there is no debt, so the card looks untouched.
function leakCardHtml(){
  var list = state.debts || [];
  var live = [];
  for (var i = 0; i < list.length; i++) {
    if (list[i] && list[i].payoff) live.push(list[i]);
  }

  var price = (state.btc && state.btc.ok && state.btc.price_bnd) ? Number(state.btc.price_bnd) : 0;

  if (!live.length) {
    var freeAmt = state.summary && state.summary.settings
      ? (Number(state.summary.settings.income_total) || 0) -
        (Number(state.summary.settings.fixed_total) || 0) -
        (Number(state.summary.settings.variable_budget) || 0)
      : 0;
    freeAmt = Math.round(freeAmt * 100) / 100;
    var g = '<div class="card leak free">';
    g += '<div class="lbl">No debt</div>';
    g += '<div class="leak-big">' + money(freeAmt) + ' <span>/month</span></div>';
    g += '<div class="leak-sub">Nothing is eating it. Every bit of that is yours to stack.</div>';
    g += satsHtml(freeAmt);
    g += '</div>';
    return g;
  }

  // THE PAGE THESIS SENTENCE, stated once. This card owns every DECISION
  // number -- what the debt costs, what clearing frees, what interest lies
  // ahead. The debt cards below are pure FACTS. Before this split the same
  // four numbers appeared on both cards, so the page read like it stuttered.
  var monthlyIntTotal = 0, payTotal = 0, intAheadTotal = 0;
  for (var j = 0; j < live.length; j++) {
    var pj = live[j].payoff || {};
    monthlyIntTotal += Number(pj.monthlyInterest) || 0;
    intAheadTotal += Number(pj.totalInterest) || 0;
    payTotal += Number(live[j].payment) || 0;
  }
  monthlyIntTotal = Math.round(monthlyIntTotal * 100) / 100;
  intAheadTotal = Math.round(intAheadTotal * 100) / 100;
  payTotal = Math.round(payTotal * 100) / 100;

  var h = '<div class="card leak">';
  h += '<div class="lbl">Your debt costs you</div>';

  if (price > 0 && monthlyIntTotal > 0) {
    var miSats = Math.round((monthlyIntTotal / price) * 100000000);
    // Red lives HERE and only here: this is the alarm. Grey elsewhere.
    h += '<div class="leak-big cost">' + miSats.toLocaleString() +
         ' <span>sats / month</span></div>';
    h += '<div class="leak-sub">' + money(monthlyIntTotal) + '/mo in interest' +
         (state.btc.stale ? ' &middot; price ' + priceAgeLabel(state.btc.fetched_at) : '') +
         '</div>';
  } else if (price > 0) {
    h += '<div class="leak-big">0 <span>sats / month</span></div>';
    h += '<div class="leak-sub">No interest accruing.</div>';
  } else {
    h += '<div class="leak-big">' + money(monthlyIntTotal) + ' <span>/month</span></div>';
    h += '<div class="leak-sub">price feed unavailable, so no sats shown</div>';
  }

  // The two facts that make the number actionable: what it buys back, and
  // what it costs if left alone. One each, never repeated below.
  h += '<div class="leak-fix">';
  h += '<div class="leak-rule"></div>';
  if (price > 0 && payTotal > 0) {
    var freeSats = Math.round((payTotal / price) * 100000000);
    h += '<div class="leak-fix row"><span>Clearing everything frees</span>' +
         '<b class="green">' + freeSats.toLocaleString() + ' sats/mo</b></div>';
  }
  if (price > 0 && intAheadTotal > 0) {
    var aheadSats = Math.round((intAheadTotal / price) * 100000000);
    h += '<div class="leak-rule"></div>';
    h += '<div class="leak-fix row"><span>Interest ahead if left alone</span>' +
         '<b class="leak-fix-red">' + aheadSats.toLocaleString() + ' sats</b></div>';
  }
  var anyRevolving = false;
  for (var v = 0; v < live.length; v++) {
    if (String(live[v].type || '') !== 'fixed') anyRevolving = true;
  }
  // Inside the block so it shares the block's rhythm and the card ends tight.
  // It has its own rule above, so it still reads as fine print rather than a
  // fourth fact.
  if (anyRevolving) {
    h += '<div class="leak-rule"></div>';
    h += '<div class="leak-fix-d">Assumes you stop adding to the card.</div>';
  }
  h += '</div>';

  // Ranked bleed, only when ranking means anything. With one debt it would
  // restate the headline in lighter type.
  if (live.length > 1 && price > 0) {
    var ranked = live.slice().sort(function(a, b){
      return (Number(b.payoff.monthlyInterest) || 0) - (Number(a.payoff.monthlyInterest) || 0);
    });
    h += '<div class="leak-list">';
    for (var k = 0; k < ranked.length; k++) {
      var d = ranked[k];
      var dSats = Math.round(((Number(d.payoff.monthlyInterest) || 0) / price) * 100000000);
      h += '<div class="leak-row">';
      h += '<span class="leak-n">' + (k === 0 ? '<i class="leak-first">1</i>' :
           '<i class="leak-idx">' + (k + 1) + '</i>') + esc(d.name) + '</span>';
      h += '<span class="leak-v">' + dSats.toLocaleString() + '</span>';
      h += '</div>';
    }
    h += '</div>';
  }

  h += '</div>';
  return h;
}

// Stackable as shown in the hero, recomputed for the before/after maths.
function currentStackable(){
  var cfg = (state.summary && state.summary.settings) || {};
  var inc = Number(cfg.income_total) || 0;
  if (inc <= 0) return 0;
  var a = inc - (Number(cfg.fixed_total) || 0) -
          (Number(cfg.variable_budget) || 0);
  a -= Number(state.summary && state.summary.debtPayment) || 0;
  return Math.round(a * 100) / 100;
}

/* ---------------- inline payment editor ---------------- */

// Opened from the pencil on a debt card. One number, one save. Deliberately not
// a copy of the Settings debt form - that form owns all fields at once, this
// owns exactly one and posts the rest back unchanged.
function payEditorHtml(){
  var did = state.payEdit;
  if (!did) return '';
  var list = state.debts || [];
  var d = null;
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].id) === String(did)) { d = list[i]; break; }
  }
  if (!d) return '';
  var h = '<div class="mask" data-act="mask"><div class="sheet">';
  h += '<h2 class="ehead">' + esc(d.name || 'Debt') + '</h2>';
  h += '<div class="hint">What you actually pay each month. The stackable figure moves with this.</div>';
  h += '<form data-act="savepay" autocomplete="off">';
  h += '<label>Monthly payment</label>';
  h += '<input type="number" step="0.01" min="0" inputmode="decimal" id="pay_amount" ' +
       'value="' + (!Number(d.payment) ? '' : Number(d.payment).toFixed(2)) + '" ' +
       'placeholder="0.00" onfocus="this.select()" required>';
  if (state.error) h += '<div class="err">' + esc(state.error) + '</div>';
  h += '<button class="btn" type="submit"' + (state.busy ? ' disabled' : '') + '>' +
       (state.busy ? 'Saving...' : 'Save payment') + '</button>';
  h += '<button class="btn alt" data-act="closePay" type="button">Cancel</button>';
  h += '</form></div></div>';
  return h;
}

function savePayFromForm(){
  var did = state.payEdit;
  var list = (state.debts || []).slice();
  var idx = -1;
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].id) === String(did)) { idx = i; break; }
  }
  if (idx < 0) return;
  var raw = valOf("pay_amount");
  var n = Number(raw);
  if (!isFinite(n) || n < 0) {
    state.error = 'Enter a monthly payment amount';
    render();
    return;
  }
  // Copy the row before mutating: list is a shallow slice, so writing through
  // it would edit the object still held by state.debts.
  list[idx] = (function(o){ var c={}; for (var k in o) c[k]=o[k]; c.payment = round2c(n); return c; })(list[idx]);

  // Close the sheet FIRST. Rendering while the form is still mounted tears the
  // input out from under the click, and the callback then updates a detached
  // node - the button looks dead. Capture the value, dismiss, then save.
  state.payEdit = null;
  state.error = null;
  state.debts = list;
  render();

  run("saveDebts", { debts: list }, function(res){
    if (!res || res.ok !== true) {
      state.error = 'Could not save: ' + ((res && res.error) ? res.error : 'unknown');
      render();
      return;
    }
    state.debts = res.debts || list;
    loadSummary();   // debtPayment feeds canStack, so the hero must re-read
    loadDebts();
    toast('Payment saved');
  });
}

function debtTabHtml(){
  var list = state.debts || [];
  // Only rows the server enriched with a payoff are real debts. A settings
  // draft row slipping in here used to render an empty tab under a populated
  // leak card - the "No debt" screen over BND 388/month.
  var real = [];
  for (var rq = 0; rq < list.length; rq++) {
    if (list[rq] && list[rq].payoff) real.push(list[rq]);
  }
  list = real;
  if (!list.length) {
    // A new user lands here with a messy credit situation and no idea what to
    // pay. One screen that explains the value, then the form. No dead end.
    var st = '<div class="card">';
    st += '<div class="starter">';
    st += '<div class="starter-t">See what your debt costs you.</div>';
    st += '<div class="starter-p">Add a card with its balance and rate. That is all ' +
          'we need to show what it costs you in sats every month, and what ' +
          'clearing it would give back.</div>';
    st += '<div class="starter-p" style="margin-top:9px">Not sure what to pay? ' +
          'Start with the minimum from your statement. You can change it later ' +
          'and watch the sats move.</div>';
    st += '<button class="starter-b" data-act="addFirstDebt" type="button">Add my card</button>';
    st += '</div></div>';
    return st;
  }

  var h = '';
  h += leakCardHtml();

  var live = [];
  for (var i = 0; i < list.length; i++) {
    if (list[i] && list[i].payoff) live.push(list[i]);
  }
  live.sort(function(a, b){
    return (Number(b.rate) || 0) - (Number(a.rate) || 0);
  });

  var price = (state.btc && state.btc.ok && state.btc.price_bnd) ? Number(state.btc.price_bnd) : 0;

  for (var k = 0; k < live.length; k++) {
    var d = live[k];
    var p = d.payoff || {};
    var bal = Number(d.balance) || 0;
    var monthlyInt = Number(p.monthlyInterest) || 0;

    // FACTS ONLY on this card. The decision numbers (monthly interest total,
    // what clearing frees, interest ahead) live once in the aggregate card
    // above. Repeating them here is what made the page stutter.
    // Same skeleton as the Summary cards: name as a whisper at the top, one
    // big centred figure, then quiet left-aligned facts. The old version put
    // the card name and the fiat balance on one line, which is why "Baiduri
    // Smart Executive Visa Platinum" wrapped to two lines and shoved "BND"
    // and "1,374.20" onto separate lines beside it.
    h += '<div class="card">';
    h += '<div class="dc-top"><span class="dc-n">' + esc(d.name) + '</span></div>';

    // Sats first: the fiat balance is the bank's unit, not the reader's.
    if (price > 0) {
      var balSats = Math.round((bal / price) * 100000000);
      h += '<div class="dc-bal">' + balSats.toLocaleString() + '</div>';
      h += '<div class="dc-bal-sub">sats owed &middot; ' + money(bal) + '</div>';
    } else {
      h += '<div class="dc-bal">' + money(bal) + '</div>';
      h += '<div class="dc-bal-sub">price feed unavailable, so no sats shown</div>';
    }

    if (p.stuck) {
      h += '<div class="dc-line warn">' + esc(p.stuckReason || 'Never clears at this payment') + '</div>';
    } else {
      // Payment is editable here, not only in Settings. One field, one save:
      // it posts the whole debt list through saveDebts, then reloads so the
      // stackable figure moves with it.
      h += '<div class="dc-facts"><span class="dc-fk">Paying</span>' +
           '<span class="dc-fv">' + money(Number(d.payment) || 0) + '/mo' +
           '<button type="button" class="dcpay-edit" data-act="editPay" data-did="' + esc(d.id) + '"' +
           ' aria-label="Change monthly payment">&#9998;</button></span></div>';
      var payAmt = Number(d.payment) || 0;
      var intShare = payAmt > 0 ? (monthlyInt / payAmt) : 0;
      if (intShare >= 0.5) {
        // More than half of every payment is interest: the date is a fiction
        // worth hiding, and the user needs to know why.
        var pctInt = Math.round(intShare * 100);
        h += '<div class="dc-line warn">' + pctInt + '% of this payment is interest. ' +
             'At this rate the debt barely moves &mdash; raise the payment or it never clears.</div>';
      } else {
        // The clear date is a per-card fact; the interest behind it is not
        // restated here because the aggregate card already owns that number.
        h += '<div class="dc-facts"><span class="dc-fk">Clears</span>' +
             '<span class="dc-fv">' + esc(p.clearsOn || '') + ' &middot; ' + p.months +
             ' month' + (p.months === 1 ? '' : 's') + '</span></div>';
      }
    }

    // One upgrade line per card, not two. suggestionHtml() below prices the
    // same 1.25x bump, so rendering both said the same thing twice.
    // Sub-2-month gains are dropped: on a small balance a 1-month gain reads
    // as noise, not a decision.

    h += suggestionHtml(d, price);
    h += payPresetsHtml(d, price);

    h += '</div>';
  }

  // No trailing card: a whole white card holding one grey sentence was pure
  // weight. The edit hint rides the aggregate card as a quiet footer.
  return h;
}

// Three payment options, priced on the spot. The backend already knew how to
// price a larger payment; this is the surface for it, so "what if I paid
// more" stops being a number you have to guess at.
// One line that prices the next step up. Reading 2 (price the choice) doing
// the work of Reading 1 (tell them what to pay) without becoming a calculator.
function suggestionHtml(d, price){
  var pay = Number(d.payment) || 0;
  if (pay <= 0) return '';
  var base = d.payoff || {};
  if (base.months === null || base.stuck) return '';
  var up = Math.round(pay * 1.25 * 100) / 100;
  if (Math.abs(up - pay) < 0.01) return '';
  var sim = payoffAt(d, up);
  var baseMonths = Math.max(1, Number(base.months) || 0);
  // Worth showing when the gain is a real share of the wait, not a count of
  // months. A flat "2 months" hid the biggest decision on a short card: on a
  // 5-month payoff, 1 month saved is a fifth of the problem. On a 30-month
  // payoff it is noise. Same proportional rule the preset chips already use.
  if (!sim || sim.monthsSaved <= 0) return '';
  if (baseMonths > 0 && (sim.monthsSaved / baseMonths) < 0.15) return '';
  // One sentence, one action. The chips below are the control; this line is the
  // same trade stated once so the two figures agree instead of repeating.
  var pct = Math.round(((up - pay) / pay) * 100);
  var h = '<div class="dc-line dcsugg">';
  h += 'Paying ' + pct + '% more (' + money(up) + '/mo) clears this ' +
       sim.monthsSaved + ' month' + (sim.monthsSaved === 1 ? '' : 's') + ' earlier';
  if (price > 0 && sim.interestSaved > 0) {
    var ss = Math.round((sim.interestSaved / price) * 100000000);
    h += ' and keeps ' + ss.toLocaleString() + ' sats';
  } else if (sim.interestSaved > 0) {
    h += ' and saves ' + money(sim.interestSaved);
  }
  h += '.</div>';
  return h;
}

// Presets that actually buy something. Multiplying the current payment by
// 1.25 and 1.5 is meaningless at a low balance: on BND 1,374 at 18%, paying
// 300 vs 450 both clear in the same month, so all three buttons answered
// "try a bigger payment" and the section looked broken. Instead, walk the
// payment up until it buys one / two / three whole months, and show those.
function payPresetsFor(d){
  var pay = Number(d.payment) || 0;
  if (pay <= 0) return [];
  var base = payoffAtBase(d);
  // A stuck card (payment below monthly interest) returns null. Do not bail
  // here: the whole point of the presets is showing what WOULD clear it.
  // monthsSaved needs a baseline, so fall back to the last month in the guard.
  var stuckBase = !base;
  if (stuckBase) base = { months: 101, monthsSaved: 0 };
  // Fractions of the remaining term, not absolute months. On a 31-month payoff
  // "1 month sooner" is invisible; on a 5-month payoff "3 months sooner" is
  // nearly the whole game. Scaling to the term keeps the rungs meaningful on
  // every balance size.
  var baseMonths = Math.max(2, Number(base.months) || 1);
  var want = [Math.max(1, Math.round(baseMonths * 0.15)),
              Math.max(2, Math.round(baseMonths * 0.30)),
              Math.max(3, Math.round(baseMonths * 0.45))];
  // Start the list EMPTY. The current payment is already stated on the
  // "Paying" line directly above, so offering it as a chip spends a slot to
  // say nothing. Every chip here must buy a different outcome.
  var out = [];
  var cur = pay;
  var wi = 0;
  var guard = 0;
  // Ceiling is generous: a stuck card needs a payment well above the balance
  // before it amortises at all, so capping at 6x the minimum hid every rung.
  var ceiling = Math.max(pay * 6, (Number(d.balance) || 0) * 0.25);
  while (wi < want.length && guard < 80) {
    guard++;
    cur = Math.round(cur * 1.05 * 100) / 100;  // 5% steps: cheap, lands close
    if (cur > ceiling) break;                  // never offer absurd jumps
    var sim = payoffAt(d, cur);
    if (!sim) continue;                        // this rung cannot clear at all
    // On a stuck card the baseline is unknown, so judge by absolute term.
    if (stuckBase) {
      sim.monthsSaved = Math.max(0, 101 - sim.months);
    }
    if (sim.monthsSaved >= want[wi]) {
      out.push(cur);
      wi++;
    }
  }
  // Nothing reachable (tiny balance, payoff already short): offer the next
  // honest increment anyway rather than an empty row.
  if (!out.length) {
    var up = Math.round(pay * 1.5 * 100) / 100;
    if (payoffAt(d, up)) out.push(up);
  }
  // Three chips max: four overflowed a 375px card.
  return out.slice(0, 3);
}

// Baseline months (current payment) used for monthsSaved comparisons.
function payoffAtBase(d){
  return payoffAt(d, Number(d.payment) || 0);
}

function payPresetsHtml(d, price){
  var pay = Number(d.payment) || 0;
  if (pay <= 0) return '';
  var amts = payPresetsFor(d);
  if (!amts.length) return '';
  var chosen = Number(state.simPay[d.id]) || pay;
  if (amts.indexOf(chosen) < 0) chosen = pay;
  // Baseline "clearing month" at the payment they already make. A rung only
  // earns a month tag if it lands in an EARLIER calendar month than this.
  var baseRung = payoffAt(d, pay);
  var h = '<div class="payset">';
  h += '<div class="payset-h">If I paid</div>';
  h += '<div class="payset-btns">';
  for (var i = 0; i < amts.length; i++) {
    var on = Math.abs(chosen - amts[i]) < 0.005;
    var rung = payoffAt(d, amts[i]);
    // The tag must measure the SAME thing the line below states. monthsSaved
    // counts amortisation months, but the line reads "Same clearing month"
    // when the calendar month does not move - so a card at 18% showed 1mo/3mo/4mo
    // above a sentence saying nothing changed. Judge by calendar month.
    var tag = (rung && clearsDifferentMonth(rung, baseRung)) ?
              '<em>' + rung.monthsSaved + 'mo</em>' : '';
    h += '<button class="payset-btn' + (on ? ' on' : '') + '" type="button" ' +
         'data-act="simPay" data-did="' + esc(d.id) + '" data-amt="' + amts[i] + '">' +
         money(amts[i]) + tag + '</button>';
  }
  h += '</div>';
  var sim = payoffAt(d, chosen);
  // The line under the chips describes the SELECTED rung. When nothing is
  // selected it would restate the current payment, which the "Paying" fact
  // above and the green box already cover - so say nothing.
  if (sim && sim.monthsSaved > 0 && Math.abs(chosen - pay) >= 0.005) {
    var out = '<b>' + sim.monthsSaved + ' month' + (sim.monthsSaved === 1 ? '' : 's') +
              '</b> sooner';
    if (price > 0 && sim.interestSaved > 0) {
      var ss = Math.round((sim.interestSaved / price) * 100000000);
      out += ' &middot; saves <span class="green">' + ss.toLocaleString() + ' sats</span>';
    } else if (sim.interestSaved > 0) {
      out += ' &middot; saves <b>' + money(sim.interestSaved) + '</b>';
    }
    out += ' &middot; clears ' + esc(sim.clearsOn);
    h += '<div class="payset-out">' + out + '</div>';
  } else if (sim && Math.abs(chosen - pay) >= 0.005) {
    // Never render this for the current payment: before any chip is tapped,
    // chosen === pay, and this branch printed "Same clearing month" under the
    // chips on first load, as if paying the same amount were an answer.
    // At small balances an extra 25% may buy no whole month, only interest.
    // Never tell the user their money does nothing: show what it does save.
    var outp = 'Same clearing month';
    if (sim.interestSaved > 0) {
      if (price > 0) {
        var s2 = Math.round((sim.interestSaved / price) * 100000000);
        outp += ' &middot; still saves <span class="green">' + s2.toLocaleString() + ' sats</span>';
      } else {
        outp += ' &middot; still saves <b>' + money(sim.interestSaved) + '</b>';
      }
    }
    h += '<div class="payset-out">' + outp + '</div>';
  }
  h += '</div>';
  return h;
}

// True only when a rung actually clears in an earlier calendar month than the
// current payment. Term length and calendar month are different claims: a card
// at 18% can shave a month of term and still clear in the same month.
function clearsDifferentMonth(rung, base){
  if (!rung || !rung.clearsOn) return false;
  if (!base || !base.clearsOn) return true;  // no baseline: trust the rung
  return rung.clearsOn !== base.clearsOn;
}

// Amortisation for the what-if buttons. Same maths as the server, run here so
// the number moves the moment a preset is tapped. Returns null when the payment
// can never clear the balance.
function payoffAt(d, pay){
  var bal = Number(d.balance) || 0;
  var rate = Number(d.rate) || 0;
  var p = Number(pay) || 0;
  if (bal <= 0 || p <= 0) return null;
  var mr = rate / 100 / 12;
  if (p <= bal * mr) return null;
  var remaining = bal, months = 0, totalInterest = 0, guard = 1200;
  while (remaining > 0.005 && months < guard) {
    var interest = remaining * mr;
    totalInterest += interest;
    remaining -= (p - interest);
    months++;
  }
  if (remaining > 0.005) return null;
  var base = d.payoff || {};
  // A stuck baseline (payment never covers interest) has months = null. The
  // old Number(null) || months collapsed to this rung's own months, so
  // monthsSaved came out 0 and the buttons showed '0mo' next to a card that
  // never cleared at all. Flag it instead of pretending nothing improved.
  var baseStuck = !!(base.stuck || base.months === null || base.months === undefined);
  var baseMonths = baseStuck ? months : (Number(base.months) || months);
  var baseInterest = baseStuck ? 0 : (Number(base.totalInterest) || 0);
  // Day-1 anchor, same reason as payoffOf(): setMonth() overflows on the
  // 29th-31st and reports the wrong clearing month.
  var nowD = new Date();
  var dd = new Date(nowD.getFullYear(), nowD.getMonth() + months, 1);
  var MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return {
    months: months,
    monthsSaved: Math.max(0, baseMonths - months),
    interestSaved: Math.max(0, baseInterest - totalInterest),
    clearsOn: MON[dd.getMonth()] + ' ' + dd.getFullYear(),
    baseStuck: baseStuck
  };
}

// Measured history only: what the app itself observed, never a backfill.
// Same cell, but the value itself is a fiat/sats switch.
// C1: read-only sats equivalent of the BTC-able amount. Hidden entirely when
// the price feed is unavailable - a stale or invented price is worse than none.
// Price age, in words. Used everywhere a sats figure is shown, so a stale
// price is always visibly stale rather than silently drifting.
function priceAgeLabel(ts){
  if (!ts) return '';
  var mins = Math.round((Date.now() - ts) / 60000);
  if (mins < 1) return 'just now';
  if (mins < 60) return mins + 'm ago';
  var hrs = Math.round(mins / 60);
  if (hrs < 24) return hrs + 'h ago';
  return Math.round(hrs / 24) + 'd ago';
}

// How old the BTC price is, on the home page, in both denominations.
// The sats line inside satsHtml() also carries an age, but that line is hidden
// in sats mode -- so switching to sats silently removed the only staleness
// signal on the page. One stamp, always visible, no state dependence.
// The other denomination, once, under the headline. "BND 241.59" alone does
// not answer how much bitcoin; "312,400 sats" alone is unverifiable. Both.
function heroAltHtml(bndVal){
  var b = state.btc;
  if (!b || !b.ok || !b.price_bnd) return '';
  var v = Number(bndVal) || 0;
  if (state.denom === 'sats') {
    return '<div class="hero-alt">' + money(v) + '</div>';
  }
  var sats = satsOf(v);
  if (sats === null) return '';
  var stale = b.stale ? ' &middot; price ' + priceAgeLabel(b.fetched_at) + ' old' : '';
  return '<div class="hero-alt">' + satsText(sats) + stale + '</div>';
}

// The rows behind the One-off figure, rendered inside the dropdown. Kept
// deliberately plain: date, vendor, amount -- the same facts as the Expenses
// list, minus the swipe actions, because this is a look, not an edit surface.
function ooDay(iso){
  // '2026-11-03' -> '3 Nov'. Built with string slicing, NOT a regex, on purpose.
  // A backslash-escaped regex inside the APP_HTML template literal was being
  // evaluated away: what the browser received was /^(d{4})-(d{2})-(d{2})$/, so
  // it never matched a date and every row quietly fell back to raw ISO. Plain
  // slicing has no escapes to lose.
  var s = String(iso || '').trim();
  if (s.length < 10 || s.charAt(4) !== '-' || s.charAt(7) !== '-') return s;
  var MO = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var mi = parseInt(s.slice(5, 7), 10) - 1;
  var di = parseInt(s.slice(8, 10), 10);
  if (!(mi >= 0 && mi < 12) || !(di > 0 && di < 32)) return s;
  return di + ' ' + MO[mi];
}

function oneoffRowsHtml(){
  var rows = (state.summary && state.summary.oneoffRows) || [];
  if (!rows.length) return '';
  var h = '<div class="oo-list">';
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    h += '<div class="oo-row"><span class="oo-d">' + esc(ooDay(r.date)) + '</span>' +
         '<span class="oo-v">' + esc(r.vendor) + '</span>' +
         '<span class="oo-a">' + money(Number(r.amount) || 0) + '</span></div>';
  }
  h += '</div>';
  return h;
}

// What has been bought this month, next to what is still available. Without
// this line the hero would look identical whether you had stacked nothing or
// half of it, because Sats never enters spent. Kept deliberately terse: two
// figures and a middle dot, same weight as the stack-rate line below it.
// Sats mode shows the same two numbers in the chosen denomination, because the
// hero figure itself already respects the denom switch.
function stackedLineHtml(stacked, left, canStack){
  if (!(stacked > 0)) return '';
  var h = '<div class="hero-stacked">';
  // Branch on CAN CAPACITY BE KNOWN, not on whether left is positive. Left is
  // floored at 0, so testing left > 0 would send a fully-deployed (or
  // over-stacked) month down the "no income configured" path and drop the
  // capacity framing entirely. null is the only "unknown" case.
  if (canStack === null) {
    // No income set: there is no capacity to compare against, so report the
    // purchase and stop. Inventing a "left" here would be a guess.
    h += '<b>' + amt(stacked) + '</b> stacked so far';
  } else if (left > 0) {
    h += '<b>' + amt(stacked) + '</b> stacked &middot; ' + amt(left) + ' left to stack';
  } else {
    // Capacity fully deployed or exceeded. Say so plainly instead of printing
    // "0 left", which reads like a broken figure.
    h += '<b>' + amt(stacked) + '</b> stacked &middot; nothing left to stack';
  }
  return h + '</div>';
}

// Rate line under the hero figure. It used to sit in the Spent card's footer,
// but it is a fact about the ANSWER, not about spending. Only rendered when
// there is an income to divide by.
//
// WORDING: this line measures CAPACITY (canStack / income), never the amount
// actually bought. Calling it "of income stacked" collided with the execution
// line below ("X stacked"), which reports a different number -- 1,000 vs 500 in
// the same hero. One word for two numbers is how a user stops trusting both.
// "Stackable" is the adjective; "stacked" stays reserved for what was bought.
// The numerator also has to stay canStack: dividing the actual purchase would
// read 0.0% until DCA day and stop reacting to daily spending.
function heroStackLineHtml(canStack, incTotal, stackStr){
  if (canStack === null || !(incTotal > 0)) return '';
  // A negative stacking rate is not a rate; it is the absence of one. Say WHY:
  // "nothing yet" read as "the app has no data", which is the opposite of the
  // truth -- this month's spending genuinely outran the income.
  if (canStack < 0) {
    return '<div class="hero-rate none">Stack rate &middot; overspent this month</div>';
  }
  return '<div class="hero-rate"><b>' + stackStr + '%</b> of income stackable</div>';
}

// Card payment as part of the monthly outflow, not a decorative hero line.
// "You pay BND 300.00/mo on the card" is a fact about money leaving the
// account -- the same family as spend -- so it sits in the Spent card.
// Tappable through to Debts, quietly, with the bank's interest in the tail.
function debtRentHtml(pay, interest){
  if (!(Number(pay) > 0)) return '';
  var price = (state.btc && state.btc.ok && state.btc.price_bnd) ? Number(state.btc.price_bnd) : 0;
  var satsTail = '';
  if (price > 0 && interest > 0) {
    satsTail = ' &middot; ' +
      Math.round((Number(interest) / price) * 100000000).toLocaleString() + ' sats interest';
  }
  // Layout: label and figure on one baseline with the chevron pinned right, and
  // the sats interest demoted to its OWN line under the figure. Inline, it
  // wrapped to "27,101 sats / interest" and dragged the label out of alignment.
  // The whole row is the tap target (no inner spans stop the hit-test).
  return '<div class="debtrow" data-act="tab" data-tab="debts" role="button" tabindex="0">' +
         '<div class="dr-top"><span class="dr-k">Card payment</span>' +
         '<span class="dr-v">' + money(pay) + '/mo</span><i>&rsaquo;</i></div>' +
         (satsTail ? '<div class="dr-sub">' + satsTail.replace(/^ &middot; /, '') + '</div>' : '') +
         '</div>';
}

function priceStampHtml(){
  var b = state.btc;
  if (!b || !b.ok || !b.price_usd) return '';
  var age = priceAgeLabel(b.fetched_at);
  if (!age) return '';
  var txt = 'BTC ' + usdLabel(b.price_usd) + ' &middot; ' + age;
  if (b.stale) txt += ' &middot; not live';
  return '<div class="pstamp' + (b.stale ? ' stale' : '') + '">' + txt + '</div>';
}

function satsHtml(bndLeft){
  // The secondary sats line under the hero only exists in fiat mode: in sats
  // mode the hero already IS sats and repeating it would be noise.
  if (state.denom === 'sats') return '';
  var b = state.btc;
  if (!b || !b.ok || !b.price_bnd) return '';
  var amt = Number(bndLeft) || 0;
  if (amt <= 0) return '';
  var sats = Math.round((amt / b.price_bnd) * 100000000);
  var ago = priceAgeLabel(b.fetched_at);
  var h = '<div class="sats' + (b.stale ? ' stale' : '') + '">';
  h += '<b>' + sats.toLocaleString() + ' sats</b>';
  // Quote the anchor in USD: it is the number anyone can check against any
  // exchange. BND is our spending currency but nobody prices a bitcoin in it.
  h += '<span>@ ' + usdLabel(b.price_usd) + '/BTC' +
       (b.price_bnd ? ' &middot; ' + money(b.price_bnd) : '') +
       (ago ? ' &middot; ' + ago : '') +
       (b.stale ? ' &middot; not live' : '') + '</span>';
  h += '</div>';
  return h;
}

function catsHtml(map){
  var keys = Object.keys(map);
  if (!keys.length) return '<div class="empty">No expenses yet</div>';
  keys.sort(function(a, b){ return map[b] - map[a]; });

  // A bar makes the ranking readable in one glance; the number alone forces
  // the reader to compare seven figures in their head.
  var top = 0;
  for (var t = 0; t < keys.length; t++) {
    if (map[keys[t]] > top) top = map[keys[t]];
  }

  var out = ['<div class="cats">'];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = map[k];
    var w = top > 0 ? Math.max(2, (v / top) * 100) : 0;
    out.push('<div class="cat"><div class="cat-top"><span class="nm">' + esc(k) +
      '</span><span class="vl">' + money(v) + '</span></div>' +
      '<div class="cat-bar"><i style="width:' + w.toFixed(1) + '%"></i></div></div>');
  }
  out.push('</div>');
  return out.join("");
}

function listHtml(){
  if (!state.rows) return '<div class="empty">Loading...</div>';
  if (!state.rows.length) return '<div class="card"><div class="empty">No expenses yet</div></div>';
  // The list carries its own context so you never have to leave the tab to
  // find out what the month cost.
  // Truncation is decided by what the SERVER said it held, not by comparing
  // rows.length to the entry count. Those two measure different things: rows
  // includes Sats rows, the count excludes them, so rows.length < listCount
  // was true while rows had still been cut, which hid the button.
  var more = !!state.listHasMore;
  var head = '<div class="listhead"><b>' + money(state.listTotal) + '</b>' +
    '<span>' + state.listCount + ' entr' + (state.listCount === 1 ? 'y' : 'ies') +
    (more ? ' &middot; newest ' + state.rows.length + ' shown' : '') +
    '</span></div>';
  // A 50-row cap with no way past it made a busy month look like rows had gone
  var out = [head, '<div class="card sw-wrap">'];
  for (var i = 0; i < state.rows.length; i++) {
    var r = state.rows[i];
    var open = (state.swipeOpen === r.rowNumber);
    // data-id is the safe handle: row numbers go stale the moment the Sheet is
    // sorted or a row is deleted elsewhere.
    out.push('<div class="sw' + (open ? ' open' : '') + '" data-rownum="' + r.rowNumber +
      '" data-id="' + esc(r.id || '') + '">');
    out.push('<div class="sw-actions">');
    out.push('<button class="sw-edit" data-act="edit" data-id="' + esc(r.id || '') + '" data-rownum="' + r.rowNumber + '">Edit</button>');
    out.push('<button class="sw-del" data-act="del" data-id="' + esc(r.id || '') + '" data-rownum="' + r.rowNumber + '">Delete</button>');
    out.push('</div>');
    out.push('<div class="sw-face row">' +
      '<div class="l">' +
      '<div class="v">' + esc(r.vendor) + '</div>' +
      '<div class="t">' + esc(ooDay(r.date)) + ' &middot; ' + esc(r.category) +
      ' &middot; ' + esc(r.payment) + '</div></div>' +
      '<div class="a">' + money(r.amount, r.currency) + '</div></div>');
    out.push('</div>');
  }
  out.push('</div>');
  // At the end of the list, after the last row: it is where you look when you
  // reach the bottom and still cannot find an entry. It was under the header,
  // which is where nobody scrolls to and where it read as a control.
  if (more) {
    out.push('<div class="listmore">' +
      '<button class="btn alt" type="button" data-act="more"' +
      (state.busy ? ' disabled' : '') + '>' +
      (state.busy ? 'Loading...' : 'Show older entries') + '</button>' +
      '<div class="hint" style="text-align:center">' + state.rows.length +
      ' newest shown</div></div>');
  }
  return out.join("");
}


function editSheetHtml(){
  var r = state.editRow;
  if (!r) return '';
  var h = '<div class="mask" data-act="mask"><div class="sheet">';
  h += '<h2 class="ehead">Edit expense</h2>';
  h += '<div class="hint">Fix anything that was misread.</div>';
  h += '<form data-act="update" autocomplete="off">';
  h += '<label>Amount</label>';
  h += '<input type="number" step="0.01" inputmode="decimal" id="e_amount" value="' + (Number(r.amount) || 0).toFixed(2) + '" required>';
  h += '<label>Vendor</label>';
  h += '<input type="text" id="e_vendor" value="' + esc(r.vendor) + '" required>';
  h += '<label>Category</label>';
  h += '<select id="e_category">' + catOptions(r.category) + '</select>';
  h += '<label>Payment</label>';
  h += '<select id="e_payment">' + payOptions(r.payment) + '</select>';
  h += '<label>Date</label>';
  h += '<input type="date" id="e_date" value="' + esc(r.date || todayISO()) + '">';
  h += '<button class="btn" type="submit"' + (state.busy ? ' disabled' : '') + '>' +
    (state.busy ? 'Saving...' : 'Save changes') + '</button>';
  h += '</form>';
  h += '<button class="btn alt" data-act="closeEdit" type="button">Cancel</button>';
  h += '</div></div>';
  return h;
}

function payOptions(sel){
  var opts = ['Cash','Card','Transfer'];
  var out = [];
  for (var i = 0; i < opts.length; i++) {
    out.push('<option' + (opts[i] === sel ? ' selected' : '') + '>' + opts[i] + '</option>');
  }
  return out.join('');
}

function navHtml(){
  return '<nav>' +
    '<button data-act="tab" data-tab="summary" class="' + (state.tab === "summary" ? "on" : "") + '">Summary<span>' + (state.tab === "summary" ? "&bull;" : "") + '</span></button>' +
    '<button data-act="tab" data-tab="debts" class="' + (state.tab === "debts" ? "on" : "") + '">Debts<span>' + (state.tab === "debts" ? "&bull;" : "") + '</span></button>' +
    '<button data-act="tab" data-tab="list" class="' + (state.tab === "list" ? "on" : "") + '">Expenses<span>' + (state.tab === "list" ? "&bull;" : "") + '</span></button>' +
    '<button class="navadd" data-act="open" aria-label="Add expense">' +
      '<svg viewBox="0 0 24 24" width="24" height="24" fill="none" ' +
      'stroke="currentColor" stroke-width="1.9" stroke-linecap="round">' +
      '<path d="M12 5v14M5 12h14"/></svg></button>' +
    '</nav>';
}

// Display-only label map. The STORED value stays 'Reimbursed' -- it is what
// the sheet holds, what the Gemini prompt expects, and what the server's sign
// negation compares against. Only the <option> text is friendlier, so a new
// user reads the mechanic at the moment they pick it. Never put the label in
// CATS: that string round-trips to the server as a category and would miss
// every 'Reimbursed' check, storing money-back rows as positive spend.
var CAT_LABELS = { 'Reimbursed': 'Reimbursed (money paid back)' };

function catOptions(sel){
  var out = [];
  for (var i = 0; i < CATS.length; i++) {
    var val = CATS[i];
    var lbl = CAT_LABELS[val] || val;
    out.push('<option value="' + esc(val) + '"' + (val === sel ? " selected" : "") + '>' +
             esc(lbl) + '</option>');
  }
  return out.join("");
}

/* ---------------- standalone install screen ---------------- */

// Shown when an unbound copy has no SHEET_ID. Everything the README used to
// ask for in Project Settings happens here instead.
function installSheetHtml(){
  var d = state.installDraft || {};
  var h = '<div class="mask" data-act="mask"><div class="sheet">';
  h += '<h2 class="ehead">Almost there</h2>';
  h += '<div class="hint">Stackable will make a spreadsheet in your own Drive to hold ' +
       'your data. Nothing is sent anywhere except that sheet and Google&rsquo;s Gemini API.</div>';
  h += '<form data-act="saveinstall" autocomplete="off">';
  h += '<label>Gemini API key &mdash; highly recommended</label>';
  h += '<input type="text" id="i_key" inputmode="text" autocomplete="off" ' +
       'autocapitalize="off" autocorrect="off" spellcheck="false" ' +
       'placeholder="AIza... or AQ.Ab..." value="' + esc(d.geminiKey || '') + '">';
  h += '<div class="hint" style="margin-top:-6px">Free from aistudio.google.com. ' +
       'Stored only in this script&rsquo;s properties, never sent to us.</div>';
  // Say plainly what the key is FOR. Without this a user thinks the app is
  // broken without one, and the key step reads as a toll gate.
  h += '<div class="hint" style="margin-top:6px">The key does one job: reading ' +
       'photographed receipts. Everything else works without it.</div>';
  if (state.error) h += '<div class="err">' + esc(state.error) + '</div>';
  h += '<button class="btn" type="submit"' + (state.busy ? ' disabled' : '') + '>' +
       (state.busy ? 'Connecting...' : 'Connect') + '</button>';
  h += '</form>';
  // Real skip path. The README promised one for a day and the screen did not
  // have it, so a user without a key was simply blocked.
  if (!state.busy) {
    h += '<button class="btn alt" data-act="skipKey" type="button">' +
         'Skip for now &mdash; I&rsquo;ll add it later</button>';
    h += '<div class="hint" style="text-align:center;margin-top:10px">You can add ' +
         'the key any time from Settings.</div>';
  }
  h += '</div></div>';
  return h;
}

// Skip the key entirely: still create the sheet, just leave GEMINI_KEY unset.
// Scanning then reports 'GEMINI_KEY is not set' inside the sheet, which is a
// real message the user can act on later from Settings.
function skipKeyInstall(){
  state.error = null;
  render();
  run("saveInstall", { sheetId: "", geminiKey: "" }, function(res){
    if (!res || res.ok !== true) {
      state.error = (res && res.error) ? res.error : "Could not set up your sheet";
      render();
      return;
    }
    state.needsInstall = false;
    state.installDraft = null;
    state.onboardChecked = false;
    toast(res.created ? "Made you a spreadsheet called Stackable" : "Connected");
    loadSummary();
  });
}

function saveInstallFromForm(){
  var sheetId = "";
  var key = valOf("i_key");
  state.installDraft = { sheetId: sheetId, geminiKey: key };
  state.error = null;
  render();
  run("saveInstall", { sheetId: sheetId, geminiKey: key }, function(res){
    if (!res || res.ok !== true) {
      state.error = (res && res.error) ? res.error : "Could not connect";
      render();
      return;
    }
    state.needsInstall = false;
    state.installDraft = null;
    state.onboardChecked = false;
    if (res.created) toast("Made you a spreadsheet called Stackable");
    loadSummary();
  });
}

// Opens the user's own spreadsheet in a new tab. Server-side because the id
// lives in Script Properties; the client never holds it.
function settingsSheetHtml(){
  var s = state.settingsDraft || (state.summary && state.summary.settings) || {};
  var first = !!state.needsSetup;
  var h = '<div class="mask" data-act="mask"><div class="sheet">';
  h += '<h2 class="ehead">' + (first ? 'Set up your numbers' : 'Settings') + '</h2>';
  h += '<div class="hint">' + (first
    ? 'Nothing here is tracked until these are yours. Takes 30 seconds.'
    : 'These drive the budget bar, the pace projection and the stack rate.') + '</div>';

  h += '<form data-act="savesettings" autocomplete="off">';
  h += '<label>Your salary, take-home</label>';
  h += sfield('s_income_salary', s.income_salary);
  h += '<label>Other household income <span style="font-weight:400;color:#8a8a90">(optional)</span></label>';
  h += sfield('s_income_wife', s.income_wife);
  h += '<div class="hint">Partner income, side work. Leave blank if it is just you.</div>';
  h += '<label>Fixed monthly costs</label>';
  h += sfield('s_fixed_total', s.fixed_total);
  // Fixed costs are LIVING COSTS, and that includes the whole instalment
  // payment on a mortgage or a car loan: rent and a loan instalment are the
  // same line in a budget. Only revolving (credit card) payments move below,
  // because totalDebtPayment() subtracts those separately -- a card payment
  // entered here would be counted twice. The copy has to match that split; it
  // used to name the car loan as a fixed cost and then forbid loan repayments
  // on the very next line.
  h += '<div class="hint">Mortgage or rent, car loan instalments, school fees, ' +
       'utilities, internet, insurance. Not groceries, eating out or fuel &mdash; ' +
       'those go below.</div>';
  h += '<div class="hint" style="margin-top:4px">Include the whole loan instalment ' +
       'here. Credit card payments are the exception &mdash; put those in Debts ' +
       'below so the interest can be priced.</div>';
  var tgt = Number(s.variable_budget) || 0;
  h += '<label>Monthly spending target</label>';
  h += sfield('s_variable_budget', tgt);
  // Say exactly what gets measured against it. The bar includes One-off, so a
  // target that was set thinking "ordinary months" will look blown through on
  // any month with a birthday in it -- that is the truth, but say so here.
  h += '<div class="hint">Everything that leaves the account: groceries, bills, ' +
       'and one-off costs like a birthday or a repair. Leave 0 and the app fills ' +
       'it from your own history as months complete.</div>';
  if (tgt <= 0) {
    var autoT = autoTargetText();
    if (autoT) h += '<div class="hint" id="auto_target" style="margin-top:2px">' + autoT + '</div>';
  } else {
    // With a target set, the history average is only worth showing if it exists
    // and says something about the number above it. It used to live under Total
    // income, where it read as a qualifier on the income figure.
    h += suggHtml(s);
  }

  h += '<label>Currency</label>';
  h += '<select id="s_currency">' + currencyOptions(s.currency) + '</select>';
  h += '<div class="hint">Used for every amount shown. Amounts are not converted &mdash; enter your numbers in this currency.</div>';

  h += '<div class="hint">Tap any big figure on the Summary to switch it between ' + esc(CUR) + ' and sats.</div>';

  var sal = Number(s.income_salary) || 0;
  var wife = Number(s.income_wife) || 0;
  var totalIncome = sal + wife;
  // Settings is a form, not a dashboard. It shows the one number the user
  // typed in, and nothing derived: income minus budget minus fixed was a
  // left-over from the old budget model that contradicted the hero.
  h += '<div class="sum" style="margin-top:22px"><span>Total income</span><b id="sum_income">' + money(totalIncome) + '</b></div>';

  h += debtSectionHtml();

  // Receipt scanning comes LAST: it is the one optional feature, and burying
  // it under the numbers people actually came here to change keeps the first
  // screen about money. Lives in Script Properties, so the field is always
  // rendered empty - a masked value would let a stray Save overwrite the real
  // key.
  var hasKey = !!(state.hasGeminiKey || (state.summary && state.summary.hasKey));
  // The SAVED model, shown as a placeholder. The field itself only ever holds
  // what the user is typing, so a re-render can never write a masked value back.
  var savedModel = (state.summary && state.summary.geminiModel) || state.geminiModel || '';
  h += '<div class="dsec">';
  h += '<div class="dsec-h">Receipt scanning</div>';
  h += '<div class="hint" style="margin-top:0">Optional. Only used to read photographed receipts; everything else works without it.</div>';
  h += '<label>Gemini API key</label>';
  h += '<input type="text" id="s_gemini_key" inputmode="text" autocomplete="off" ' +
       'autocapitalize="off" autocorrect="off" spellcheck="false" placeholder="' +
       (hasKey ? 'Key saved \u2713 \u2014 paste a new one to replace it' : 'AIza... or AQ.Ab...') +
       '" value="' + esc(s.geminiKey || '') + '">';
  h += '<div class="hint">' + (hasKey
    ? 'Saved. Paste a new key only if you want to replace it.'
    : 'Free from aistudio.google.com.') + '</div>';
  h += '<label>Gemini model <span style="font-weight:400;color:#8a8a90">(optional)</span></label>';
  // CLIENT_SCOPE: DEFAULT_GEMINI_MODEL is a SERVER constant. Referencing it
  // here threw "DEFAULT_GEMINI_MODEL is not defined" on every gear tap, which
  // killed render() and froze the tab bar. The client keeps its own copy,
  // exactly like CCY.
  h += '<input type="text" id="s_gemini_model" autocomplete="off" autocapitalize="off" ' +
       'spellcheck="false" placeholder="' + esc(savedModel || CLIENT_GEMINI_MODEL) +
       '" value="' + esc(s.geminiModel || '') + '">';
  h += '<div class="hint">' + (savedModel
    ? 'Using ' + esc(savedModel) + '. Leave blank unless a scan fails with 403/404.'
    : 'Leave blank unless a scan fails with 403/404 &mdash; then name a current Flash model.') + '</div>';
  h += '</div>';

  // Dismiss the keyboard before submitting. On a phone the numeric keypad
  // covers the bottom of the sheet, which is exactly where these buttons sit,
  // so a tap on Save landed on the keyboard instead and the field reopened.
  // This button is labelled by what THIS FORM is doing, not by the global
  // request counter. Background loads (suggestion, debts) kept the counter up
  // and the Save button flickered to "Saving..." while the user was only
  // scrolling. state.saving is set by doSaveSettings and nothing else.
  var saving = !!state.saving;
  h += '<button class="btn" type="submit" onclick="if(document.activeElement&&document.activeElement.blur)document.activeElement.blur();"' + (saving ? ' disabled' : '') + '>' +
    (saving ? 'Saving...' : (first ? 'Start tracking' : 'Save')) + '</button>';
  if (saving) {
    h += '<div class="hint" style="text-align:center">Saving your numbers&hellip;</div>';
  }
  h += '</form>';
  h += '<button class="btn alt" data-act="closeSettings" type="button">' + (first ? 'Not now' : 'Cancel') + '</button>';
  // Extra room so the last control is never under the home indicator or a
  // still-closing keyboard.
  h += '<div style="height:26px"></div>';
  h += '</div></div>';
  return h;
}

// Debt rows live inside the settings form so there is one place to enter data.
// Fields are plain inputs named d_<index>_<field>, read back on submit.
function debtSectionHtml(){
  var list = state.debts || [];
  var h = '<div class="dsec" id="dsec">';
  h += '<div class="dsec-h">Debts</div>';
  h += '<div class="hint">Balance, rate and what you actually pay. Leave empty if none.</div>';
  // The type picker decides whether the payment is subtracted from stackable.
  // Say which is which HERE, where the choice is made, instead of leaving the
  // user to discover that a car loan is already inside Fixed costs.
  h += '<div class="hint" style="margin-top:2px">Credit cards are subtracted from ' +
       'your stackable amount. Instalments (mortgage, car loan) are already inside ' +
       'Fixed costs, so they are shown here for the interest maths only.</div>';

  for (var i = 0; i < list.length; i++) {
    var d = list[i];
    h += '<div class="dcard">';
    h += '<div class="drow">';
    h += '<input type="text" id="d_' + i + '_name" class="dname" placeholder="Name (e.g. Chase Sapphire)" value="' + esc(d.name) + '">';
    h += '<button type="button" class="dx" data-act="delDebt" data-didx="' + i + '" aria-label="Remove"' +
         (state.busy ? ' disabled' : '') + '>&times;</button>';
    h += '</div>';
    h += '<select id="d_' + i + '_type" class="dtype">' +
         '<option value="revolving"' + (d.type === 'revolving' ? ' selected' : '') + '>Revolving (credit card)</option>' +
         '<option value="fixed"' + (d.type === 'fixed' ? ' selected' : '') + '>Instalment (mortgage, car loan)</option>' +
         '</select>';
    h += '<div class="dgrid">';
    h += dfield('Balance', 'd_' + i + '_balance', d.balance);
    h += dfield('Rate %/yr', 'd_' + i + '_rate', d.rate);
    h += dfield('Minimum', 'd_' + i + '_minimum', d.minimum);
    h += dfield('You pay', 'd_' + i + '_payment', d.payment);
    h += '</div>';
    h += '<input type="hidden" id="d_' + i + '_id" value="' + esc(d.id) + '">';
    h += '<input type="hidden" id="d_' + i + '_notes" value="' + esc(d.notes) + '">';
    h += '</div>';
  }

  h += '<button type="button" class="dadd" data-act="addDebt"' + (state.busy ? ' disabled' : '') +
       '>+ Add a debt</button>';
  // Settings is for entering debts, not for reading back totals. The Debts
  // tab owns those figures, so a half-populated form never shows a figure
  // that disagrees with what the app actually computed.
  h += '</div>';
  return h;
}

function dfield(lbl, id, val){
  var n = Number(val) || 0;
  var shown = (n === 0) ? '' : n.toFixed(2);
  return '<div class="df"><label class="dl">' + lbl + '</label>' +
    '<input type="number" step="0.01" min="0" inputmode="decimal" id="' + id +
    '" value="' + shown + '" placeholder="0.00" onfocus="this.select()"></div>';
}

// Auto-fill: while the target is 0, show what the user's own history implies,
// so the field is never a number they had to invent.
function autoTargetText(){
  var sg = state.suggestion;
  if (!sg || !sg.ok) return '';
  if (sg.ready && sg.average > 0) {
    return 'From your last ' + sg.months + ' month' + (sg.months === 1 ? '' : 's') +
           ': <b>' + money(sg.average) + '</b> &mdash; used automatically until you set your own.';
  }
  if (sg.months && sg.months > 0) {
    return 'Tracking so far. Need 2 complete months before a target can be measured.';
  }
  return 'No complete month yet. Log for 30 days and a target appears here on its own.';
}

function suggHtml(s){
  var sg = state.suggestion;
  if (!sg || !sg.ok) return '';
  // No not-ready branch: with too little history, autoTargetText() under the
  // target field already explains it. Two boxes said the same thing.
  if (!sg.ready) return '';
  var cur = Number(s.variable_budget) || 0;
  var diff = Math.round((sg.average - cur) * 100) / 100;
  var h = '<div class="sugg">';
  h += '<div class="sugg-t">Your last ' + sg.months + ' month' + (sg.months === 1 ? '' : 's') +
       ' averaged <b>' + money(sg.average) + '</b></div>';
  var parts = [];
  for (var i = 0; i < sg.used.length; i++) {
    parts.push(monthLabel(sg.used[i].month) + ' ' + money(sg.used[i].spent));
  }
  h += '<div class="sugg-d">' + parts.join(' &middot; ') + '</div>';
  if (Math.abs(diff) >= 1) {
    h += '<button class="sugg-b" type="button" data-act="useSuggestion" data-amt="' + sg.average + '">Use ' +
         money(sg.average) + '</button>';
  } else {
    h += '<div class="sugg-d">Your budget is already in line with this.</div>';
  }
  h += '</div>';
  return h;
}

function currencyOptions(sel){
  var codes = CCY.slice();
  codes.sort();
  var out = [];
  var cur = String(sel || 'USD').toUpperCase();
  for (var i = 0; i < codes.length; i++) {
    out.push('<option value="' + codes[i] + '"' +
      (codes[i] === cur ? ' selected' : '') + '>' + codes[i] + '</option>');
  }
  return out.join('');
}

function sfield(id, val){
  var n = Number(val);
  if (isNaN(n)) n = 0;
  // Blank when empty/zero so the field does not greet you with "0.00" that you
  // must delete before typing. select() grabs whatever is there on tap.
  var shown = (n === 0) ? '' : n.toFixed(2);
  return '<input type="number" step="0.01" min="0" inputmode="decimal" id="' + id +
    '" value="' + shown + '" placeholder="0.00" onfocus="this.select()" required>';
}

function sheetHtml(){
  var h = '<div class="mask" data-act="mask"><div class="sheet">';
  h += '<h2>Add expense</h2>';
  h += '<div class="hint">Photograph a receipt, or type it in.</div>';

  h += '<label>Receipt photo</label>';
  h += '<input type="file" accept="image/*" capture="environment" id="file" style="display:none">';
  // Busy is true while scanning AND while saving a typed expense. Only the
  // scan should relabel this button; otherwise typing an amount and saving
  // made the photo button claim it was reading a receipt.
  h += '<button class="btn ghost" data-act="pick" type="button"' + (state.scanning ? " disabled" : "") + '>' +
    (state.scanning ? "Reading receipt..." : (state.photo ? "Retake photo" : "Take photo")) + '</button>';
  if (state.photo) {
    h += '<img class="prev" src="' + state.photo + '" alt="receipt">';
  }
  if (state.scanning && state.photo) {
    h += '<div class="hint" style="text-align:center;margin-top:12px">Reading the receipt&hellip;</div>';
  }
  // The scan failure used to set state.error, which render() draws in the page
  // body BEHIND this modal. The user saw the button flip back to "Retake photo"
  // and nothing else. Errors from the scan belong inside the sheet.
  if (!state.scanning && state.scanError) {
    h += '<div class="err">' + esc(state.scanError) + '</div>';
  }

  h += '<div class="sep">or enter manually</div>';
  h += '<form data-act="save" autocomplete="off">';
  h += '<label>Amount</label>';
  h += '<input type="number" step="0.01" inputmode="decimal" id="f_amount" placeholder="0.00" required>';
  h += '<label>Vendor</label>';
  h += '<input type="text" id="f_vendor" placeholder="Where" required>';
  h += '<label>Category</label>';
  h += '<select id="f_category">' + catOptions("Groceries & Household") + '</select>';
  h += '<label>Payment</label>';
  h += '<select id="f_payment"><option>Cash</option><option>Card</option><option>Transfer</option></select>';
  h += '<label>Date</label>';
  h += '<input type="date" id="f_date" value="' + todayISO() + '">';
  h += '<button class="btn" type="submit"' + (state.busy ? " disabled" : "") + '>' +
    (state.busy ? "Saving..." : "Save expense") + '</button>';
  h += '</form>';

  h += '<button class="btn alt" data-act="close" type="button">Cancel</button>';
  h += '</div></div>';
  return h;
}

function draftSheetHtml(){
  var d = state.draft;
  if (!d) return '';
  var conf = (typeof d.confidence === "number") ? d.confidence : null;
  var warn = "";
  // A date the model guessed wrong moves the expense out of this month, which
  // is how a saved grocery bill vanished from the Expenses list. Say it here.
  if (d.dateWarn) {
    warn += '<div class="err" style="background:#3a2e1a;border-color:#5a4a26;color:#ffd88a">' +
      esc(d.dateWarn) + ' Tap the date below to fix it.</div>';
  }
  if (conf !== null && conf < 0.7) {
    warn += '<div class="err" style="background:#3a2e1a;border-color:#5a4a26;color:#ffd88a">' +
      'Low confidence (' + Math.round(conf * 100) + '%) - check the amount and category.</div>';
  }
  var h = '<div class="mask"><div class="sheet">';
  h += '<h2 class="ehead">Check the receipt</h2>';
  h += '<div class="hint">Check what was read. Fix anything wrong, then save.</div>';
  h += warn;
  h += '<div class="card" style="margin:0 0 4px;text-align:center">' +
    '<div class="lbl">Amount</div>' +
    '<div class="big" style="font-size:34px">' + money(d.amount) + '</div></div>';
  h += '<form data-act="savedraft" autocomplete="off">';
  h += '<label>Amount</label>';
  h += '<input type="number" step="0.01" inputmode="decimal" id="d_amount" value="' +
    (isNaN(Number(d.amount)) ? "" : Number(d.amount).toFixed(2)) + '" required>';
  h += '<label>Category</label>';
  h += '<select id="d_category">' + catOptions(d.category) + '</select>';
  h += '<label>Vendor</label>';
  h += '<input type="text" id="d_vendor" value="' + esc(d.vendor) + '" required>';
  h += '<label>Payment</label>';
  h += '<select id="d_payment">' + payOptions(d.payment) + '</select>';
  h += '<label>Date</label>';
  h += '<input type="date" id="d_date" value="' + esc(d.date || todayISO()) + '">';
  h += '<button class="btn" type="submit"' + (state.busy ? ' disabled' : '') + '>' +
    (state.busy ? 'Saving...' : 'Save expense') + '</button>';
  h += '</form>';
  h += '<button class="btn alt" data-act="discardDraft" type="button">Discard</button>';
  h += '</div></div>';
  return h;
}

/* ---------- data ---------- */

// Requests that never answer are the one failure Apps Script does not report:
// neither handler fires, the counter never drops, and the UI says "Saving..."
// for ever. A watchdog per call converts that silent hang into a visible error.
var RUN_TIMEOUT_MS = 45000;

function run(fnName, payload, onOk){
  if (!window.google || !google.script || !google.script.run) {
    state.error = "Not running inside Apps Script. Open the /exec URL.";
    render();
    return;
  }
  // Breadcrumb: remembers which request is outstanding, so a stuck screen can
  // say what it is waiting for instead of just spinning.
  state.waitingOn = fnName;
  // A count, not a flag. Boot fires loadSummary -> btc + debts in parallel and
  // the first to finish used to clear 'busy' while the second was in flight,
  // which enabled Save/Add Debt mid-load and made beginScan swallow taps.
  busyStart();
  var settled = false;
  var timer = setTimeout(function(){
    if (settled) return;
    settled = true;
    busyEnd();
    state.scanning = false;
    // A timed-out settings save must not leave the Save button on "Saving..."
    // for the rest of the session. Nothing is in flight any more, so clear it.
    state.saving = false;
    state.waitingOn = null;
    state.error = fnName + " did not answer in " + Math.round(RUN_TIMEOUT_MS / 1000) +
      "s. The server may be busy or the deployment stale - reload the page and try again.";
    render();
  }, RUN_TIMEOUT_MS);
  var done = function(){
    if (settled) return;
    settled = true;
    clearTimeout(timer);
    state.waitingOn = null;
    busyEnd();
  };
  var runner = google.script.run
    .withSuccessHandler(function(res){
      done();
      try { onOk(res); }
      catch (e) { fail("Render error", e); }
    })
    .withFailureHandler(function(err){
      done();
      // A server throw or an Apps Script timeout skips the success callback
      // entirely, so any flag that callback was going to clear stayed set.
      // That is how a failed scan left the photo button stuck on "Reading
      // receipt..." until the sheet was reopened.
      state.scanning = false;
      // Same rule as the watchdog: a failed call is not a saving call.
      state.saving = false;
      state.error = "Server error: " + ((err && err.message) ? err.message : String(err));
      render();
    });
  runner[fnName](payload);
}

function loadSummary(){
  state.error = null;
  run("getSummary", { month: state.month }, function(res){
    if (res && res.ok === false) { state.error = "Server said: " + res.error; render(); return; }
    state.summary = res || { total:0, byCategory:{}, count:0 };
    // Every amount in the app reads.CUR, so it must follow the real setting.
    if (res && res.settings && res.settings.currency) CUR = String(res.settings.currency);
    maybeOnboard();
    render();
    if (!state.btc) loadBtcPrice();
    if (!state.debts) loadDebts();
    // Calibrating mode needs the measured average to offer a budget.
    if (!(Number(state.summary.budget) > 0) && !state.suggestion) loadSuggestion();
  });
}

// First run: a stranger who has never seen the app gets one screen explaining
// what it does and where the data lives, instead of a form demanding pay.
function welcomeHtml(){
  var h = '<div class="card welcome">';
  h += '<div class="wl-t">Know what you can stack.</div>';
  h += '<div class="wl-p">Log what you spend. The app works out what is ' +
       'genuinely left over at the end of the month &mdash; from your own ' +
       'numbers, not a guess.</div>';
  h += '<div class="wl-list">' +
       '<div class="wl-row"><i>1</i><span>Set your income and fixed costs &mdash; 30 seconds</span></div>' +
       '<div class="wl-row"><i>2</i><span>Photograph entries as you spend</span></div>' +
       '<div class="wl-row"><i>3</i><span>Watch what is stackable move</span></div>' +
       '</div>';
  h += '<div class="wl-note">Your receipts stay in your own Google Drive. ' +
       'The app writes to a spreadsheet you own and sends nothing to anyone else.</div>';
  h += '<button class="btn" data-act="beginSetup" type="button">Set up my numbers</button>';
  h += '</div>';
  return h;
}

// First run only: if the install was never configured, force the setup screen.
// The template ships with the author's numbers, and a stranger trusting those
// is worse than a blank app. Runs once per page load, never after saving.
function maybeOnboard(){
  if (state.onboardChecked) return;
  state.onboardChecked = true;
  if (!state.summary || state.summary.configured) return;
  // Show what the app does first; the form opens only when they ask for it.
  state.needsWelcome = true;
}

function loadList(){
  state.error = null;
  run("getList", { limit: state.listLimit, month: state.month }, function(res){
    if (res && res.ok === false) { state.error = "Server said: " + res.error; render(); return; }
    state.rows = (res && res.rows) ? res.rows : [];
    state.listTotal = (res && typeof res.total === 'number') ? res.total : 0;
    state.listCount = (res && typeof res.count === 'number') ? res.count : state.rows.length;
    state.listHasMore = !!(res && res.hasMore);
    render();
  });
}

function openSettings(){
  state.settingsOpen = true;
  state.settingsDraft = null;
  state.error = null;
  // A freshly opened form is not saving anything. If a previous session left
  // busy set (an abandoned request, a stale callback), the sheet used to open
  // already showing "Saving..." with values that had not arrived yet. Reset
  // here so the form always starts honest; the two loads below re-raise it.
  state.pending = 0;
  state.busy = false;
  state.waitingOn = null;
  // Do NOT start the two loads with the sheet already painted: their callbacks
  // each re-render, and a user who starts typing between them gets their input
  // replaced. Paint once from what we have, then let the loads update in place.
  render();
  loadSuggestion();
  loadDebts();
}

function loadDebts(){
  // Only a load issued BY the Settings form cares whether that form is still
  // open. A Debts-tab or boot load has no form to protect, and bailing on
  // settingsOpen made it a silent no-op - the tab then showed "No debt" with
  // two debts saved in the sheet. The caller states its intent instead of
  // the callback guessing it.
  var forSettings = !!state.settingsOpen;
  run("getDebts", {}, function(res){
    // A FAILED load must not erase what we already have. Before this, any
    // error or timeout response replaced state.debts with [] - so the Debts tab
    // lost two saved debts and said "No debt" after a slow getDebts. Empty is
    // a valid answer; failure is not an answer at all.
    if (!res || res.ok === false) {
      if (res && res.error) state.error = "Debts could not load: " + res.error;
      render();
      return;
    }
    if (forSettings && !state.settingsOpen) { render(); return; }
    // Never overwrite the list while the Settings form is open. The user has
    // been editing those rows (maybe adding one), and this response is a
    // SNAPSHOT of what the sheet held when the request left. Applying it
    // replaced their rows with the old set - Add a debt then Save wrote the
    // stale list back, and the debts tab showed "no debt" because the sheet
    // really had none. The form's own captured state wins while it is open.
    if (state.settingsOpen) {
      // A tab load that resolves while the form is open must NOT touch the
      // form's rows - the user may already be typing in them.
      if (!forSettings) return;
      // The user may already have tapped "+ Add a debt" since this request left.
      // This response is a snapshot from before that tap, so applying it would
      // delete a row they just created. Once the form holds rows, its own state
      // wins - capturing them into state.debts keeps them, then we bail.
      //
      // NB: this must NOT be the old "if settingsOpen return" guard. That one
      // also rejected the legitimate first load and produced the "No debt"
      // screen over a live figure.
      if (state.debts && state.debts.length) {
        captureSettingsDraft();
        captureDebts();
        render();
        return;
      }
    }
    // Only a successful response may replace the list, and keep a copy of the
    // ENRICHED version: a Settings draft (no payoff) must never clobber what
    // the Debts tab needs, and Cancel should render instantly from this.
    state.debts = res.debts || [];
    state.debtsGood = state.debts;
    // History is measured, not assumed: keep whatever the server observed.
    state.trackedSince = (res && res.trackedSince) ? res.trackedSince : '';
    state.trackedInterest = (res && typeof res.trackedInterest === 'number') ? res.trackedInterest : 0;
    state.trackedMonths = (res && typeof res.trackedMonths === 'number') ? res.trackedMonths : 0;
    state.trackedError = (res && res.trackedError) ? String(res.trackedError) : '';
    state.debtLoadedAt = new Date().toLocaleTimeString();
    state.lastRaw = res;
    // Same trap as loadSuggestion: never re-render the open form empty-handed.
    if (state.settingsOpen) { captureSettingsDraft(); captureDebts(); }
    render();
  });
}

// Reads the current form values into state before a re-render, so adding or
// removing a row never silently wipes what was already typed.
function captureDebts(){
  var list = state.debts || [];
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var old = list[i];
    var nm = document.getElementById("d_" + i + "_name");
    if (!nm) { out.push(old); continue; }
    out.push({
      id: valOf("d_" + i + "_id"),
      name: nm.value,
      type: valOf("d_" + i + "_type"),
      balance: numOf("d_" + i + "_balance"),
      rate: numOf("d_" + i + "_rate"),
      minimum: numOf("d_" + i + "_minimum"),
      payment: numOf("d_" + i + "_payment"),
      notes: valOf("d_" + i + "_notes")
    });
  }
  // Any row still lacking an id (older state, a partial render) gets one now,
  // and keeps it: the hidden field carries it back on the next render.
  for (var j = 0; j < out.length; j++) {
    if (!out[j].id) out[j].id = "d" + Date.now() + j + Math.floor(Math.random() * 1000);
  }
  state.debts = out;
}

function valOf(id){
  var el = document.getElementById(id);
  return el ? String(el.value) : "";
}

function numOf(id){
  var el = document.getElementById(id);
  if (!el) return 0;
  var v = parseFloat(el.value);
  return isNaN(v) ? 0 : v;
}

// Any re-render rebuilds the Settings form from state, so EVERY field the user
// typed must be captured first. captureDebts() existed but nothing captured
// the income/target/currency fields, so tapping "Add a debt" during first-run
// setup wiped everything typed above it. That was silent data loss.
function captureSettingsDraft(){
  var curEl = document.getElementById("s_currency");
  var prev = state.settingsDraft || {};
  state.settingsDraft = {
    income_salary: readSettingField("s_income_salary"),
    income_wife: readSettingField("s_income_wife"),
    fixed_total: readSettingField("s_fixed_total"),
    variable_budget: readSettingField("s_variable_budget"),
    currency: curEl ? curEl.value : (prev.currency || "BND"),
    // Without these, any re-render while the user was typing a key (debts
    // loading, Add a debt, the suggestion landing) wiped the field mid-paste.
    geminiKey: (typeof valOf === "function" ? valOf("s_gemini_key") : "") || prev.geminiKey || "",
    geminiModel: (typeof valOf === "function" ? valOf("s_gemini_model") : "") || prev.geminiModel || ""
  };
}

function addDebtRow(){
  captureSettingsDraft();
  captureDebts();
  // Mint the id HERE, once, when the row is created. Leaving it blank meant the
  // server invented one at save time via Date.now(), so every rename produced a
  // new id and a new orphan row in DebtHistory.
  state.debts = (state.debts || []).concat([{
    id: "d" + Date.now() + Math.floor(Math.random() * 1000),
    name: "", type: "revolving", balance: 0, rate: 0,
    minimum: 0, payment: 0, notes: ""
  }]);
  render();
}

function deleteDebtRow(idx){
  captureSettingsDraft();
  captureDebts();
  var list = (state.debts || []).slice();
  if (idx < 0 || idx >= list.length) return;
  list.splice(idx, 1);
  state.debts = list;
  render();
}

// One-tap adoption of the measured budget. Reuses saveSettings so the
// Settings tab stays the single source of truth.
function adoptBudget(amt){
  render();
  run("saveSettings", { settings: { variable_budget: amt } }, function(res){
    if (!res || res.ok === false) {
      state.error = "Could not set budget: " + ((res && res.error) ? res.error : "unknown");
      render();
      return;
    }
    state.calibDismissed = true;
    if (state.summary && res.settings) state.summary.settings = res.settings;
    loadSummary();
    toast("Budget set to " + money(amt));
  });
}

function loadBtcPrice(){
  run("getBtcPrice", {}, function(res){
    state.btc = (res && res.ok) ? res : null;
    render();
  });
}

function loadSuggestion(){
  run("getBudgetSuggestion", { months: 3 }, function(res){
    state.suggestion = (res && res.ok) ? res : null;
    // Same rule as loadDebts: a late answer for a closed form is stale.
    if (!state.settingsOpen) { render(); return; }
    // This callback lands AFTER the Settings form is on screen, so capture the
    // live DOM before re-rendering or the typed values are lost.
    if (state.settingsOpen) { captureSettingsDraft(); captureDebts(); }
    render();
  });
}

function closeSettings(){
  // Cancel used to leave state.debts holding the form's DRAFT rows: name,
  // balance, rate and nothing else. Those have no payoff, so the Debts tab
  // drew the leak card and then skipped every debt card - a screen reading
  // "No debt" over a figure that came from the debts. Save avoided it only
  // because saveDebts returns the enriched list. Drop the draft and fetch the
  // real one.
  // Restore the last ENRICHED list before painting. Draft rows carry no
  // payoff, so painting them was the whole "No debt" bug; painting cached good
  // rows makes Cancel instant AND correct.
  if (state.debtsGood) state.debts = state.debtsGood;
  state.settingsOpen = false;
  state.settingsDraft = null;
  state.needsSetup = false;
  // Abandon anything still in flight. openSettings fires getBudgetSuggestion
  // and getDebts, and their callbacks outlive the sheet; a stale response
  // arriving after the close used to leave the counter and the form's own
  // flags disagreeing, which is the "reopen and it hangs" report. Every
  // callback now checks state.settingsOpen and bails, so the counter must be
  // cleared here rather than waiting for responses that are being dropped.
  state.pending = 0;
  state.busy = false;
  state.saving = false;
  state.waitingOn = null;
  // Paint immediately. Waiting on a refetch here made Cancel feel like a 3-4s
  // freeze; the tab renders now from cached rows and the refresh lands after.
  render();
  // Refresh in the background after painting. settingsOpen is already false,
  // so loadDebts is allowed to replace the list when the answer arrives.
  if (state.tab === "debts") loadDebts();
}

function readSettingField(id){
  var el = document.getElementById(id);
  if (!el) return null;
  var v = parseFloat(el.value);
  if (isNaN(v)) return null;
  return Math.round(v * 100) / 100;
}

function doSaveSettings(){
  var payload = {
    income_salary: readSettingField("s_income_salary"),
    income_wife: readSettingField("s_income_wife"),
    fixed_total: readSettingField("s_fixed_total"),
    variable_budget: readSettingField("s_variable_budget"),
    currency: (document.getElementById("s_currency") ? document.getElementById("s_currency").value : "BND"),
    // Non-empty only; blank means "keep the stored key/model".
    geminiKey: (valOf("s_gemini_key") || "").trim(),
    geminiModel: (valOf("s_gemini_model") || "").trim()
  };
  // Only salary is truly required. "Other income" is blank for a single user
  // and zero fixed costs is a real answer for someone living at home, so both
  // blank fields read as 0 rather than blocking the save. A single user who
  // could not save Settings was the worst of these.
  if (payload.income_salary === null) { toast("Enter your take-home salary", true); return; }
  var optional = ["income_wife","fixed_total","variable_budget"];
  for (var i = 0; i < optional.length; i++) {
    if (payload[optional[i]] === null) payload[optional[i]] = 0;
  }
  var all = ["income_salary"].concat(optional);
  for (var j = 0; j < all.length; j++) {
    if (payload[all[j]] < 0) { toast("No negative numbers", true); return; }
  }

  // Read the debt rows BEFORE any re-render: render() rebuilds the form DOM
  // from state, so capturing afterwards would read a fresh, empty form and
  // silently drop whatever the user just typed.
  captureDebts();
  var debtsPayload = [];
  for (var di = 0; di < (state.debts || []).length; di++) {
    var dd = state.debts[di];
    if (!String(dd.name || "").trim()) continue;
    debtsPayload.push(dd);
  }

  // No external busy pair: run() owns the counter, so a failure or the 45s
  // watchdog cannot leave pending stuck at 1. state.saving is the button's own
  // flag and is cleared by run() on failure/timeout and by finishSettingsSave
  // on success.
  state.saving = true;
  state.error = null;
  render();

  run("saveSettings", { settings: payload }, function(res){
    if (!res || res.ok === false) {
      state.saving = false;
      state.error = "Save failed: " + ((res && res.error) ? res.error : "unknown");
      render();
      return;
    }
    run("saveDebts", { debts: debtsPayload }, function(dres){
      finishSettingsSave(res, dres);
    });
  });
}

function finishSettingsSave(res, dres){
    state.saving = false;
    if (!dres || dres.ok === false) {
      state.error = "Debts not saved: " + ((dres && dres.error) ? dres.error : "unknown");
      render();
      return;
    }
    state.debts = dres.debts || [];
    // Keep the Debts tab's cache in step with what was just saved. Without
    // this, reopening Settings and pressing Cancel before visiting the tab
    // restored the PRE-save list and the edit appeared to vanish.
    state.debtsGood = state.debts;
    {
    var wasSetup = state.needsSetup;
    state.settingsOpen = false;
    state.settingsDraft = null;
    state.needsSetup = false;
    // A price fetched under the old currency is quoted in the old currency.
    // Keeping it means the sats line lags one currency change behind -- that
    // is exactly how SGD ended up showing a USD-converted figure for 24 min.
    var newCur = res.settings && res.settings.currency;
    if (newCur && String(newCur).toUpperCase() !== String(CUR || 'USD').toUpperCase()) {
      state.btc = null;
    }
    if (newCur) CUR = String(newCur).toUpperCase();
    if (state.summary && res.settings) state.summary.settings = res.settings;
    loadSummary();
    toast(wasSetup ? "You're set up" : "Settings saved");
    }
}

function openSheet(){
  state.sheet = true;
  state.photo = null;
  state.error = null;
  // A failed scan left its red message in state, so reopening the sheet showed
  // the old error as if the new attempt had failed too.
  state.scanError = null;
  state.scanning = false;
  render();
}

function closeSheet(){
  state.sheet = false;
  state.photo = null;
  state.draft = null;
  // Reset the COUNTER too, or the next request's completion recomputes
  // busy = pending > 0 and the whole UI locks in its saving state.
  state.pending = 0;
  state.busy = false;
  state.scanError = null;
  state.scanning = false;
  render();
}


// Resolve the clicked element to the expense it refers to. Id first; the row
// number is only a fallback for a row the server could not give an id for.
function findRow(el){
  var id = el.getAttribute("data-id");
  if (!id || !state.rows) return null;
  for (var i = 0; i < state.rows.length; i++) {
    if (String(state.rows[i].id || '') === String(id)) return state.rows[i];
  }
  return null;
}

function openEdit(rowOrId){
  if (!state.rows) return;
  for (var i = 0; i < state.rows.length; i++) {
    if (Number(state.rows[i].rowNumber) === Number(rowOrId) ||
        (typeof rowOrId === 'object' && rowOrId && state.rows[i].id === rowOrId.id)) {
      state.editRow = state.rows[i];
      state.swipeOpen = null;
      state.error = null;
      render();
      return;
    }
  }
  toast("Could not find that expense", true);
}

function confirmDelete(rowOrId){
  var target = (rowOrId && typeof rowOrId === 'object') ? rowOrId : null;
  var rowNumber = target ? target.rowNumber : rowOrId;
  if (!target && state.rows) {
    for (var i = 0; i < state.rows.length; i++) {
      if (Number(state.rows[i].rowNumber) === Number(rowNumber)) target = state.rows[i];
    }
  }
  if (!target) { toast("Could not find that expense", true); return; }
  var label = target.vendor + " - " + money(target.amount);
  if (!window.confirm("Delete this expense?\\n\\n" + label)) return;

  state.error = null;
  render();
  run("deleteExpense", { id: target.id, rowNumber: rowNumber }, function(res){
    if (!res || res.ok === false) {
      state.error = "Delete failed: " + ((res && res.error) ? res.error : "unknown");
      render();
      return;
    }
    state.swipeOpen = null;
    loadList();
    toast("Deleted");
  });
}

function doUpdate(){
  var r = state.editRow;
  if (!r) return;

  var amtEl = document.getElementById("e_amount");
  var venEl = document.getElementById("e_vendor");
  var catEl = document.getElementById("e_category");
  var payEl = document.getElementById("e_payment");
  var datEl = document.getElementById("e_date");

  var amount = amtEl ? parseFloat(amtEl.value) : NaN;
  var vendor = venEl ? String(venEl.value).trim() : "";

  // A 'Reimbursed' row is stored as a negative number. Rejecting <= 0 here made
  // every reimbursed expense uneditable: the box pre-fills -50.00 and save
  // refused it. Only a zero or non-number is invalid.
  if (isNaN(amount) || amount === 0) { toast("Enter an amount", true); return; }
  if (!vendor) { toast("Enter a vendor", true); return; }

  state.error = null;
  render();

  run("updateExpense", {
    id: r.id,
    rowNumber: r.rowNumber,
    amount: amount,
    vendor: vendor,
    category: catEl ? catEl.value : r.category,
    payment: payEl ? payEl.value : r.payment,
    date: (datEl && datEl.value) ? datEl.value : r.date
  }, function(res){
    if (!res || res.ok === false) {
      state.error = "Update failed: " + ((res && res.error) ? res.error : "unknown");
      render();
      return;
    }
    state.editRow = null;
    loadList();
    toast("Saved changes");
    if (state.tab === "summary") loadSummary();
  });
}

function toggleSwipe(rowNumber){
  state.swipeOpen = (state.swipeOpen === rowNumber) ? null : rowNumber;
  render();
}

/* ---------- events (delegated, never lost) ---------- */

document.addEventListener("click", function(e){
  // Walk up to the nearest ancestor that OWNS a data-act. The old loop stopped
  // at the first node exposing getAttribute -- which is every element -- so a
  // tap on a child span of the Card payment row found no data-act and died.
  // That is why only the row's bare padding responded.
  var el = e.target;
  while (el && el !== document && !(el.getAttribute && el.getAttribute("data-act"))) {
    el = el.parentNode;
  }
  if (!el || el === document) return;
  var act = el.getAttribute("data-act");
  if (!act) return;

  if (act === "open") { openSheet(); return; }
  if (act === "close") { closeSheet(); return; }
  if (act === "oneoffDisc") { state.oneoffOpen = !state.oneoffOpen; render(); return; }
  if (act === "settings") { openSettings(); return; }
  if (act === "dismissCalib") { state.calibDismissed = true; render(); return; }
  if (act === "adoptBudget") {
    var amt = Number(el.getAttribute("data-amt")) || 0;
    if (amt > 0) adoptBudget(amt);
    return;
  }
  if (act === "simPay") {
    var sid = el.getAttribute("data-did") || "";
    var samt = Number(el.getAttribute("data-amt")) || 0;
    state.simPay[sid] = samt;
    render();
    return;
  }
  if (act === "addDebt") { addDebtRow(); return; }
  if (act === "addFirstDebt") {
    // One tap: open Settings, seed a blank row, land on the debt fields.
    openSettings();
    if (!state.debts || !state.debts.length) {
      state.debts = [{ id: "", name: "", type: "revolving", balance: 0, rate: 0,
                       minimum: 0, payment: 0, notes: "" }];
      render();
    }
    return;
  }
  if (act === "delDebt") {
    var di = Number(el.getAttribute("data-didx"));
    deleteDebtRow(di);
    return;
  }
  if (act === "useSuggestion") {
    var amt = Number(el.getAttribute("data-amt")) || 0;
    var vb = document.getElementById("s_variable_budget");
    if (vb) {
      vb.value = amt.toFixed(2);
      vb.dispatchEvent(new Event("input", { bubbles: true }));
    }
    return;
  }
  if (act === "denom") { toggleDenom(); return; }
  if (act === "heroDisc") {
    state.heroOpen = !state.heroOpen;
    render();
    return;
  }
  if (act === "beginSetup") {
    state.needsWelcome = false;
    state.settingsOpen = true;
    state.needsSetup = true;
    render();
    return;
  }
  // Form acts are NEVER routed from the CLICK handler. Any tap inside a
  // <form data-act="..."> walks up to the form and matches here, so the handler
  // fired on focus, not submit. saveinstall and savepay both live in the SUBMIT
  // handler only (see document.addEventListener("submit", ...)).
  if (act === "editPay") {
    state.payEdit = el.getAttribute("data-did");
    state.error = null; render(); return;
  }
  if (act === "closePay") { state.payEdit = null; state.error = null; render(); return; }
  // savepay is NOT routed here. It lives in the SUBMIT handler only - a tap on
  // #pay_amount bubbles to <form data-act="savepay"> and would save the old
  // value before the user typed anything.
  if (act === "closeSettings") { closeSettings(); return; }
  if (act === "skipKey") { skipKeyInstall(); return; }
  // Backdrop tap must close whichever sheet is open. It used to call
  // closeSheet() only, so tapping the dim area behind Settings, Edit or the Pay
  // editor did nothing and the app looked frozen.
  if (act === "mask") {
    if (e.target !== el) return;
    if (state.payEdit) { state.payEdit = null; state.error = null; render(); return; }
    if (state.editRow) { state.editRow = null; render(); return; }
    if (state.settingsOpen) { closeSettings(); return; }
    closeSheet();
    return;
  }
  if (act === "prevMonth" || act === "nextMonth") {
    var d = Number(el.getAttribute("data-d")) || 0;
    state.month = shiftMonth(state.month, d);
    state.summary = null;
    render();
    if (state.tab === "summary") loadSummary();
    else if (state.tab === "list") loadList();
    return;
  }
  if (act === "tab") {
    state.tab = el.getAttribute("data-tab");
    render();
    if (state.tab === "summary") loadSummary();
    else if (state.tab === "debts") {
      if (!state.btc) loadBtcPrice();
      // ALWAYS refresh on entering the Debts tab. The tab needs the enriched
      // list (payoff, clearsOn, whatIf); the Settings form only ever holds the
      // raw draft rows. Rendering the tab from a settings-shaped list produced
      // the exact ""No debt" but money is stacking" screen: the leak card drew
      // fine, every debt card was skipped for lacking a payoff, and the empty
      // state never appeared because the array was not empty.
      // One request per entry. This used to fire twice when there were no
      // cached debts - the unconditional call plus the conditional one below.
      loadDebts();
    }
    else loadList();
    return;
  }
  if (act === "pick") {
    // Clear the previous failure so a retake starts clean and a stale message
    // cannot sit under the new proof photo.
    state.scanError = null;
    var fi = document.getElementById("file");
    if (fi) fi.click();
    return;
  }
  if (act === "scan") { beginScan(); return; }
  if (act === "discardDraft") {
    if (window.confirm("Discard this receipt?")) { state.draft = null; closeSheet(); }
    return;
  }
  if (act === "swap") { return; }
  if (act === "more") {
    // Server caps at 500; this just raises the ask until everything is shown.
    // Ignore the repeat during the in-flight request. run() raises state.busy
    // and render() repaints the button as disabled, but a tap landing in the
    // same frame still reaches this handler -- the first tap appeared to do
    // nothing and the second one worked, because the render had landed by then.
    if (state.busy) return;
    state.listLimit = Math.min((state.listLimit || 50) + 50, 500);
    render();
    loadList();
    return;
  }
  if (act === "edit") {
    openEdit(findRow(el) || el.getAttribute("data-rownum"));
    return;
  }
  if (act === "del") {
    confirmDelete(findRow(el) || el.getAttribute("data-rownum"));
    return;
  }
  if (act === "closeEdit") {
    state.editRow = null;
    render();
    return;
  }
});

document.addEventListener("change", function(e){
  if (e.target && e.target.id === "file") handleFile(e.target.files);
});

// Live totals in the settings form. Without this the figures only refreshed on
// a full render, so typing a new number left the old one on screen.
document.addEventListener("input", function(e){
  if (!e.target || !e.target.id) return;
  if (e.target.id.indexOf("s_") !== 0) return;
  updateSettingTotals();
});

function fieldNum(id){
  var el = document.getElementById(id);
  if (!el) return 0;
  var v = parseFloat(el.value);
  return isNaN(v) ? 0 : v;
}

function updateSettingTotals(){
  var sal = fieldNum("s_income_salary");
  var wife = fieldNum("s_income_wife");
  var fixed = fieldNum("s_fixed_total");
  var variable = fieldNum("s_variable_budget");
  var totalIncome = sal + wife;
  // The derived 'free' figure is gone: Settings only echoes the one sum the
  // user's own inputs produce, and nothing the app calculates behind it.
  var incEl = document.getElementById("sum_income");
  if (incEl) incEl.textContent = money(totalIncome);
}

document.addEventListener("submit", function(e){
  var f = e.target;
  if (!f || !f.getAttribute) return;
  var a = f.getAttribute("data-act");
  if (a === "save") { e.preventDefault(); doSave(); return; }
  if (a === "update") { e.preventDefault(); doUpdate(); return; }
  if (a === "savedraft") { e.preventDefault(); doSaveDraft(); return; }
  if (a === "savesettings") { e.preventDefault(); doSaveSettings(); return; }
  if (a === "savepay") { e.preventDefault(); savePayFromForm(); return; }
  // Missing until v55: the install form's Connect button and Enter key were
  // never prevented, so the browser did a native submit and blanked the iframe
  // before saveInstall could finish.
  if (a === "saveinstall") { e.preventDefault(); saveInstallFromForm(); return; }
});

function handleFile(files){
  if (!files || !files.length) return;
  var file = files[0];
  var reader = new FileReader();
  reader.onload = function(){
    var img = new Image();
    img.onload = function(){
      try {
        var max = 1024;
        var scale = Math.min(1, max / Math.max(img.width, img.height));
        var cv = document.createElement("canvas");
        cv.width = Math.round(img.width * scale);
        cv.height = Math.round(img.height * scale);
        cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
        state.photo = cv.toDataURL("image/jpeg", 0.8);
        // 2 steps, not 3: picking the photo starts the scan immediately.
        render();
        beginScan();
      } catch (err) {
        toast("Could not read that image", true);
      }
    };
    img.onerror = function(){ toast("Could not open that image", true); };
    img.src = reader.result;
  };
  reader.readAsDataURL(file);
}

function beginScan(){
  // Guard on scanning, not busy: busy is now a counter shared with every
  // background load, so a summary or price still in flight covered taps on
  // Take photo and the button did nothing. The photo button is already
  // disabled while scanning, so this only blocks a double-scan.
  if (!state.photo || state.scanning) return;
  var b64 = state.photo.split(",")[1];
  if (!b64) { toast("Bad image", true); return; }
  state.scanning = true;
  state.error = null;
  state.scanError = null;
  state.draft = null;
  render();
  run("extractReceipt", { image: b64, mimeType: "image/jpeg" }, function(res){
    state.scanning = false;
    if (!res || res.ok === false) {
      // Stay on the add screen with the manual fields right there. The reason
      // must be shown INSIDE the sheet - state.error renders behind it.
      state.scanError = (res && res.error) ? res.error : "Could not read that receipt.";
      render();
      return;
    }
    state.draft = res.draft;
    state.scanError = null;
    state.error = null;
    render();
  });
}

function doSaveDraft(){
  var amtEl = document.getElementById("d_amount");
  var catEl = document.getElementById("d_category");
  var venEl = document.getElementById("d_vendor");
  var payEl = document.getElementById("d_payment");
  var datEl = document.getElementById("d_date");

  var amount = amtEl ? parseFloat(amtEl.value) : NaN;
  var vendor = venEl ? String(venEl.value).trim() : "";

  if (isNaN(amount) || amount <= 0) { toast("Enter an amount", true); return; }
  if (!vendor) { toast("Enter a vendor", true); return; }

  state.error = null;
  render();

  run("addReceipt", {
    amount: amount,
    vendor: vendor,
    category: catEl ? catEl.value : "Groceries & Household",
    payment: payEl ? payEl.value : "Cash",
    date: (datEl && datEl.value) ? datEl.value : todayISO(),
    currency: (state.draft && state.draft.currency) ? state.draft.currency : CUR,
    confidence: (state.draft && typeof state.draft.confidence === "number") ? state.draft.confidence : null
  }, function(res){
    if (!res || res.ok === false) {
      state.error = "Save failed: " + ((res && res.error) ? res.error : "unknown");
      render();
      return;
    }
    state.draft = null;
    closeSheet();
    if (state.tab === "summary") loadSummary(); else loadList();
    toast("Saved - " + money(amount));
  });
}

function doSave(){
  var amtEl = document.getElementById("f_amount");
  var venEl = document.getElementById("f_vendor");
  var catEl = document.getElementById("f_category");
  var payEl = document.getElementById("f_payment");
  var datEl = document.getElementById("f_date");

  var amount = amtEl ? parseFloat(amtEl.value) : NaN;
  var vendor = venEl ? String(venEl.value).trim() : "";

  if (isNaN(amount) || amount <= 0) { toast("Enter an amount", true); return; }
  if (!vendor) { toast("Enter a vendor", true); return; }

  state.error = null;
  render();

  run("addManual", {
    amount: amount,
    vendor: vendor,
    category: catEl ? catEl.value : "Groceries & Household",
    payment: payEl ? payEl.value : "Cash",
    date: (datEl && datEl.value) ? datEl.value : todayISO(),
    // Missing until v55: without this the server fell back to 'BND' and every
    // manually typed expense was written as BND whatever the user's currency.
    currency: CUR
  }, function(res){
    if (!res || res.ok === false) {
      state.error = "Save failed: " + ((res && res.error) ? res.error : "unknown");
      render();
      return;
    }
    state.sheet = false;
    if (state.tab === "summary") loadSummary(); else loadList();
    toast("Saved - " + money(amount));
  });
}


/* ---------- swipe gesture ---------- */
(function(){
  var startX = 0, startY = 0, active = null, moved = false;

  document.addEventListener("touchstart", function(e){
    var t = e.target;
    while (t && t !== document && !(t.classList && t.classList.contains("sw"))) t = t.parentNode;
    if (!t || t === document) { active = null; return; }
    active = t;
    moved = false;
    startX = e.touches[0].clientX;
    startY = e.touches[0].clientY;
  }, { passive: true });

  document.addEventListener("touchmove", function(e){
    if (!active) return;
    var dx = e.touches[0].clientX - startX;
    var dy = e.touches[0].clientY - startY;
    if (Math.abs(dy) > Math.abs(dx)) { active = null; return; }
    if (Math.abs(dx) > 8) moved = true;
  }, { passive: true });

  document.addEventListener("touchend", function(e){
    if (!active || !moved) { active = null; return; }
    var dx = e.changedTouches[0].clientX - startX;
    var num = Number(active.getAttribute("data-rownum"));
    if (dx < -40) { state.swipeOpen = num; render(); }
    else if (dx > 40) { state.swipeOpen = null; render(); }
    active = null;
    moved = false;
  }, { passive: true });
})();

/* ---------- boot ---------- */

loadDenom();
try {
  // A bound copy is already configured; a standalone copy may not be. Ask the
  // server before reading data, or a fresh install errors instead of offering
  // the connect screen.
  run("getInstallState", {}, function(res){
    if (res && res.ok && !res.configured) {
      state.needsInstall = true;
      state.needsWelcome = false;
      state.onboardChecked = true;
      render();
      return;
    }
    loadSummary();
  });
} catch (e) {
  state.error = "Boot error: " + ((e && e.message) ? e.message : String(e));
  render();
}

})();
</script>
</body>
</html>
`;
