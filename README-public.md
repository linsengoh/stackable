# Stackable

**How much bitcoin can you actually stack this month?**

Most budgeting apps tell you what you spent. Stackable answers the question you actually have: after rent, after the card payments, after the groceries — how much is genuinely left to put into bitcoin this month? And it answers it in sats.

One Google Sheet. One script. No account, no subscription, no company holding your data. Free.

---

## What it does

- **Log an expense in seconds.** Photograph a receipt, or type three fields.
- **See your stackable figure** in dollars and in sats, recalculated the moment you spend.
- **See what your debt costs you** — in interest per month, and in sats — and what paying it down faster would actually change.
- **Your data stays in a spreadsheet you own.** Open it, read it, edit it, export it, delete it.

## What it doesn't do

- No account, no sign-up beyond your Google account.
- No bank connection. You photograph or type; it records.
- No telemetry, no analytics, no tracking.
- **No investment advice.** It prices the choice. It does not tell you what to buy.

---

## Install (about 5 minutes)

You need one Google account. Everything below is free.

### 1. Copy the code

Open **[code.gs](https://github.com/linsengoh/stackable/blob/main/code.gs)** and press the **copy** button (the little clipboard icon, top right of the code box).

If that link is broken, open the repo and click `code.gs`.

### 2. Create the script

1. Go to **[script.new](https://script.new)**
2. Name the project `Stackable`
3. Delete the sample code in the editor
4. Paste what you copied
5. Press **Save** (⌘S / Ctrl+S)

### 3. Deploy it

1. **Deploy → New deployment**
2. Click the gear icon next to "Select type" → **Web app**
3. **Description:** `v1`
4. **Execute as:** **Me**
5. **Who has access:** **Only myself**
6. **Deploy**

Google will warn you the app isn't verified. That's expected — it's your own script, not someone else's:

> **Advanced** → **Go to Stackable (unsafe)**

It will ask for permission to access your spreadsheets and Google Drive. That's how it creates and reads your ledger.

### Why "Only myself" matters

A web app set to **Anyone** runs as you, reading your sheet, for anyone who has the URL. A shared link would be a shared bank statement. Keep it on **Only myself**. If you ever change it for testing, change it back.

### 4. Open it and connect

You'll get a URL like:

```
https://script.google.com/macros/s/AKfycb.../exec
```

Open it. The app offers to create a spreadsheet called **Stackable** in your own Drive — tap Connect. You don't need to make one or paste any IDs.

Then **add it to your home screen** so it opens like an app:
- **iPhone:** Share → Add to Home Screen
- **Android:** Menu → Add to Home screen

### 5. Fill in your numbers

Settings asks for five things:

- **Salary, take-home** — what actually lands in your account, after tax and deductions
- **Partner / other income**
- **Fixed monthly costs** — rent or mortgage, car payment, childcare or tuition, utilities, internet, phone, insurance
- **Monthly spending target** — leave it at 0 and the app works it out from your own history once two complete months exist
- **Currency** — defaults to USD

Until you fill these in, the app shows "Spent this month" instead of a stackable figure. That's not a bug; it has nothing to calculate from yet.

**The loan rule, in one place.** Fixed costs is a living-costs bucket — a mortgage or car payment belongs there, whole. The one exception is **credit card payments**: those go in **Debts**, because that's where the interest gets priced. Enter a card payment in both places and your stackable figure drops twice.

---

## Optional: receipt scanning

Scanning is the one feature that needs an API key. **The app works fully without it** — you just type expenses instead of photographing them. Typing takes about twenty seconds; a photo takes three.

To turn it on:

1. Go to **[aistudio.google.com](https://aistudio.google.com)**, signed in with the same Google account
2. **Get API key** → **Create API key**
3. Copy it. No spaces, no line breaks
4. In Stackable: **Settings → Receipt scanning**, paste, Save

Keys come in two shapes — older ones start `AIza...`, newer ones `AQ.Ab...`. Both work.

**Honest caveat:** Google has been tightening the free tier. Some new accounts are asked to enable billing on a Cloud project before a key works, even inside the free allowance. If Google asks you for a card, that's between you and Google. A few receipts a day should cost nothing, but I can't promise your account won't be asked.

---

## Using it

### The stackable figure

Income, minus fixed costs, minus your card payments, minus what you've spent, minus what the rest of the month will cost at your current pace.

Every term comes from something you entered or something you actually spent. Nothing is a placeholder, and nothing is optimistic. Spend more today and the number moves today.

### Logging an expense

Tap **+**. Photograph the receipt — the app reads vendor, amount, date and category and shows you a draft to check before saving — or type it.

Green entries are money coming back (Reimbursed). They reduce what the month cost, rather than counting as spending.

### Categories

The app files every expense into one of these. Use these exact names if you're pasting rows into the sheet:

`Groceries & Household` · `Eating Out` · `Transport` · `Health` · `Fun` · `Kids` · `Gifts` · `Sats` · `One-off` · `Reimbursed`

- **Sats** — bitcoin you bought. Doesn't count as spending; it's the point.
- **One-off** — a repair, a holiday booking. Real money, but not a daily rate, so it's tracked separately from your pace.
- **Reimbursed** — money someone paid you back. Enter as a negative amount.

Older names like `groceries` still fold into `Groceries & Household`, so old history never orphans itself.

### Debts

Each card shows **facts**: balance, rate, what you pay, when it clears at that payment.

The card at the top shows the **decision**: what the debt costs you in interest every month, and what clearing it would free up.

The **If I paid** chips show what a higher payment actually changes — months saved, interest saved. If a chip shows no earlier month, no higher payment buys you one; it still saves interest, and the line says so.

### Your data

**Settings → Open my spreadsheet** takes you to the Sheet.

Tabs: **Expenses**, **Settings**, **Debts**, **DebtHistory**, **DebtPayments**.

Deleting the app does not delete your data. Deleting the spreadsheet does.

---

## Use it as a spreadsheet

This is the part most people miss, and it's why the ledger is a plain table instead of a black box.

The **Expenses** tab has fixed columns and a stable id per row:

```
date | vendor | amount | currency | category | payment | photo | confidence | logged_at | id | updated_at
```

That means you can **paste a bank CSV straight into the sheet** and the app picks it up immediately. Summary, pace, by-category and the budget suggestion all recompute — including months from before you installed anything.

Three things make pasted data behave:

- **Blank currency means your ledger currency.** Leave it empty and the row counts as whatever is in Settings. Put in a different code (`EUR` on a trip) and it converts at the live rate before it touches any total.
- **Old category names fold silently.** `groceries` becomes `Groceries & Household`.
- **Columns are matched by position.** Paste in the same order and the app treats those rows as real expenses.

Useful for: backfilling three months from your bank export so the target calculation has something to work with on day one; bulk-adding a holiday's receipts; fixing a mistyped category across fifty rows at once.

Same caveat as any spreadsheet: a wrong date files a row in the wrong month, and text in `amount` gets skipped rather than guessed at.

---

## Troubleshooting

**The app won't load**
Check **Deploy → Manage deployments**. It must say **Execute as: Me** and **Who has access: Only myself**, and you must be signed into that same Google account. The URL only works for you.

**I changed the code and nothing happened**
Edit the existing deployment and choose **New version**. Creating a new deployment makes a new URL and you'd have to re-add it to your home screen.

**Scanning says a Gemini key is needed**
Settings → Receipt scanning, paste it, Save. Typing an expense never needs a key.

**Scanning fails with 403 or 404**
The model name is retired or wrong. Clear the Gemini model field — the app then uses its built-in default, which is what most people want.

**An expense is missing**
Check its date in the sheet. Wrong year, wrong month. The list shows one month at a time; use the arrows beside the month name.

**Summary and Expenses disagree**
They shouldn't — both convert foreign rows at the same rate. If they differ, one row has an unrecognised currency code. Check the `currency` column for a typo.

**Photos scan slowly**
Free-tier Gemini can take 20–40 seconds on a cold start. Better light, closer framing. Or just type it.

---

## What it costs

- Apps Script: free
- Google Sheets: free
- Bitcoin price and exchange rates: free public APIs
- Gemini: free tier, roughly 20 receipts a day
- **Total: nothing**, unless Google asks you to enable billing on your own account

---

## Privacy, plainly

Your spreadsheet is yours. The only outbound requests are bitcoin price and exchange-rate lookups, plus Google's Gemini API when you scan a receipt. No analytics, no telemetry, no copy of your data anywhere but your own Drive.

Everyone who uses Stackable deploys their own copy, with their own sheet and their own URL. Nobody shares an instance. Nobody can reach anyone else's data.

---

## Feedback

This is a v1, built by one person, released free. If something is broken, confusing, or missing — tell me. Open an issue in this repo.

If it's genuinely saved you money, there's a [donation link](https://github.com/linsengoh/stackable#support). No obligation, ever. The app is and stays free.

---

## Licence

Do what you like with it. Copy it, change it, share it.
