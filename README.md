# 📊 Ultimate Habit & Sleep Tracker (2026 – 2030)

Welcome to the **Ultimate Habit & Sleep Tracker**! This workbook is an all-in-one monthly habit, goal, and sleep tracking system spanning **52 monthly sheets from September 2026 through December 2030**.

The workbook is located at:
📁 **[`Habit_Tracker_2026_2030.xlsx`](./Habit_Tracker_2026_2030.xlsx)**

---

## 📑 Table of Contents
1. [Overview & Key Highlights](#-overview--key-highlights)
2. [Workbook Layout & Anatomy](#-workbook-layout--anatomy)
3. [Step-by-Step Usage Guide](#-step-by-step-usage-guide)
   - [1. Monthly Setup (Targets & Notes)](#1-monthly-setup-targets--notes)
   - [2. Defining Your Protocols (Habits)](#2-defining-your-protocols-habits)
   - [3. Daily Check-Ins (Tick & Cross Flow)](#3-daily-check-ins-tick--cross-flow)
   - [4. Understanding Scores & Percentages](#4-understanding-scores--percentages)
   - [5. Stakes: Punishments & Rewards](#5-stakes-punishments--rewards)
   - [6. Sleep Tracking & Live Chart](#6-sleep-tracking--live-chart)
4. [Color Coding & Visual Aesthetics](#-color-coding--visual-aesthetics)
5. [Frequently Asked Questions (FAQ)](#-frequently-asked-questions-faq)
6. [Developer & Regeneration Guide](#-developer--regeneration-guide)

---

## 🌟 Overview & Key Highlights

- **52 Complete Monthly Sheets**: Every single month from **September 2026** to **December 2030**.
- **Dynamic Calendar Logic**: Exact number of days for 30-day months, 31-day months, and leap years (e.g., February 2028 with 29 days).
- **20 Protocol Slots**: Room for up to 20 daily habits with customizable targets.
- **Smart Auto-Discard Logic**: Incomplete/unselected habit rows do not penalize your daily score.
- **Dynamic Conditional Formatting**: Cells automatically light up in green for completed tasks (`✓`) and red for missed tasks (`✗`).
- **Done vs Target Tracking**: View exact counts (`24/30`) alongside percentage completions.
- **Dedicated Sleep Tracker & Live Chart**: Track sleep hours (1–24) with an interactive line chart.
- **Accountability Stakes**: Clear sections for monthly reward and penalty commitments.

---

## 🗺️ Workbook Layout & Anatomy

Each sheet is structured vertically in a clean, logical sequence:

```text
┌─────────────────────────────────────────────────────────────────────────────┐
│ 1. HEADER: Sheet Title (Month & Year) + Rotating Motivational Quote        │
├─────────────────────────────────────────────────────────────────────────────┤
│ 2. TOP SECTION:                                                             │
│    • MONTHLY TARGETS: 8 major monthly milestone boxes (in 2 columns)        │
│    • NOTES CONTAINER: Soft yellow memo space for reflections & strategies   │
├─────────────────────────────────────────────────────────────────────────────┤
│ 3. MAIN PROTOCOLS TABLE (20 Rows):                                          │
│    • S.No. (1 to 20) | Protocol Name | Target (1-31 dropdown)               │
│    • Days (1 to 30/31): In-cell dropdowns for '✓' or '✗'                   │
│    • Success %: Automated completion rate for the habit                     │
│    • Done / Target: Exact count ratio (e.g., 25/30)                         │
├─────────────────────────────────────────────────────────────────────────────┤
│ 4. DAILY PERFORMANCE ROWS:                                                  │
│    • DAILY TOTAL SCORE: Evaluated ratio for that day (e.g., 8/10)           │
│    • DAILY SUCCESS %: Formatted percentage (e.g., 80.0%)                    │
├─────────────────────────────────────────────────────────────────────────────┤
│ 5. ACCOUNTABILITY STAKES:                                                   │
│    • PUNISHMENT: Stakes for failing to meet monthly goals                   │
│    • REWARD: Celebration incentive for achieving your goals                 │
├─────────────────────────────────────────────────────────────────────────────┤
│ 6. SLEEP TRACKING TABLE:                                                    │
│    • Night Indigo header with 1–24 hours dropdown per day                   │
├─────────────────────────────────────────────────────────────────────────────┤
│ 7. LIVE SLEEP TREND CHART:                                                  │
│    • Interactive line graph displaying Sleep Hours vs. Day of the Month     │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 🚀 Step-by-Step Usage Guide

### 1. Monthly Setup (Targets & Notes)
At the beginning of each month:
1. Navigate to the month tab (e.g., `Sep 2026`).
2. Under **MONTHLY TARGETS** (Rows 6–9), write down up to 8 core priorities or milestones for the month across the two columns (`1.` through `8.`).
3. Under **NOTES:** (Row 12), jot down focus areas, mindset cues, or weekly gameplans.

### 2. Defining Your Protocols (Habits)
In the main protocol table (Rows 18–37):
1. **Protocol Name (Column B)**: Enter your habit or routine (e.g., *Morning Run*, *Read 20 Pages*, *Meditation*, *No Sugar*).
2. **Target (Column C)**: Click the cell to open the dropdown and select your goal for the number of days you aim to complete this habit this month (from `1` to `31`).

### 3. Daily Check-Ins (Tick & Cross Flow)
For each day:
1. Click the cell corresponding to your habit and current day (Columns D through AH/AI/AJ).
2. Use the dropdown arrow to select:
   - `✓` (Checkmark): You completed the protocol. The cell turns **Soft Green**.
   - `✗` (Cross): You missed or skipped the protocol. The cell turns **Soft Red**.
   - *Blank / Nothing*: The habit was not tracked or not active yet. The cell remains its neutral pastel color.

### 4. Understanding Scores & Percentages

#### How the Auto-Discard Logic Works
Only active items are evaluated. If you have only 8 active habits marked with `✓` or `✗`, the denominator is **8**, not 20:
$$\text{Daily Total Score} = \frac{\text{Count of } ✓}{\text{Count of } ✓ + \text{Count of } ✗}$$

- **Example 1**: You evaluate 10 habits today. 8 are `✓` and 2 are `✗`.
  - **DAILY TOTAL SCORE (Row 38)** displays: `8/10`
  - **DAILY SUCCESS % (Row 39)** displays: `80.0%`
- **Example 2**: You only track 5 habits on a busy day. 5 are `✓` and 0 are `✗`.
  - **DAILY TOTAL SCORE** displays: `5/5`
  - **DAILY SUCCESS %** displays: `100.0%`
- **Unmarked cells**: Any habit left blank is completely excluded so your score is never unfairly penalized.

#### Summary Columns on the Right
- **`Success %`**: Shows $\frac{\text{Total Days Completed}}{\text{Target}}$ as a percentage.
- **`Done / Target`**: Displays the exact fractional score (e.g. `24/30`).
- **Monthly Totals**: The bottom-right summary cell under `Done / Target` shows your cumulative monthly score vs. total targets (e.g., `185/220`).

### 5. Stakes: Punishments & Rewards
Accountability creates consistency:
- **PUNISHMENT (Row 41)**: Enter what you must do if you fail to reach your monthly commitment (e.g., *Donate $50*, *Cold shower for a week*).
- **REWARD (Row 42)**: Enter your celebration reward for hitting your targets (e.g., *Weekend trip*, *Buy new running shoes*).

### 6. Sleep Tracking & Live Chart
1. In the **Sleep Tracking** section (Row 47), click on the day cell.
2. Select your hours of sleep from the dropdown (`1` to `24`).
3. The in-cell number will be clearly visible.
4. The **Daily Sleep Trend** line chart directly below (cell `A49`) will dynamically plot your sleep curve across the month.

---

## 🎨 Color Coding & Visual Aesthetics

| Element | Color Palette | Hex Code | Description |
| :--- | :--- | :--- | :--- |
| **Habit Group 1** | Sky Blue | `#E0F2FE` | Clean, refreshing start (Rows 18–21) |
| **Habit Group 2** | Soft Lavender | `#F3E8FF` | Calming mental focus (Rows 22–25) |
| **Habit Group 3** | Warm Sand | `#FEF3C7` | Energizing amber pastel (Rows 26–29) |
| **Habit Group 4** | Cool Slate | `#F1F5F9` | Neutral modern slate (Rows 30–33) |
| **Habit Group 5** | Periwinkle | `#EEF2FF` | Harmonious twilight blue (Rows 34–37) |
| **Success Cell (`✓`)** | Mint Green | `#DCFCE7` | High-contrast accomplishment indicator |
| **Missed Cell (`✗`)** | Soft Coral Red | `#FEE2E2` | Gentle accountability alert |
| **Sleep Header** | Deep Night Indigo | `#1E1B4B` | Elegant night-sky theme |
| **Sleep Day Cells** | Bedtime Blue | `#EEF2FF` | High contrast for dropdown numbers |
| **Notes Box** | Pale Memo Yellow | `#FFFFE6` | Familiar sticky-note aesthetic |

---

## ❓ Frequently Asked Questions (FAQ)

### Q: Can I use this in Google Sheets or Microsoft Excel?
**A**: Yes! The workbook is standard `.xlsx` (OpenXML). It opens natively in **Microsoft Excel (Desktop, Web, & Mobile)**, **Google Sheets**, and **LibreOffice Calc**. Dropdowns and formulas are standard Excel formulas (`COUNTIF`, `SUM`, `IFERROR`).

### Q: What if I only want to track 6 habits instead of 20?
**A**: Simply fill in 6 rows and leave the other 14 rows blank. Because of the built-in auto-discard logic, blank rows will never count toward your daily score or percentages.

### Q: Can I add a new habit in the middle of the month?
**A**: Absolutely. Type the new habit name in an empty protocol row, choose its target, and start selecting `✓` or `✗` from today onwards. Past unselected days for that habit will not affect your earlier scores.

### Q: Why do the column widths look slightly wider than normal?
**A**: Day column widths are calibrated to `4.8` so that Excel's dropdown arrow button does not hide or truncate two-digit numbers (like `10`, `12`, or `24`).

---

## 💻 Developer & Regeneration Guide

If you ever need to customize the generator or add more years beyond 2030:

- **Generator Script**: [`create_excel_tracker.py`](./create_excel_tracker.py)
  - Contains full layout logic, styles, conditional formatting, and formulas.
  - To regenerate the workbook:
    ```bash
    python create_excel_tracker.py
    ```
- **Automated Verification Script**: [`verify_excel_tracker.py`](./verify_excel_tracker.py)
  - Validates all 52 sheets, headers, colors, conditional formatting, and formulas:
    ```bash
    python verify_excel_tracker.py
    ```

---

*Enjoy tracking your habits and building momentum every single day from 2026 through 2030!* 🚀
