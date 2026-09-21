# 📊 Ultimate Habit, Sleep & Analytics Tracker (2026 – 2030)

Welcome to the **Ultimate Habit, Sleep & Analytics Tracker**! This workbook is an all-in-one monthly habit execution, domain analytics, and sleep tracking system spanning **52 monthly sheets from September 2026 through December 2030**.

The workbook is located at:
📁 **[`Habit_Tracker_2026_2030.xlsx`](./Habit_Tracker_2026_2030.xlsx)**

---

## 📑 Table of Contents
1. [Overview & Highlights](#-overview--highlights)
2. [Visual Architecture & Anatomy](#-visual-architecture--anatomy)
3. [Step-by-Step Usage Guide](#-step-by-step-usage-guide)
   - [1. Monthly Setup (Targets & Notes)](#1-monthly-setup-targets--notes)
   - [2. Protocols & Domain Categories](#2-protocols--domain-categories)
   - [3. Daily Tracking Flow & "Today" Highlighting](#3-daily-tracking-flow--today-highlighting)
   - [4. Scoring, Formulas & Auto-Discard](#4-scoring-formulas--auto-discard)
   - [5. Sleep Tracking & Dynamic Line Chart](#5-sleep-tracking--dynamic-line-chart)
4. [End-of-Month Analytics Suite](#-end-of-month-analytics-suite)
   - [Executive KPI Cards](#1-executive-monthly-kpi-cards)
   - [Weekly Slump Detector & Chart](#2-weekly-performance-slump-detector--chart)
   - [Category Domain Breakdown & Chart](#3-life-domain--category-breakdown--chart)
   - [Sleep vs. Productivity Correlation Analysis](#4-sleep--productivity-correlation-analysis)
   - [Annual & Quarterly Review Data Bridge](#5-standardized-annual--quarterly-data-bridge)
   - [Monthly Retrospective Memo Box](#6-monthly-retrospective--action-plan)
5. [Color Palette Reference](#-color-palette-reference)
6. [Frequently Asked Questions (FAQ)](#-frequently-asked-questions-faq)
7. [Developer & Verification Guide](#-developer--verification-guide)

---

## 🌟 Overview & Highlights

- **52 Complete Monthly Sheets**: Every single month from **September 2026 through December 2030**.
- **Dynamic Calendar Logic**: Exact number of days for 30-day months, 31-day months, and leap years (e.g., February 2028 with 29 days).
- **20 Protocol Slots**: Room for up to 20 habits with category tagging and monthly targets (1–31).
- **Life Domain Tagging**: Classify habits across 6 domains (`Health & Fitness`, `Mindset & Learning`, `Career & Focus`, `Personal & Home`, `Finance & Wealth`, `Other`).
- **Dynamic "Today" Highlighter**: The current day's column lights up automatically in warm gold based on `=TODAY()`.
- **Smart Auto-Discard Logic**: Incomplete/unselected habit rows are auto-discarded from daily scores.
- **Dynamic Conditional Formatting**: Cells automatically turn **Soft Green** for completed (`✓`) and **Soft Red** for missed (`✗`).
- **Done vs Target Tracking**: View exact counts (`24/30`) alongside percentage completions.
- **Dedicated Sleep Tracking & Live Chart**: Track sleep hours (1–24) with a synchronized trend curve.
- **3 Interactive Charts per Sheet**:
  1. *Daily Sleep Trend Line Chart*
  2. *Weekly Momentum Trend Column Chart*
  3. *Performance by Domain Bar Chart*
- **Sleep & Productivity Correlation Engine**: Measures your quantified productivity boost on well-rested days (≥7 hrs) vs. deficit days (<7 hrs).
- **Quarterly & Annual Data Bridge**: Standardized export row allowing seamless multi-month aggregation.
- **Monthly Retrospective Coaching Box**: End-of-month reflection on Wins, Friction Points, and Next Month's Action Plan.

---

## 🗺️ Visual Architecture & Anatomy

Each monthly sheet follows an intuitive top-to-bottom layout:

```text
┌─────────────────────────────────────────────────────────────────────────────┐
│ 1. HEADER: Title (Month & Year) + Rotating Motivational Quote               │
├─────────────────────────────────────────────────────────────────────────────┤
│ 2. TOP SECTION:                                                             │
│    • MONTHLY TARGETS: 8 milestone goal boxes across 2 columns               │
│    • NOTES CONTAINER: Soft yellow memo space for strategies & reflections   │
├─────────────────────────────────────────────────────────────────────────────┤
│ 3. MAIN PROTOCOLS TABLE (20 Rows):                                          │
│    • S.No. (1 to 20) | Protocols | Category (Dropdown) | Target (Dropdown)  │
│    • Days (1 to 30/31): In-cell dropdowns for '✓' or '✗'                   │
│    • Success %: Automated completion rate for the habit                     │
│    • Done / Target: Exact count ratio (e.g., 25/30)                         │
├─────────────────────────────────────────────────────────────────────────────┤
│ 4. DAILY PERFORMANCE ROWS:                                                  │
│    • DAILY TOTAL SCORE: Evaluated ratio for that day (e.g., 8/10)           │
│    • DAILY SUCCESS %: Formatted percentage (e.g., 80.0%)                    │
├─────────────────────────────────────────────────────────────────────────────┤
│ 5. ACCOUNTABILITY STAKES:                                                   │
│    • PUNISHMENT: Stakes for failing to meet monthly goals (<80%)            │
│    • REWARD: Celebration incentive for achieving your goals (>90%)          │
├─────────────────────────────────────────────────────────────────────────────┤
│ 6. SLEEP TRACKING SECTION & DYNAMIC CHART:                                  │
│    • Twilight-themed Sleep Hours row (1–24 dropdown)                        │
│    • Interactive Daily Sleep Trend Line Chart (Hours vs. Day)               │
├─────────────────────────────────────────────────────────────────────────────┤
│ 7. EXECUTIVE MONTHLY KPI CARDS:                                             │
│    [ Monthly Success Rate ] [ Active Protocols ] [ Perfect Days ] [ Avg Sleep]
├─────────────────────────────────────────────────────────────────────────────┤
│ 8. END-OF-MONTH ANALYTICS & CHARTS:                                         │
│    • Left: Weekly Slump Detector Table + Weekly Momentum Column Chart       │
│    • Right: Life Domain Breakdown Table + Domain Performance Bar Chart      │
│    • Sleep & Productivity Correlation Engine (Impact of ≥7h vs <7h sleep)  │
│    • Standardized Annual / Quarterly Review Data Bridge (Row 92)            │
│    • Monthly Retrospective Box (Wins, Friction Points, Action Plan)         │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 🚀 Step-by-Step Usage Guide

### 1. Monthly Setup (Targets & Notes)
At the beginning of each month:
1. Navigate to the month tab (e.g., `Sep 2026`).
2. Under **MONTHLY TARGETS** (Rows 6–9), write down up to 8 core priorities or milestones for the month across the two columns (`1.` through `8.`).
3. Under **NOTES:** (Row 12), write your focus areas, mindset cues, or weekly gameplans in the pale yellow memo container.

### 2. Protocols & Domain Categories
In the main protocol table (Rows 18–37):
1. **Protocols (Column B)**: Enter your habit or routine (e.g., *Gym Session*, *Deep Work 2h*, *Read 20 Pages*, *Meditation*).
2. **Category (Column C)**: Select the life domain from the dropdown:
   - `Health & Fitness`
   - `Mindset & Learning`
   - `Career & Focus`
   - `Personal & Home`
   - `Finance & Wealth`
   - `Other`
3. **Target (Column D)**: Select your monthly target days from the dropdown (`1` to `31`).

### 3. Daily Tracking Flow & "Today" Highlighting
- **"Today" Column Highlight**: When opening the sheet for the current month, the current day column header automatically lights up in warm gold (`#FEF08A`) so your eyes immediately locate today's column!
- For each day:
  1. Click the cell corresponding to your habit and today's day (Columns E through AH/AI/AJ).
  2. Use the dropdown arrow to select:
     - `✓` (Checkmark): Task completed. The cell turns **Soft Green** (`#DCFCE7`).
     - `✗` (Cross): Task missed. The cell turns **Soft Red** (`#FEE2E2`).
     - *Blank*: Unmarked / inactive habit. The cell retains its gentle pastel group color.

### 4. Scoring, Formulas & Auto-Discard

#### Smart Auto-Discard Logic
Unmarked habits are excluded from that day's score so new habits added mid-month never damage your record:
$$\text{Daily Total Score} = \frac{\text{Count of } ✓}{\text{Count of } ✓ + \text{Count of } ✗}$$

- **Example**: If 10 habits are marked today with 8 `✓` and 2 `✗`:
  - **DAILY TOTAL SCORE (Row 38)** displays: `8/10`
  - **DAILY SUCCESS % (Row 39)** displays: `80.0%`
- **Summary Columns on the Right**:
  - `Success %`: $\frac{\text{Days Done}}{\text{Target}}$ as a percentage.
  - `Done / Target`: Displays the exact count (e.g., `24/30`).
  - Monthly Totals (bottom-right cell) shows total checks done vs. total targets (e.g., `185/220`).

### 5. Sleep Tracking & Dynamic Line Chart
1. In the **Sleep Tracking** section (Row 47), select your sleep hours from the dropdown (`1` to `24`).
2. Numbers appear centered with high-contrast night-indigo font.
3. The **Daily Sleep Trend** chart (Row 49) automatically updates its blue curve across the month.

---

## 📈 End-of-Month Analytics Suite

Scroll below the sleep chart to access the automated monthly analytics:

### 1. Executive Monthly KPI Cards
Four color-coded executive stat cards summarizing monthly performance:
- 🟢 **Monthly Success Rate**: Overall completion percentage against targets.
- 🔵 **Active Protocols**: Total count of active habits out of 20 slots.
- 🟡 **Perfect Days (100%)**: Number of days where all marked habits were achieved.
- 🟣 **Avg Sleep Duration**: Average sleep hours per night across logged days.

### 2. Weekly Performance (Slump Detector) & Chart
- Breaks down your month into **Week 1 (Days 1–7)**, **Week 2 (Days 8–14)**, **Week 3 (Days 15–21)**, **Week 4 (Days 22–28)**, and **Week 5 (Days 29–end)**.
- Compares Checks Done, Evaluated Count, and Success Rate for each week.
- Displays the **Weekly Momentum Trend** column chart to immediately identify mid-month slumps or weekend drop-offs.

### 3. Life Domain / Category Breakdown & Chart
- Aggregates your active habits, targets, completed checks, and completion percentage by life domain:
  - *Health & Fitness*
  - *Mindset & Learning*
  - *Career & Focus*
  - *Personal & Home*
  - *Finance & Wealth*
- Includes a dedicated **Domain Performance** column chart to visualize which areas of your life received the most consistent execution.

### 4. Sleep & Productivity Correlation Analysis
- Categorizes days into **Optimal Rest (≥ 7 Hours)** vs. **Sleep Deficit (< 7 Hours)**.
- Automatically calculates:
  - Average habit completion % on well-rested days.
  - Average habit completion % on sleep deficit days.
  - **Rest Advantage**: Quantifies the exact percentage boost in habit execution when you get 7+ hours of sleep!
  - Dynamic AI Callout: *"You complete X% MORE habits on days with 7+ hours of sleep!"*

### 5. Standardized Annual & Quarterly Data Bridge
- Standardized data row at **Row 92** containing:
  `Month | Year | Active Habits | Target Days | Done Checks | Success % | Perfect Days | Avg Sleep | Best Week | Top Life Domain`
- **How to Use**: Reference or copy Row 92 directly into your Annual or Quarterly review dashboards.

### 6. Monthly Retrospective & Action Plan
Three structured coaching boxes for end-of-month self-reflection:
- 🏆 **Monthly Wins & Highlights**: Top victories and effortless consistency.
- ⚠️ **Friction Points & Root Causes**: Triggers of skipped days and energy drains.
- 🎯 **3 Key Adjustments for Next Month**: Environment tweaks and non-negotiable protocols.

---

## 🎨 Color Palette Reference

| Element | Theme Name | Hex Code | Visual Purpose |
| :--- | :--- | :--- | :--- |
| **Habit Group 1** | Sky Blue | `#E0F2FE` | Fresh, focused start (Rows 18–21) |
| **Habit Group 2** | Lavender | `#F3E8FF` | Calm mental flow (Rows 22–25) |
| **Habit Group 3** | Warm Sand | `#FEF3C7` | Warm, energizing amber (Rows 26–29) |
| **Habit Group 4** | Cool Slate | `#F1F5F9` | Balanced slate (Rows 30–33) |
| **Habit Group 5** | Periwinkle | `#EEF2FF` | Harmonious twilight (Rows 34–37) |
| **Completed Cell (`✓`)** | Mint Emerald | `#DCFCE7` | High-contrast accomplishment |
| **Missed Cell (`✗`)** | Soft Coral | `#FEE2E2` | Gentle accountability |
| **"Today" Column** | Golden Glow | `#FEF08A` | Automatic location indicator |
| **Sleep Header** | Night Indigo | `#1E1B4B` | Sleep theme header |
| **Sleep Day Cells** | Bedtime Blue | `#EEF2FF` | Calming contrast for sleep hours |
| **Notes Box** | Pale Memo Yellow | `#FFFFE6` | Sticky-note memo space |

---

## ❓ Frequently Asked Questions (FAQ)

### Q: Can I use this in Google Sheets, Microsoft Excel, or LibreOffice?
**A**: Yes! The file is standard `.xlsx` (OpenXML). All formulas (`COUNTIF`, `SUMIF`, `AVERAGEIF`, `IFERROR`, `INDEX/MATCH`) and charts are natively supported across Microsoft Excel (Desktop, Web, Mobile), Google Sheets, and LibreOffice Calc.

### Q: What happens if I track fewer than 20 habits?
**A**: Simply leave unused rows blank. Because of the built-in auto-discard logic, blank rows will never count toward your daily score, weekly totals, or domain statistics.

### Q: How do I add a new habit mid-month?
**A**: Enter the habit name in Column B, select its category in Column C, select its target in Column D, and start selecting `✓` or `✗` from today onwards. Past blank days are excluded automatically.

---

## 💻 Developer & Verification Guide

- **Generator Script**: [`create_excel_tracker.py`](./create_excel_tracker.py)
  - Contains complete workbook generation logic, conditional formatting, analytics tables, and charts.
  - To regenerate all 52 sheets:
    ```bash
    python create_excel_tracker.py
    ```
- **Automated Verification Script**: [`verify_excel_tracker.py`](./verify_excel_tracker.py)
  - Tests all 52 sheets, verifying column structures, categories, KPI cards, tables, correlation formulas, and all 3 charts per sheet:
    ```bash
    python verify_excel_tracker.py
    ```

---

*Transform your consistency, sleep, and performance month by month from 2026 to 2030!* 🚀
