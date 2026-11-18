# HR Leave and Onboarding Tracker

An Excel workbook for the everyday records an HR administrator keeps: annual leave requests and balances, absence by department, and a checklist for new starters.

**All employees and requests are made up. No real personal data is stored.**

[Open the workbook](hr_leave_onboarding.xlsx)

## What it does

**Leave**
- Counts the working days in each request, skipping weekends and Irish public holidays for 2026 ([source](https://www.citizensinformation.ie/en/employment/employment-rights-and-conditions/leave-and-holidays/public-holidays/))
- Shows each employee's entitlement, leave taken, leave booked, pending requests and days remaining
- Flags anyone over their entitlement, with pending requests above their balance, or with 10 or more days left late in the year

**Absence summary**
- Approved days off by department and leave type, with a chart

**Onboarding**
- An 8-step checklist for each new starter: contract, payroll details, IT setup, access card, induction, GDPR training, health and safety, and a 30-day check-in
- Marks tasks as Done, Open or Overdue against the report date, and shows % complete per person

## Excel skills used

`NETWORKDAYS`, `SUMIFS`, `COUNTIFS`, `COUNTIF`, `INDEX` + `MATCH`, `IFERROR`, `TEXT`, drop-down lists, conditional formatting, filters and a stacked bar chart.

## Try it

1. Add a leave request and pick the employee from the drop-down list.
2. Change a request from Pending to Approved and watch the balance update.
3. Change the **report date** on the Settings sheet to see onboarding tasks become overdue.

Leave entitlement is typed in per employee. Check [Citizens Information](https://www.citizensinformation.ie/en/employment/employment-rights-and-conditions/leave-and-holidays/annual-leave/) for the legal minimums before using this for real staff.

## Rebuilding the file

```
pip install openpyxl
python build_workbook.py
```
