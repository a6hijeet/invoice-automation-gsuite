# Google Form Fields

The Apps Script depends on **exact form field labels**.

⚠️ Field names are CASE-SENSITIVE

---

## Section 1 – Room Invoice

- Invoice Number
- Invoice Date
- Company Name
- Company Address
- Customer Email
- Customer Email CC
- GSTIN
- Check In (Date and time)
- Check Out (Date and Time)
- Guest Name
- Stayed At
- Persons
- Days
- Rate/Day

Repeat optional blocks:
- Guest Name 2 → Rate/Day 2
- Guest Name 3 → Rate/Day 3
- Guest Name 4 → Rate/Day 4

---

## Section 2 – Food Invoice (Optional)

Food Name fields:
- Food1 → Food7

Food Quantity:
- Food1 Qty → Food7 Qty

Food Rate:
- Food1 Rate → Food7 Rate

📌 Food invoice is generated **only if Food1 Qty is filled**
