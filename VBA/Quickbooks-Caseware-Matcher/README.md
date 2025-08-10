# QuickBooks → CaseWare Account Matcher (VBA)

This VBA macro automates the process of matching QuickBooks trial balance account names to CaseWare account numbers and names.  
It speeds up the import process into CaseWare by performing **exact matches** first, then using **fuzzy matching** if exact matches fail.

---

## ✨ Features

- **Exact Match Pass**  
  Uses Excel’s `XLOOKUP` to instantly match account names between QuickBooks and CaseWare.  

- **Fuzzy Match Pass**  
  For unmatched accounts, applies a **Levenshtein distance** + token overlap similarity check to suggest the closest match.  

- **Customizable Match Thresholds**  
  - **Good Match:** ≥ 0.84 → auto-assign & highlight yellow.  
  - **Possible Match:** ≥ 0.74 → assign & highlight peach.  
  - Below threshold → mark as “No good match” in red.  

- **Formatting Helper**  
  Auto-cleans and formats the trial balance for easier processing.

---

## 📂 Repository Structure

    QuickBooks-CaseWare-Matcher/
    │
    ├─ src/
    │   └─ MatchAccounts.bas           # VBA macro module
    │
    ├─ sample-data/
    │   └─ trial_balance_sample.xlsx   # Sample QuickBooks TB for testing
    │
    ├─ README.md                       # Project documentation
    └─ LICENSE                         # (Optional) Usage license

---

## 📋 Requirements

- Microsoft Excel (tested on Excel 2016+)
- VBA enabled (press `Alt + F11` to open the VBA editor)

---

## 🚀 Setup & Usage

1. **Download the repository**  
   - Click **Code → Download ZIP**, or clone with:
     ```bash
     git clone https://github.com/YOUR-USERNAME/QuickBooks-CaseWare-Matcher.git
     ```

2. **Import the VBA module into Excel**  
   - Open Excel.  
   - Press `Alt + F11` to open the VBA editor.  
   - Go to `File → Import File…` and select `src/MatchAccounts.bas`.

3. **Prepare your data**  
   - **Sheet1** (trial balance):  
     - Column B: QuickBooks account names (to match)  
     - Column A & C: Leave blank — macro will fill CaseWare account # (A) and CaseWare name (C)  
   - **Sheet2** (CaseWare chart of accounts):  
     - Column A: CaseWare account numbers  
     - Column B: CaseWare account names

4. **Run the macro**  
   - Press `Alt + F8`, select `MatchAccounts_All`, and click **Run**.  
   - Review highlighted results:  
     - **Yellow** → Strong match (≥ 0.84)  
     - **Peach** → Possible match (≥ 0.74)  
     - **Red** → No good match found

---

## 🧪 Example Output

| Account # | Excel Account Name        | CaseWare Name            | Match Type/Score |
|-----------|---------------------------|--------------------------|------------------|
| 1000      | RBC Loan                  | RBC Loan                 | Exact            |
| 2100      | Accounts Receivable       | Accounts Receivable      | Fuzzy (0.87)     |
|           | Random Misc Expense       |                          | No good match    |

---

## ⚙️ Configuration

At the top of the VBA module, you can adjust:
```vba
Const S1_NAME As String = "Sheet1" ' Trial balance sheet
Const S2_NAME As String = "Sheet2" ' CaseWare COA sheet
Const GOOD_MATCH As Double = 0.84
Const MIN_MATCH  As Double = 0.74
