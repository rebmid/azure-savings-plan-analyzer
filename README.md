# Azure Savings Plan Analyzer

Read-only tooling that measures an Azure estate's Compute Savings Plan opportunity and
builds a PowerPoint deck showing, per subscription and per business unit, where a savings
plan beats the current negotiated on-demand discount. You edit ONE file (`Config.ps1`),
then run the numbered scripts in order.

## What it produces
- CSVs: per-subscription eligible compute, 1-year and 3-year savings-plan recommendations,
  VM instance-hours, a full readable-subscription cost audit, and a business-unit rollup.
- A PowerPoint deck: a methodology slide, a summary slide, and one detail slide per
  business unit.

## Prerequisites
- **Azure access:** Reader (plus Cost Management Reader / Advisor Reader) across the
  subscriptions you want analyzed, in your own tenant.
- **PowerShell** with the **Az** module (`Az.Accounts`, `Az.Resources`, `Az.CostManagement`).
- **Desktop PowerPoint on Windows** for the deck-building scripts (06, 07). The data-pull
  scripts (01-05) run anywhere the Az module runs; only the slide steps need Office.

## The economic model (IMPORTANT: confirm against your own contract)
The dollar figures depend on two contract-specific inputs set in `Config.ps1`:

- `$retailDiscount` - your current negotiated discount off public list.
- `$spd` - the additional savings-plan discount your agreement stacks on top.

The two published Compute Savings Plan rates (`$standardSp1`, `$standardSp3`) are Microsoft
list values and normally do not change.

Model (multiplicative, computed off the public PAYG list):
- `publicRetail    = current spend / (1 - retailDiscount)`
- covered/steady portion `= publicRetail * (1 - standardSp) * (1 - spd)` (the SP rate and
  the SPD both stack off list)
- uncovered portion pays full list (the negotiated on-demand discount is forfeited on a plan)

Coverage % is a pure ratio, so it is price-level independent; only the dollar deltas move
when you change the discount inputs. If you do not update `$retailDiscount` and `$spd` to
your own numbers, the savings dollars will be wrong.

## How to run

### Step 0 - Connect to your Azure tenant (read access is enough)
```powershell
Connect-AzAccount -Tenant <your-tenant-guid>
(Get-AzContext).Tenant.Id          # confirm you are on the right tenant
Set-Location "<this folder>"
```

### Step 1 - Fill in the top of Config.ps1
- `$tenant` : your tenant GUID
- `$periodFrom` / `$periodTo` : a full recent complete month
- `$retailDiscount` / `$spd` : your negotiated discount + savings-plan discount
- `$deckName` : keep or rename

Leave `Get-BU` and `$businessUnits` empty for now; you fill them after Step 2.

### Step 2 - Pull the data (read-only Azure queries)
```
.\01-Analyze-PerSubCompute.ps1     -> PerSubCompute.csv        (every sub + eligible compute)
.\02-Analyze-SavingsPlanBySub.ps1  -> SavingsPlanBySub.csv     (3-year, 30-day look-back)
.\03-Pull-1yr3yr.ps1               -> SavingsPlan_1yr_3yr.csv  (1-year + 3-year)
.\04-Compute-EstateVMHours.ps1     -> VMHoursBySub.csv         (VM instance-hours)
.\05-Audit-AllSubCost.ps1          -> AllSubCostAudit.csv      (all readable subs)
```

### Step 3 - Fill Get-BU and $businessUnits in Config.ps1
Open `PerSubCompute.csv` and `AllSubCostAudit.csv` to see the real subscription NAMES.
Group them under each business unit, in the order you want the slides to appear.

### Step 4 - Build the deck
```
.\06-Add-MethodologySlide.ps1      (creates the deck + slide 1: methodology)
.\07-Build-BUSavingsBreakout.ps1   (slide 2 summary + one detail slide per business unit)
```

## Notes
- The deck's column labels and footnotes contain the literal text "32.5%" (matching the
  sample `$spd`). The math itself is driven by `$spd`, but if you change `$spd` update that
  label text in scripts 06 and 07 so the slides read correctly.
- Re-run 02-05 any time to refresh the 30-day look-back, then re-run 07.
- Script 07 deletes and rebuilds every slide after slide 1, so run 06 first, then 07.
- The Cost Management `benefitRecommendations` API (2024-08-01) returns
  `commitmentAmount`, `benefitCost`, `totalCost`, and `costWithoutBenefit`; the scripts
  derive coverage and savings from those fields (with a fallback to the legacy
  `coveragePercentage` / `savingsAmount` fields if a tenant still returns them).
- Generated CSV, PPTX and PNG outputs are git-ignored; they contain tenant-specific data.

## Deck column reference
Each detail row shows: current $/mo and $/hr, Steady % (savings-plan coverage), public PAYG
$/mo, 1-year and 3-year savings-plan cost, commit $/hr (1yr / 3yr), the recommendation, and
the 3-year annual savings versus the current negotiated discount.
