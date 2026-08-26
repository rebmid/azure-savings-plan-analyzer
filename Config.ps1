# ============================================================================
#  AZURE SAVINGS PLAN ANALYSIS - CENTRAL CONFIG   (edit this one file)
#  Every numbered script dot-sources this file. Fill in the << >> placeholders.
#
#  WHAT YOU MUST CHANGE FOR YOUR ENVIRONMENT:
#    1. $tenant                 -> your Azure tenant GUID
#    2. $periodFrom / $periodTo -> a full recent complete month
#    3. $retailDiscount / $spd  -> YOUR negotiated discount + savings-plan discount
#                                  (these are contract-specific; do NOT assume the
#                                   sample values below apply to your agreement)
#    4. Get-BU + $businessUnits -> your real subscription names / groupings
# ============================================================================

# --- Your Azure tenant GUID.  After Connect-AzAccount run: (Get-AzContext).Tenant.Id
$tenant = '<<your-tenant-guid>>'

# --- Usage period for the cost / VM-hours pulls. Use a full, recent, complete month.
$periodFrom = '<<yyyy-MM-01T00:00:00Z>>'
$periodTo   = '<<yyyy-MM-lastdayT23:59:59Z>>'   # last day of that month

# --- Output PowerPoint deck (created by script 06, rebuilt by 07).
$deckName = 'Savings Plan Opportunity.pptx'

# --- OPTIONAL prior-access baseline: a CSV with a 'SUBSCRIPTION NAME' column, used only
#     to flag which subscriptions are newly readable. Leave '' on the first run.
$priorSubsCsv = ''

# ============================================================================
#  ECONOMIC MODEL  (CONTRACT-SPECIFIC - confirm against your own agreement)
#
#  $retailDiscount : your current negotiated discount off public list (e.g. 0.40 = 40% off).
#  $spd            : the additional savings-plan discount your agreement adds on top,
#                    applied to the covered/steady portion (e.g. 0.325 = 32.5%).
#  $standardSp1/3  : FALLBACK published Compute Savings Plan discounts off list.
#                    Analyze-SavingsPlanAuto.ps1 prices the covered portion from each subscription's
#                    live Cost Management recommendation (benefitCost); these constants are used only
#                    when a subscription/term has no usable recommendation.
#
#  Model (multiplicative, off public PAYG list):
#    publicRetail    = current40Spend / (1 - retailDiscount)
#    coveredPortion  = publicRetail * (1 - standardSp) * (1 - spd)   [SP rate AND spd both
#                                                                     stacked off list]
#    uncoveredPortion= publicRetail  (full list; the negotiated on-demand discount is
#                                     forfeited on a savings plan)
#  Coverage % is a pure ratio, so it is price-level independent; only the dollar deltas
#  move when you change the discount inputs.
# ============================================================================
$retailDiscount = 0.40    # <<< YOUR negotiated discount off list (sample: 0.40)
$spd            = 0.325   # <<< YOUR additional savings-plan discount (sample: 0.325)
$standardSp1    = 0.322   # FALLBACK 1-year Compute Savings Plan discount off list (auto script derives this live)
$standardSp3    = 0.528   # FALLBACK 3-year Compute Savings Plan discount off list (auto script derives this live)

# ============================================================================
#  BUSINESS UNITS  (fill in after you can see your subscription names)
#  Get-BU        : pattern rules the DATA pulls use to tag each subscription.
#  $businessUnits: exact subscription membership used to BUILD the slides,
#                  in the order you want the slides to appear.
#
#  Run scripts 01 and 05 first, open PerSubCompute.csv / AllSubCostAudit.csv to see the
#  real subscription NAMES, then map them below.
# ============================================================================

function Get-BU([string]$name){
    switch -Regex ($name){
        # '^(Prod|Production)'  { 'Workloads'; break }
        # '^(Hub|Identity|Mgmt)' { 'Platform';  break }
        default { 'Unassigned' }
    }
}

$businessUnits = [ordered]@{
    # 'Workloads' = @('<<Subscription A>>','<<Subscription B>>')
    # 'Platform'  = @('<<Subscription C>>','<<Subscription D>>')
}
