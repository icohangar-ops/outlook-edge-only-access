<#
    Sample category map for Group-RecipientsByCategory.ps1.

    Copy this, then fill it in for your own use. Every entry below is a
    well-known public company used purely as an illustration - none of this
    reflects any real relationship.

    Format: 'domain.com' = @('Category', 'Display name')

    Assign a category only where you have evidence for it. Anything you are
    unsure about is better left out: the script routes unmapped domains to
    'Review - Unclassified' so they surface rather than being quietly
    mislabelled.
#>
@{
    # --- capital providers ---
    'example-ventures.com'   = @('Investor - Venture',        'Example Ventures')
    'example-pe.com'         = @('Investor - Private equity',  'Example PE Partners')
    'example-strategic.com'  = @('Investor - Strategic',       'Example Strategic Corp')

    # --- intermediaries running a raise ---
    'example-bank.com'       = @('Advisor - Investment bank',  'Example Bank')
    'example-broker.com'     = @('Advisor - Broker',           'Example Broker')

    # --- debt and asset finance ---
    'example-lender.com'     = @('Lender - Bank',              'Example Bank of Commerce')
    'example-leasing.com'    = @('Lender - Leasing',           'Example Leasing')

    # --- explicitly not capital providers ---
    'example-law.com'        = @('Excluded - Legal',           'Example LLP')
    'example-audit.com'      = @('Excluded - Auditor',         'Example Audit')
    'example-saas.com'       = @('Excluded - Vendor',          'Example SaaS')
}
