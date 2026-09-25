# Minimal test harness shared by the test scripts (no Pester needed). Dot-source it.
$script:passed = 0; $script:failed = 0
function It {
    param([string]$Name, [scriptblock]$Test)
    try { & $Test; $script:passed++; Write-Host "  [pass] $Name" -ForegroundColor Green }
    catch { $script:failed++; Write-Host "  [FAIL] $Name`n         $($_.Exception.Message)" -ForegroundColor Red }
}
function Assert-Equal {
    param($Expected, $Actual, [string]$Because = '')
    $e = ($Expected | ForEach-Object { "$_" }) -join ','; $a = ($Actual | ForEach-Object { "$_" }) -join ','
    if ($e -ne $a) { throw "Expected [$e] but got [$a]. $Because" }
}
function Assert-True { param($Condition, [string]$Because = '') if (-not $Condition) { throw "Expected true. $Because" } }
function Complete-Tests {
    Write-Host ''
    Write-Host "$script:passed passed, $script:failed failed" -ForegroundColor $(if ($script:failed) { 'Red' } else { 'Green' })
    if ($script:failed) { exit 1 }
}
