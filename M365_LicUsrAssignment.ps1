<# 
 .Synopsis
  M365 license assignment to user accounts
  
 .Description
  When AzureAD P1 is not available, no group assignment is possible
  In the first step, the actual settings will be exported
  The CSV file can be edited
  Run the script a second time with files in place
  The settings defined in the CSV file wil be applied

 .PARAMETER
  
 .NOTES
  (c) 2020 ByteRunner/Frank Morstadt
  
  VERSION:
  17.09.2020 V1.0 Base-Version
  23.09.2020 V1.1 Redesign CSV file: UPN;SKU_OLD;SKU_NEW
                  Adding Transcript
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region MyInvocation
if ($hostinvocation -ne $null)
{ $strps1dir = Split-Path $hostinvocation.MyCommand.path }
else
{ $strps1dir = Split-Path $MyInvocation.MyCommand.Path }
if ($hostinvocation)
{ $strps1nam = $hostinvocation.MyCommand.Name.split(".")[0] }
else
{ $strps1nam = $MyInvocation.MyCommand.Name.split(".")[0] }
#endregion MyInvocation

#region Transcript
$strTrsPfd = $strps1dir + "\" + $strps1nam + "_" + (Get-Date -Format yyyyMMddHHmm) + ".log"
Start-Transcript -Path $strTrsPfd
#endregion Transcript

$InformationPreference = "Continue"
$bolAdmMfa = $false # Azure Admin account needs MFA for authentication?
$bolLocAdm = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") #actual context local Admin?
$objUsrCsv = @()
$strUsrCsv = $strps1dir + "\" + $strps1nam + ".csv"
$strLicCsv = $strps1dir + "\" + $strps1nam + "_Count.csv"

if (!(Get-Module -ListAvailable -Name "MSOnline")) { 
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    if ($bolLocAdm) {
        Install-Module -Name MSOnline -Scope AllUsers 
    }
    else {
        Install-Module -Name MSOnline -Scope CurrentUser
    }
}
try { Import-Module -Name MSOnline } catch { Throw }

if ($bolAdmMfa) {
    Connect-MsolService
}
else {
    if (!($objMsoCre)) { $objMsoCre = Get-Credential }
    Connect-MsolService -Credential $objMsoCre
}

$objUsrLst = Get-MsolUser -All
$objLicLst = Get-MsolAccountSku

if (Test-Path $strUsrCsv) {
    Write-Host "Import file found - applying user license assignments"
    try { $objCsvImp = Import-Csv $strUsrCsv -Delimiter ";" } catch { throw }
    $objCsvImp = $objCsvImp | Sort-Object UPN, SKU_NEW

    foreach ($objRow in $objCsvImp) {
        $strSrcUpn = $objrow.UPN
        $strSkuNew = ($objrow.SKU_NEW).trim()
        $strSkuOld = ($objrow.SKU_OLD).trim()
        $objMsoUsr = $objUsrLst | Where-Object { $_.UserPrincipalName -eq $strSrcUpn }
        $objLicArr = $objMsoUsr.Licenses | Select-Object AccountSkuId  #get all assigned license SKUs
        
        if ($strSkuNew -ne $strSkuOld) {
            if ($strSkuOld) {
                if ($objLicArr | Where-Object { $_.AccountSkuId -eq $strSkuOld }) {
                    Write-Host "Remove $strSkuOld from User $strSrcUpn"
                    try {Set-MsolUserLicense -UserPrincipalName $strSrcUpn -RemoveLicenses $strSkuOld} catch {Write-Error "Remove $strSkuOld from User $strSrcUpn failed"}
                }
                else {
                    Write-Warning "$strSkuOld not found - User $strSrcUpn"
                }
            }
            if ($strSkuNew) {
                if (!($objLicArr | Where-Object { $_.AccountSkuId -eq $strSkuNew })) {    #check if license already assigned
                    Write-Host "Add $strSkuNew to User $strSrcUpn"
                    try {Set-MsolUserLicense -UserPrincipalName $strSrcUpn -AddLicenses $strSkuNew} catch {Write-Error "Add $strSkuNew to User $strSrcUpn failed"}
                }
                else {
                    Write-Warning "$strSkuNew already applied - User $strSrcUpn"
                }
            }
        }
    }
}
else {
    Write-Host "Import files NOT found - generating CSV files from Azure AD"
    foreach ($objUsr in $objUsrLst) {
        if ($objUsr.isLicensed) {
            $objUsrLic = $objUsr.Licenses
            foreach ($objLic in $objUsrLic) {
                $objTmpRes = New-Object PSCustomObject
                $objTmpRes | Add-Member -type NoteProperty -Name UPN -Value $objUsr.UserPrincipalName
                $objTmpRes | Add-Member -type NoteProperty -Name SKU_OLD -Value $objLic.AccountSkuId
                $objTmpRes | Add-Member -type NoteProperty -Name SKU_NEW -Value $objLic.AccountSkuId
                $objUsrCsv += $objTmpRes
            }
        }
        else {
            $objTmpRes = New-Object PSCustomObject
            $objTmpRes | Add-Member -type NoteProperty -Name UPN -Value $objUsr.UserPrincipalName
            $objTmpRes | Add-Member -type NoteProperty -Name SKU_OLD -Value $null
            $objTmpRes | Add-Member -type NoteProperty -Name SKU_NEW -Value $null
            $objUsrCsv += $objTmpRes
        }
    }
    $objUsrCsv | Sort-Object UPN | Export-Csv $strUsrCsv -Delimiter ";" -NoTypeInformation
    $objLicLst | Select-Object -Property AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits | Export-Csv $strLicCsv -Delimiter ";" -NoTypeInformation
}

Stop-Transcript