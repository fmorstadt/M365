<# 
 .Synopsis
  M365 license feature assignment to license on a per user base
  
 .Description
  When AzureAD P1 is not available, no group assignment is possible
  so, no group based license feature set is available
  In the first step, the actual settings will be exported
  The CSV file can be edited
  Run the script a second time with files in place
  The settings defined in the CSV file wil be applied

 .PARAMETER
  
 .NOTES
  (c) 2020 ByteRunner/Frank Morstadt
  Reference:
  https://docs.microsoft.com/de-de/azure/active-directory/users-groups-roles/licensing-service-plan-reference
  
  VERSION:
  17.09.2020 V1.0 Basis-Version (getestet und freigegeben)
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

$objUsrLst = (Get-MsolUser | Where-Object { $_.islicensed } | Select-Object -Property UserPrincipalName -ExpandProperty licenses) | Select-Object UserPrincipalName, AccountSkuId
$objLicLst = Get-MsolAccountSku
$objLicCsv = @()
$strLicCsv = $strps1dir + "\" + $strps1nam + ".csv"

##### License Expansion and Export to CSV #####
if (Test-Path $strLicCsv) {
    Write-Host "Import file found - applying license service plan assignments"
    try { $objCsvImp = Import-Csv $strLicCsv -Delimiter ";" } catch { throw }
      
   
    foreach ($objRow in $objUsrLst) {
        $strLicSku = $objrow.AccountSkuId
        $strUsrUpn = $objrow.UserPrincipalName
        $objLicSpl = $objCsvImp | Where-Object { $_.SKU -eq $strlicsku }
        $objLicDis = $objLicSpl | Where-Object { $_.Active -ne "1" } | Select-Object ServicePlan 

        $hshLicDis = @()
        $objLicDis | ForEach-Object { $hshLicDis += $_.ServicePlan }
        $objLicOpt = New-MsolLicenseOptions -AccountSkuId $strLicSku -DisabledPlans $hshLicDis
        Write-Host "Service plan $strLicSku for User $strUsrUpn"
        try {Set-MsolUserLicense -UserPrincipalName $strUsrUpn -LicenseOptions $objLicOpt} catch {Write-Error "Set $strLicSku for User $strUsrUpn failed"}
    }
}
else {
    Write-Host "Import files NOT found - generating CSV files from Azure AD"
    $objLicArr = @()
    foreach ($objLic in $objLicLst) {
    
        $objLicSpl = $objLic | Select-Object -ExpandProperty ServiceStatus
        foreach ($objSpl in $objLicSpl) {
            if (($objSpl.ProvisioningStatus -eq "Success") -or ($objSpl.ProvisioningStatus -eq "PendingProvisioning")) {
                $objTmpRes = New-Object PSCustomObject
                $objTmpRes | Add-Member -type NoteProperty -Name SKU -Value $objLic.AccountSkuId
                $objTmpRes | Add-Member -type NoteProperty -Name ServicePlan -Value $objSpl.ServicePlan.ServiceName
                $objTmpRes | Add-Member -type NoteProperty -Name Active -Value "1"
                $objLicArr += $objTmpRes
            }
        }
    }
    $objLicArr | Export-Csv $strLicCsv -Delimiter ";" -NoTypeInformation
}
##### 

Stop-Transcript


    


