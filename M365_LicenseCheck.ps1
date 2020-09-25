<# 
 .Synopsis
  M365 license overview and assignment
  
 .Description
  Connects to MSOnline and checks license stat and assignments
  Export to CSV

 .PARAMETER
  
 .NOTES
  (c) 2020 ByteRunner/Frank Morstadt
   
  VERSION:
  17.09.2020 V1.0 Basis-Version (getestet und freigegeben)
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$bolAdmMfa = $false # Azure Admin account needs MFA for authentication?

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

Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
if (!(Get-Module -ListAvailable -Name "MSOnline")) { Install-Module -Name MSOnline -Scope AllUsers }
try { Import-Module -Name MSOnline } catch { Throw }

if ($bolAdmMfa) {
    Connect-MsolService
}
else {
    if (!($objMsoCre)) { $objMsoCre = Get-Credential }
    Connect-MsolService -Credential $objMsoCre
}

try { Connect-MsolService -Credential $objMsoCre } catch { throw; break }

$licensePlanList = Get-MsolAccountSku
$objUsr = $null
$objUsrLst = Get-MsolUser -All | Where-Object { $_.islicensed }
$objResult = @()
$strUsrCsv = $strps1dir + "\" + $strps1nam + ".csv"
$strLicCsv = $strps1dir + "\" + $strps1nam + "_Count.csv"
$strLicUsr = $strps1dir + "\" + $strps1nam + "_User.csv"

foreach ($objUsr in $objUsrLst) {

    $objUsrLic = $objUsr.Licenses
    
    #$objUsrLic.AccountSkuId

    foreach ($objLic in $objUsrLic) {
        $objLicSta = $objlic.ServiceStatus | Where-Object { ($_.ProvisioningStatus -eq "Success") -or ($_.ProvisioningStatus -eq "PendingProvisioning") }

        foreach ($objRow in $objLicSta) {

            $objTmpRes = New-Object PSCustomObject
            $objTmpRes | Add-Member -type NoteProperty -Name UPN -Value $objUsr.UserPrincipalName
            $objTmpRes | Add-Member -type NoteProperty -Name SKU -Value $objLic.AccountSkuId
            $objTmpRes | Add-Member -type NoteProperty -Name ServicePlan -Value $objRow.ServicePlan.ServiceName
            $objResult += $objTmpRes
        }
    }
}

$objResult | Export-Csv $strUsrCsv -Delimiter ";" -NoTypeInformation
$objResult | Select-Object -Unique UPN, SKU | Sort-Object UPN | Export-Csv $strLicUsr -Delimiter ";" -NoTypeInformation
$licensePlanList | Export-Csv $strLicCsv -Delimiter ";" -NoTypeInformation
