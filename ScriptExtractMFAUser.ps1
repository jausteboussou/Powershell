# Ensure AzureAD module is installed and imported
Import-Module AzureAD


$ClientId = '***'
$TenantId = '***'
$CertThumbprint = '***'
$Cert = Get-ChildItem Cert:\LocalMachine\My\$CertThumbprint
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $Cert

# Connect to Azure AD
Connect-AzureAD
Connect-MsolService

# Get all Azure AD users
$users = Get-AzureADUser -All $true

# Filter out guest and resource accounts
$filteredUsers = $users #| Where-Object { $_.UserType -ne 'Guest' -and $_.ObjectType -ne 'ServicePrincipal' }

$results=@();
Write-Host  "`nRetreived $($users.Count) users";
#loop through each user account
foreach ($user in $users) {

    Write-Host  "`n$($user.UserPrincipalName)";
    $myObject = [PSCustomObject]@{
        User                   = "-"
        MFAstatus              = "_"
        Email                  = "-"
        Fido2                  = "-"
        MicrosoftAuthenticator = "-"
        Password               = "-"
        SMS                    = "-"
        OtherApp               = "-"
        TempAccess             = "-"
        HelloBusiness          = "-"
        JobTitle               = "-"
        OfficeLocation         = "-"
        OnPremiseAccount       = "-"
        EmployeeType           = "-"
        AccountType            = "-"
    }

    $MFAData=Get-MgUserAuthenticationMethod -UserId $user.UserPrincipalName #-ErrorAction SilentlyContinue

    $myObject.user = $user.UserPrincipalName
    $myObject.jobtitle = $user.JobTitle
    $myObject.officelocation = $user.OfficeLocation
    $myObject.OnPremiseAccount = 'No' 
    if ($user.DirSyncEnabled) { 
        $myObject.OnPremiseAccount = 'Yes' 
    }
    if ($user.ExtensionProperty.EmployeeType) { 
        $myObject.EmployeeType = $user.ExtensionProperty.EmployeeType
    }
    if ($user.AccountType) { 
        $myObject.AccountType = $user.ObjectType
    }

    #check authentication methods for each user
    ForEach ($method in $MFAData) {
    
        Switch ($method.AdditionalProperties["@odata.type"]) {
        "#microsoft.graph.emailAuthenticationMethod"  { 
            $myObject.Email = $true 
            $myObject.MFAstatus = "Enabled"
        } 
        "#microsoft.graph.fido2AuthenticationMethod"                   { 
            $myObject.Fido2 = $true 
            $myObject.MFAstatus = "Enabled"
        }    
        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"  { 
            $myObject.MicrosoftAuthenticator = $true 
            $myObject.MFAstatus = "Enabled"
        }    
        "#microsoft.graph.passwordAuthenticationMethod"                {              
                $myObject.Password = $true 
                # When only the password is set, then MFA is disabled.
                if($myObject.MFAstatus -ne "Enabled")
                {
                    $myObject.MFAstatus = "Disabled"
                }                
        }     
        "#microsoft.graph.phoneAuthenticationMethod"  { 
            $myObject.SMS = $true 
            $myObject.MFAstatus = "Enabled"
        }   
            "#microsoft.graph.softwareOathAuthenticationMethod"  { 
            $myObject.OtherApp = $true 
            $myObject.MFAstatus = "Enabled"
        }           
            "#microsoft.graph.temporaryAccessPassAuthenticationMethod"  { 
            $myObject.TempAccess = $true 
            $myObject.MFAstatus = "Enabled"
        }           
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod"  { 
            $myObject.HelloBusiness = $true 
            $myObject.MFAstatus = "Enabled"
        }                   
        }
    }
    ##Collecting objects
    $results+= $myObject;

}

$result | Export-Excel -Path "D:\Scripts\MFA\20231117_MFAStatus_AAD.xlsx" -TableStyle Medium9

$$user = Get-MsolUser -UserPrincipalName "***"
$$user | Select-Object DisplayName, UserPrincipalName, @{Name="StrongAuthenticationMethods";Expression={$_.StrongAuthenticationMethods}}, 
@{Name="StrongAuthenticationPhoneAppDetails";Expression={$_.StrongAuthenticationPhoneAppDetails}}, StrongAuthenticationProofupTime, 
@{Name="StrongAuthenticationRequirements";Expression={$_.StrongAuthenticationRequirements}}, @{Name="StrongAuthenticationUserDetails";
Expression={$_.StrongAuthenticationUserDetails}}, StrongPasswordRequired
