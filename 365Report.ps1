function Connect_Graph{
    #check if graph is installed.
    $MSGRAPH = get-module -name microsoft.graph -ListAvailable
    if ($msgraph.count -eq 0){
        Install-Module Microsoft.Graph -repository PSGallery -AllowClobber -Force
    }
    Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All","Organization.Read.All","UserAuthenticationMethod.Read.All","Auditlog.Read.All"
    Connect-exchangeonline
}
function Graph_info{
    $FriendlyLicenseNames = Get-Content -raw -path .\LicenseFriendlyName.txt | ConvertFrom-StringData
    $MFAFriendlyNames = Get-content -raw -path .\mfamethods.txt | ConvertFrom-StringData
    $Users = get-mguser -property "DisplayName","UserPrincipalName","ID","SigninActivity"
    $disabledusers= Get-MgUser -Filter 'accountEnabled eq false' -All
    $exportobject=@()
    foreach ($user in $users){
        $AuthMethods=@()
        $license=@()
        $SKUs= Get-MgUserLicenseDetail -userID $user.ID
        $LastSignIn = $user.SignInActivity.LastSignInDateTime
        
        if ($Disabledusers.Displayname -contains $User.DisplayName){
            $SigninStatus='Blocked'
        }
        else{
            $signinstatus='Allowed'
        }

        $mfatypes = Get-MgUserAuthenticationMethod -UserID $user.UserPrincipalName
        foreach ($mfatype in $mfatypes.AdditionalProperties){
            $MFAKey = $mfatype['@odata.type']-replace'#'
            $methodtoadd = $MFAFriendlyNames[$MFAKey]
            $AuthMethods+= $methodtoadd
        }

        foreach($sku in $skus){
                $license += $FriendlyLicenseNames[$sku.SkuPartNumber]
        }

        $MailBoxType = (get-mailbox -identity $user.UserPrincipalName -erroraction 'silentlycontinue').recipienttypedetails
        if ($MailboxType -eq $null){
            $Mailboxtype = 'No Mailbox'
        }
        $exportobject += new-object psobject -property ([ordered] @{Name=$user.Displayname;UPN=$user.UserPrincipalName;Licenses=$license -join ", ";SigninStatus=$SigninStatus;LastSignin=$lastSignIn;Authmethods=$authmethods -join ", ";Mailboxtype=$Mailboxtype})
    }  
    $exportobject | export-csv .\365report.csv -notypeinformation
}

Connect_Graph
Graph_info
disconnect-mggraph 
