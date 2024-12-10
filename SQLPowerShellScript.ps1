# Use your Azure account that has User administration permissions and Directory Read permissions. Make sure you disconnect after the script is completed.. line 61

Connect-AzureAD -AzureEnvironmentName AzureUSGovernment -TenantDomain "xyz.onmicrosoft.com" -AccountId firstlastname@xyz.com

 

# Make sure that there are not spaces after any of the fields in the line 1 - csv header rows. Spaces breaks the script

$invitecsvfile = 'C:\CollabInvite\Import.csv'; # Specify correct path to the csv import file

$invitedUsers = import-csv $invitecsvfile | select-object @{Name='ValidRow';Expression={''}}, UserSourceID, UPN, @{Name='ENTRAID';Expression={''}}, `

                @{Name='ObjectID';Expression={''}}, FirstName, LastName, @{Name='DisplayName';Expression={''}}, email, notes

 

$messageInfo = New-Object Microsoft.Open.MSGraph.Model.InvitedUserMessageInfo

$messageInfo.customizedMessageBody = "Welcome."

$xyzDomain = "xyz.com"

 

foreach ($userInvite in $invitedUsers) {

    # Check to see if of the required columns hav the data, otherwise fail the import for that row.

    if ($userInvite.Email -ne '' -and  $userInvite.FirstName.Trim() -ne '' `

            -and $userInvite.LastName.Trim() -ne '' -and $userInvite.upn.Trim() -ne '' -and $userInvite.UserSourceID.Trim() -eq '2') {

        $userInvite.Email = $userInvite.Email.Trim();

        $userInvite.FirstName = (Get-Culture).TextInfo.ToTitleCase($userInvite.FirstName.ToLower()).Trim();

        $userInvite.LastName = (Get-Culture).TextInfo.ToTitleCase($userInvite.LastName.ToLower()).Trim();

        $userInvite.DisplayName = $userInvite.FirstName + " " + $userInvite.LastName;

        $userInvite.ValidRow = 1;

       

        # Start the invite for the user, no email is actually sent to the user

        New-AzureADMSInvitation -InvitedUserEmailAddress $userInvite.Email `

            -InvitedUserDisplayName $userInvite.DisplayName `

            -InviteRedirectUrl 'https://hometest.xyz.xyz.com' `

            -InvitedUserMessageInfo $messageInfo `

            -SendInvitationMessage $False

    } else {

        $userInvite.ValidRow = 0;

        $userInvite.Notes = "User not invited, data provided is incomplete or not a Collaboration user";

    }

}

 

# Added delay to ensure that the Azure AD users are provisioned completely, sometimes the AAD sync delays the availability of the new accounts

Start-Sleep -Seconds 30

 

foreach ($userInvite in $invitedUsers) {

    # Only process the previously marked rows with complete data, ignore othewise

    If ($userInvite.ValidRow -eq 1) {

       # $xyzzSearchName = "*" + $userInvite.Email.Split("@")[0] + "*onmicrosoft.com*";

        $xyzzSearchName = $userInvite.Email.Split("@")[0];

        # Search for the invited user in AAD using the name part of upn with type:guest, creationtype: invitation, user will have an xyzz tenant domain.

        # If the account does not have this status, it is not an account that should be processed

        # $xyzAADUser = Get-AzureADUser -Filter "UserType eq 'Guest' and CreationType eq 'Invitation'" | Where-Object userPrincipalName -like $xyzzSearchName;

        $xyzAADUser = Get-AzureADUser -Filter "startswith(UserPrincipalName, '$xyzzSearchName')" | Where-Object {$_.usertype -EQ 'Guest' -and $_.creationtype -eq 'invitation'} | Select-Object *;

        $userInvite.ObjectID = $xyzAADUser.ObjectID;

        if ($xyzAADUser -ne $null) {

            $userInvite.ENTRAID = $xyzAADUser.UserPrincipalName.Split("@")[0] + "@" + $xyzDomain;  # Converted UPN to have the xyz.com domain instead of onmicrosoft.com

            # Update the UPN to xyz.com from the xyzz tenant domain, name of the user and the office location for IAM contact email

            Set-AzureADUser -ObjectId $xyzAADUser.UserPrincipalName  `

                -UserPrincipalName $userInvite.ENTRAID `

                -PhysicalDeliveryOfficeName $userInvite.Email `

                -GivenName $userInvite.FirstName `

                -Surname $userInvite.LastName

        } else {    # The user was not found in the xyzz mod tenant, user invitation not processed. If delayed, this row may require reprocessing

            $userInvite.ValidRow = 0;

            $userInvite.Notes = "User has been invited completely or the account is not available yet in xyzz AAD";

        }

    }

}

 

# MUST disconenct from the Azure AD session after making the changes.

Disconnect-AzureAD

 

$invitedUsers | Out-GridView    # displays the list with the status in a grod/ table view popup
