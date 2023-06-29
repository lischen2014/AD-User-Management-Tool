Clear-Host
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0


#####################################################################################
############################# Update History ########################################
#####################################################################################

$UpdateHistory = @"

======================================= Start =======================================

    v0.1.0 =  Initial version.
    v0.1.1 =  Add feature to reset user password.
    v0.1.2 =  Check/Confirm window will pop up before action.
              Modified menu layout.
    v0.1.4 =  Add feature to get IMU application access.
              Show selected user above selection line.
              Optimized menu function.
    v0.1.5 =  Application access result now include IMU user profile.
    v0.1.6 =  Add feature to enable ad user.
              Highlighted id and other important message.
    v0.1.7 =  When search user, the manager will display normally.
    v0.1.8 =  When search returns multiple user, you can specify a user to proceed.
              Remove previous selected user and history when run this script.
              Removed GST user from search result.
              Add OU into user detail information.
              Optimized tip message.
    v0.1.9 =  Now you can search a user in main menu directly.
    v0.2.0 =  Add switch for GST account search.
              Add change language feature for Office 365 Webpage
    v0.2.1 =  Indented code to improve readability.
              Add back gst account search with Network id or uid. also password reset.
    v0.2.2 =  Password now will have random words from a list.
    v0.2.3 =  Removed mobile and telephone number from result.
              Add feature to search with Email Address.
              Add some name for Get-AdUser result filter.
    v0.2.4 =  Add Manager Email in detail information.
              Removed language change feature, add disable password change after logon.
    v0.2.5 =  Add CDS user group information, changed layout.
    v0.2.6 =  Highlight VPN access.
              Add O365 License into brief info.
    v0.2.7 =  Add Update History into main menu.
    v0.2.8 =  Add Password Expiry date and Password Expire Alert.
    v0.2.9 =  Fixed VPN access shows wrong result, changed menu option.
    v0.3.0 =  Fixed bug unable to enable AD User.
              Modified multiple color output code.
    v0.3.1 =  Chnaged default color to ui default color.
    v0.3.2 =  Fixed a bug to display color.
    v0.3.3 =  Fixed a bug with color display.
              Fixed a bug with enable AD Account, result changes after enable done.
              changed prompt confirm message.
    v0.3.4 =  Fixed bug when with id selection.
              Fixed a bug when go to option 1, the no user match message returned wrong result.
    v0.3.5 =  Add feature to change user AD language.
    v0.3.6 =  Optimized language change message.
    v0.3.7 =  Optimized message showing in wrong position.
    v0.3.8 =  Text modification.
              Fixed a bug in password.
              Modified the count method of gst account.
              Enhanced password complexity with new AD password policy.
    v0.3.9 =  Reduce password complexity for regular accounts.
    v0.4.0 =  Changed Script Name to from 'Get ADUser Tool' to 'AD User Management Tool for ITSD'.
              Add Version Number in Menu Title.
              Modified Update Log Display.
    v0.4.1 =  Add User Computer in main menu.
    v0.4.2 =  More words in password dictionary.
              Fixed bug about cannot disable password change after logon.
    v0.4.3 =  Add feature to unlock AD Account.
    v0.4.4 =  Modified syntax.
              Highlighted MFA Phone.
    v0.4.5 =  Fixed version display bug.
    v0.4.6 =  Display MFA Phone on search result page.
    v0.4.7 =  Modified prompt code.
    v0.4.8 =  Added detection of AD Module
              Modified code, enhanced performance and stability.
    v0.4.9 =  Add feature of password expired alert.
              Fixed a bug in version display.
    v0.5.0 =  Fixed a bug of password reset.
              Fixed a bug of prompt confirm.
    v0.5.1 =  Modified feature of enable ad user.
              Modified code with search result.
    v0.5.2 =  Modified some code.
              Add more words in password dictionary.
              Fixed a bug in enable AD user.
    v0.5.3 =  Fixed a bug in enable AD user,
              Add Easy Password switch.
    v0.5.4 =  Fixed a bug of password expiry alert.
              change easy password to default, if need more complexed password, please change `$EasyPasswordEnabled to `$false.
    v0.5.5 =  Modified Code with app access filter.
    v0.5.6 =  Removed duplicate password expiry alert.
              Fixed a bug in disable change user password.
              Text modification.
    v0.5.7 =  Changed version ID, now started with 0.1
              Modify version check function.
    v0.5.8 =  Fixed a bug when search user.
    v0.5.9 =  Fixed a bug which cannot find correct Computer Name.
    v0.6.0 =  Fixed a bug of detect user selection.
              Text modification.
    v0.6.1 =  Fixed a bug of display user computer.
    v0.6.2 =  Added a feature to display account expiry.
              Fixed a bug of display user computer. 
    v0.6.3 =  Added a feature to send data to collect usage.
    v0.6.4 =  Modified Code.
              Fixed wrong error message.
    v0.6.5 =  Fixed a bug in enable AD User.
    v0.6.6 =  Add a feature to exclude service account from password reset.
    v0.6.7 =  Modified text in change Office language.
    v0.6.8 =  Modified layout, added account locked out alert.
    v0.6.9 =  Fixed a bug that did not recognize AD lock successfully in specific situation.
    v0.7.0 =  Fixed a bug in password reset.
    v0.7.1 =  Added a feature to distinguish User temporarily lock out and manualy lock out.
    v0.7.2 =  Fixed a bug when reset password for anonymous account.
              Fixed a bug that unable to quit with Q keyword.
              Text modification.
    v0.7.3 =  Highlighted VPN access group.
              Text modification.
              Now it can display multiple VPN access.
    v0.7.4 =  Modified layout when display multiple VPN access.
    v0.7.5 =  Added PAM accounts reset notification.
    v0.7.6 =  Fixed a bug that unable to unlock user when account is locked in AD.

    Latest Releases: 
    https://github.com/lischen2014/AD-User-Management-Tool

                                                   Author: Leon Jiang
                                                   Email: linchen2014@foxmail.com
======================================== End ========================================

"@


#####################################################################################
################################ Functions ##########################################
#####################################################################################

function Show-Menu {
    param (
    [string]$Title = "AD User Management Tool $Current_Version"
    )
    Write-Host "=============== $Title ==============="
    Write-Host ""
    Write-Host "1: Press '1' to select/select user."
    Write-Host "2: Press '2' to view user detail information."
    Write-Host "3: Press '3' to view user access."
    Write-Host "4: Press '4' to enable user in AD."
    Write-Host "5: Press '5' to reset user AD password."
    Write-Host "6: Press '6' to disable password expiry after password reset."
    Write-Host "7: Press '7' to unlock account"
    Write-Host "8: Press '8' to change Office web language."
    Write-Host "9: Press '9' to view update log."
    Write-Host "Q: Press 'Q' to exit this script."
}


function Start-Menu {
    do
    {
        Write-Host ""
        Show-Menu
        Write-Host ""
        if (!$uidinfo){
            Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
            Write-Host ""
        }
        else{
            Write-Msg @("Important: ", "You are now selected user ","$id",".") -Color @("Red",$defaultcolor,"Red",$defaultcolor)
            Write-Host ""
        }
        $selection = Read-Host "Select an user/option"
    
        switch($selection)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               {
        
        '1'{
            $id = Input-Id
            Get-User -filter (Build-Filter -id $id)
        }
        
        '2'{
            if(!$uidinfo){
                Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
                Write-Host ""
            }
            else{
                Write-Host ($ustdinfo | Out-String)
            }
        }
        
        '3'{
            if (!$uidinfo){
                Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
                Write-Host ""
            }
            else{
                Write-Host ""
                Write-Host "User have below application access: "
                Write-Host ""
                Get-Group
            }
            
        }
        
        '4'{
            if (($uidinfo) -and ($ustdinfo.("AD Enabled") -eq $False)){
                [System.Windows.Forms.MessageBox]::Show("This option will only enable user in AD!!!")
                $decision = Prompt-Confirm -action "Enable AD account" -id $id -prompt $true
                if ($decision -eq 0){
                    Write-Host "Please wait a second..."
                    Enable-Account -filter $filter
                    Write-Msg ("Message: Account of", " $id ", "enabled successfully.") -Color @("Cyan","Red","Cyan")
                    Write-Host ""
            }}
            elseif($false -eq $filter){
                Write-Host "No user selected" -f Red
            }
            elseif(($false -ne $uidinfo) -and ($ustdinfo.("AD Enabled") -eq $True)){
                Write-Host ""
                Write-Msg @("Message: Account of ","$id"," is already enabled, please double confirm if id is correct.") -Color @($defaultcolor, "Red", $defaultcolor)
                Write-Host ""
            }
            else{
                Write-Host ""
                Write-Msg @("Message: ", "Enable Account failed, please double check if this id is correct: "," '$id'") -Color @($defaultcolor, "Red", $defaultcolor)
                Write-Host ""
            }
        }
        
        '5'{
            if(!$uidinfo){
                Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
                Write-Host ""
            }
            else{
                if ($false -ne $uidinfo){
                    $decision = Prompt-Confirm -action "Password Reset" -id $id -prompt $true
                    if ($decision -eq 0){
                        Set-Password $filter
                }
                }
                elseif($false -eq $filter){
                    Write-Host "No user selected" -f Red
                    Build-Filter -id (Input-Id)
                    Write-Host ""
                }
                else{
                    Write-Host ""
                    Write-Warning "Update password failed, please double check if user id is correct and active: '$id'"
                    Write-Host ""
                }
            } 
        }
        
        '6'{
            if(!$uidinfo){
                Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
                Write-Host ""
            }
            else{
                Disable-PasswdAtLogon -filter $filter
            }
        }

        '7'{
            if(!$uidinfo){
                Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
                Write-Host ""
            }
            else{
                Unlock-AD
            }

        }

        '8'{
            if(!$uidinfo){
                Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
                Write-Host ""
            }
            else{
                Write-Output $languagelist
                Write-Host ""
                $currentlanguage = (Get-AdUser -LDAPfilter $filter -Properties PreferredLanguage).PreferredLanguage
                Write-Msg ("User ", $id, " currently language is ", $currentlanguage, ".") -Color @($defaultcolor,"Red",$defaultcolor,"Red", $defaultcolor)
                Write-Host ""
                Write-Msg ("Note: Please type ", "Quit"," or ", "Exit", " if you want back to main menu.") -Color @($defaultcolor,"Red",$defaultcolor,"Red", $defaultcolor)
                Write-Host ""
                $language = (read-host "Please input the Language Culture Name").trim()

                if ($languagelist.keys -contains $language){
                    $language = $languagelist.$language
                    Write-Host "'$language' is selected, changing user language"
                    Set-Language -filter $filter
                    Write-Msg @("The language of user ",$id, " has been changed to ", $language, "`nPlease wait up to 48 hours to take effect.") -Color @($defaultcolor,"red", $defaultcolor, "red", $defaultcolor)
                }
                elseif($languagelist.values -contains $language){
                    Write-Host "'$language' is selected, changing user language"
                    Set-Language -filter $filter
                    Write-Msg @("The language of user ",$id, " has been changed to ", $language, "`nPlease wait up to 48 hours to take effect.") -Color @($defaultcolor,"red", $defaultcolor, "red", $defaultcolor)
                }
                elseif(($language -eq "quit") -or ($language -eq "exit")){
                    Write-Host "Back to main menu"
                }
                else{
                    Write-Host ""
                    Write-Host "Language '$language' is not exist, back to main menu" -f red
                }
                Write-Host ""
            }

        }

        '9'{
            $UpdateHistory
            pause

        }

        'Q'{
            $decision = Prompt-Confirm -action "exit this tool?" -prompt $true
            if ($decision -eq 0){
                exit
            }
        }

        default {
            Write-Host ""
            if ($selection.length -gt 1){
                $id = ($selection).trim()
                Get-User -filter (Build-Filter -id $id)
            }
            else{
                Write-Msg @("Important: ", "No user selected, ","input a User ID to continue.") -Color @("Red", "Red", $defaultcolor)
                Write-Host ""
            }
            }
        }
    }
    until($UserInput -eq 'Q')
}


function Input-Id {
    Write-Host ""
    Write-Msg @("Note: To search ","gst"," account, do not use Display Name") -Color @($defaultcolor, "red", $defaultcolor)
    $id = (read-host "Please input either Networkid, Uid or Full Display Name").trim()
    return $id
}


function Build-Filter {
    param($id)
    If ($id.Contains("@")){
        Write-Host "You typed Email Address: $id"
        $idtmp = $id.Substring(0, $id.IndexOf('@'))
        $script:filter = "(|(CN=$idtmp))"
        $script:exclude = "XXX"
    }
    elseif($id.Contains(".")){
        Write-Host "You typed uid: $id"
        $script:filter = "(|(CN=$id))"
        $script:exclude = "XXX"
    }
    elseif($id.Contains(" ")){
        Write-Host "You typed display name: $id"
        $script:filter = "(|(displayname=$id*))"
        $script:exclude = "gst.*"
        
    }
    elseif($id.Contains(" ")){
        Write-Host "You typed Supplier Number: $id"
        $script:filter = "(|(Office=SUP $id))"
        $script:exclude = "gst.*"
    }

    else{
        Write-Host "You typed network id: $id"
        $script:filter = "(|(SamAccountName=$id))"
        $script:exclude = "XXX"
    }
    
}


function Get-User {

    $users = Get-AdUser -LDAPFilter $filter | ForEach-Object {
        Get-AdUser -Identity $_ -Properties "msDS-UserPasswordExpiryTimeComputed", 'pwdLastSet',*
    }

    foreach ($item in $exclude) 
    {
        
          $Users = $Users | Where-Object {$_.CN -notlike "*$item*"} 
    }
    
    $script:userfullinfo = ($users)
    $script:uidinfo = $userfullinfo| Select-Object $idfilter
    $script:upn = $uidinfo.EmailAddress
    $script:idcount = ($uidinfo.("Network ID")).count
    Write-Host ($uidinfo | Out-String)

    if ($idcount -gt 1){
        Write-Host "Please sepcify a user to proceed"
        $script:selectid = (read-host "Which user you would like to select?(network id)").trim()
        $script:id = $selectid
        $script:filter = "(|(SamAccountName=$selectid))"
        Write-Msg @("You selected user Network id ","$id") -Color @($defaultcolor,"Red")
        Get-User
    }
    elseif($idcount -eq 1){
        # no action needed
    }
    elseif($idcount -eq 0){
        Write-Msg @("Important: ", "Search of ","'$id' returned no result."," Please check if id is correct.") -Color @("Red",$defaultcolor, "Red",$defaultcolor)
        Write-Host ""
    }
    else{
        Write-Host "you did not specify a user id" -f Red
        Write-Host ""
    }

    $script:ustdinfo = $userfullinfo | Select-Object $stdfilter
    $script:uappaccessinfo = $userfullinfo | Select-Object -ExpandProperty memberOf
    
    
    if ($ustdinfo){
        # AD Disabled Alert
        if(!$ustdinfo.("AD Enabled")){
            Write-Warning "User is disabled or deactivated."
        }
        else{
            Write-Host "Message: User is active in in AD." -f Green
        }
        
        # Accolunt Locked Out Alert
        if($userfullinfo.AccountLockoutTime){
            Write-Warning "User is temporarily locked out."
        }
        else{
            Write-Host "Message: User is not being locked out."-f Green
        }

        # Password Expired Alert
        if($ustdinfo.PasswordExpired){
            Write-Warning "User password is Expired."
        }
        else{
            Write-Host "Message: User password is not expired."-f Green
        }
        
        # AD Account Expired Alert
        if($ustdinfo.AccountExpiredinAD -eq 'True'){
            Write-Warning "User account is Expired in AD."
        }
        else{
            Write-Host "Message: User account is not expired."-f Green
        }

        # Change Password at Next Logon Alert
        if($userfullinfo.pwdLastSet -eq '0'){
            Write-Warning "User need change password at next login."
            Write-Host "Note: use '6' when user keeps prompt to change a new password."
        }
        else{
            # No action needed
        }

    }
    Write-Host ""
    break # Not quite sure why it worked, but if removed the break, when search user with display name(which will return multiple result) then specify a user, it will gives double AD and Passowrd Message.
}


function Write-Msg {
     # This function should only use when you need display multiple colors.
     param(
         $Text,
         $Color
     )

     foreach($i in 0..($Text.length-1)){
        if ($false -eq $Text){
            Write-Host ""
        }
        else{
                if ($Color -is [String]){
                    Write-Host $Text[$i] -ForegroundColor $Color
                }
                elseif ($Color -is [Array]){
                    Write-Host $Text[$i] -NoNewline -ForegroundColor $Color[$i]
                }
        }
     }
}


function Get-LatestVersionId {
    
    $pattern = 'v(\d+(\.\d+){0,3})'
    $matches = [regex]::Matches($UpdateHistory, $pattern)
    
    if ($matches.Count -gt 0) {
        $latestMatch = $matches[$matches.Count - 1]
        return $latestMatch.Groups[0].Value
    }
    
    return $null
}

function Prompt-Confirm {
    # Version: 1.2
    # Author: Leon Jiang

    param(
        [Parameter(mandatory=$false)]
        $action = "continue this action",
        [Parameter(mandatory=$false)]
        $id,
        [Parameter(mandatory=$false)]
        [Bool]$prompt
    )

    # Determine prompt or not
    if($prompt){
        # Return value: 0 for True, 1 for False

        $title = "Confirm '$action'"
    
        if($id){
            $question = "Do you want to '$action' for '$id'?"
        }
        else{
            $question = "Do you want to '$action'?"
        }

        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
    
        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        # or we can add global variable then remove the return line.
        # $global:decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        if ($decision -eq 0) {
            Write-Host ' Confirmed'
        }
        else{
            Write-Host ' Cancelled'
        }
    }
    else{
        # Return value: 0 for True, 1 for False
        Do {
            $decision = Read-Host -Prompt "Do you want to '$action'? (y/n)"
        }
        Until ($decision -eq 'y' -or $decision -eq 'n')
    }

    return $decision
}


function Get-Group {
    $UserProfile=@()
    $AppAccess=@()
    $UserGroup=@()
    $vpn = @()

    foreach($i in $uappaccessinfo){
    
        if($i -like "*OU=UserProfiles*"){
            # get user profile
            $index = $i.IndexOf("@")
            $UserProfileTmp = $i.Substring(6, ($index-9))
            $UserProfile += $UserProfileTmp

        }
        elseif($i -like "*OU=Applications*"){
            # get application access
            $AppAccessTmp = $i -replace "CN=A-|,?(OU|DC)=[^,]+", ""
            $AppAccess += $AppAccessTmp
        }
        if($i -like "*OU=UserGroups*"){
            # get user group
            $UserGroupTmp = ($i -replace ",OU=UserGroups.*","")
            $UserGroupTmp = $UserGroupTmp.substring(3)
                if(($UserGroupTmp -like "*RAS*") -or ($UserGroupTmp -like "*VPN*")){
                    $vpn += $UserGroupTmp
                }
            $UserGroup += $UserGroupTmp
        }
     
    }

    Write-Host "====== User Profile ======" -ForegroundColor Cyan
    Write-Output $UserProfile | Sort-Object
    Write-Host "=== Application Access ===" -ForegroundColor Cyan
    Write-Output $AppAccess | Sort-Object
    Write-Host "======= User Group =======" -ForegroundColor Cyan
    Write-Output $UserGroup | Sort-Object

    if($vpn){
        Write-Host "`nUser has VPN access:"-ForegroundColor Cyan
        foreach($i in $vpn) {
            Write-Host $i -ForegroundColor Cyan
        }
    }
    else{
        Write-Host "User do not have VPN access."-ForegroundColor Red
    }
}


Function Set-Language {
    param($filter)
    # "https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20)?redirectedfrom=MSDN"
    Get-AdUser -LDAPfilter $filter -Properties PreferredLanguage | Set-ADUser -Replace @{PreferredLanguage = "$language"}

}


function Set-Password {
    param($filter)
    
    $date = Get-Date -format HHmm
    if($uidinfo.("NetWork ID") -match "-"){
        # Service account
        Write-Warning "Currently unable to reset password for service account,`nPlease go to Identity Management Portal to proceed."
    }
    else{
        if(!$EasyPasswordEnabled){
            Write-Host "Please wait a second..."
            $randpasswdindex = (Get-random -Maximum  $wordlist1.count -SetSeed ([Int32](Get-Date).Ticks.ToString().substring(10)))
            $randpasswdindex2 = (Get-random -Maximum  $wordlist2.count -SetSeed ([Int32](Get-Date).Ticks.ToString().substring(13)))
            $randpasswdindex3 = (Get-random -Maximum  $wordlist3.count -SetSeed ([Int32](Get-Date).Ticks.ToString().substring(15)))
            if($uidinfo.("NetWork ID") -match "gst"){
                # GST
                $Password = ("$($wordlist1[$randpasswdindex])" + "$($wordlist2[$randpasswdindex2])" + "$($wordlist3[$randpasswdindex3])" + "$date" + "!")

            }
            else{
                # Normal account
                $Password = ("$($wordlist1[$randpasswdindex])" + "$($wordlist2[$randpasswdindex2])" + "$date")

            }
        }
        else{
            # Easy Password
            $longdate = get-date -format yyyyMMdd
            if($uidinfo.("NetWork ID") -match "gst"){
                # GST
                $Password = "Welcome" + "$longdate" + "!"
            }
            else{
                # Normal account
                $Password = "Welcome" + "$longdate"
            }
        }
        try{
            Get-AdUser -LDAPfilter $filter | Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$password" -Force)
            Write-Host "The password of user id '$id' has been reset to " -NoNewline
            Write-Host "'$password'" -ForegroundColor Red
            Write-Host "Password Reset complete"
        }
        catch{
            write-host ""
            $e = $_.Exception.Message 
            write-msg -Text @("Password Reset Failed, Error Message: ", $e) -Color @("Red", "Red")
            write-host "`nPlease double check account status."
            if($uidinfo.("NetWork ID") -match "gst"){
                # GST
                write-host "`nFor gst account, please guide user to reset password from PAM Portal" -f "Yellow"
            }
            write-host ""
        }
    }
}


function Enable-Account {
    param(
        [Parameter(Position=0,mandatory=$true)]
        [string] $filter
        )

    if(!$ustdinfo.("AD Enabled")){
        # Write-Warning "User is disabled in AD."
        try{
            Get-AdUser -LDAPFilter $filter | Enable-ADAccount  
        }
        catch{
            Write-warning $_.Exception.Message
        }
    }
    else{
        Write-Host "Message: User is active in AD, no need to enable."
    }
}


function Disable-PasswdAtLogon {
    param($filter)

    if ($userfullinfo.pwdLastSet -eq '0'){
        Write-Host "User need change password after logon, do you want to disable this option?"
        $decision = Prompt-Confirm -action "Disable password reset after logon" -id $id -prompt $true
        if ($decision -eq 0){
            Get-AdUser -LDAPfilter $filter | Set-ADUser -ChangePasswordAtLogon $false
            Write-Host "Change password at next logon has been disabled." -f Cyan
        }
    }
    else{
        Write-Host ""
        Write-Host "Change Password After Logon is not enabled, no action needed."
    }
    
    
}


function Unlock-AD {
                
                if($false -eq $uidinfo){
                    Write-Host "No user selected" -f Red
                }
                else{
                    if(((Get-AdUser -LDAPFilter $filter -properties LockedOut).LockedOut -eq $True) -or ($userfullinfo.AccountLockoutTime)){
                        Write-Host ""
                        Write-Msg ("User ", $id, " is locked out, ", "do you want to unlock this account?") -Color @($defaultcolor,"Red",$defaultcolor,$defaultcolor)
                        $decision = Prompt-Confirm -action 'Unlock account' -id $id -prompt $true
                        if ($decision -eq 0){
                            Get-AdUser -LDAPFilter $filter | Unlock-ADAccount
                            Write-Msg ("User ", $id, " is now unlocked!") -Color @($defaultcolor,"Red",$defaultcolor)
                        }
                        else{
                            Write-Host "Cancelled."
                        }
                    }
                    else{
                        Write-Host ""
                        Write-Host "Account is not locked in AD."
                    }
                    Write-Host ""
                }
            }

function Search-ComputerName {
    param($list)

    $result = @()
    $RegExComputerNumber = "CN=([A-Za-z0-9]{7,8}-[A-Za-z]{2}\d{4})"

    if ($list){
        foreach ($i in $list){
            if ($i -match $RegExComputerNumber){
                $result += ($i.Split(",")[0].substring(3))
            }
        }
    }
    return $result[0]
}

function Send-Stat{

    $counterfile = "\\Hostname\counter.txt"
    $fileName = Split-Path $MyInvocation.ScriptName -Leaf
    [string]$agent = [string]($env:UserName).ToLower()
    $time = get-date

    try{
        $usage = Get-Content $counterfile -ErrorAction SilentlyContinue
        $newline = "$time,$agent,$fileName"
        $newline | add-content -path $counterfile -ErrorAction SilentlyContinue
    }
    catch{}
}

#####################################################################################
############################## Variables ############################################
#####################################################################################
[Bool]$EasyPasswordEnabled = $true
$Current_Version = Get-LatestVersionId
$date = Get-Date
# Send-Stat
$defaultcolor = (Get-Host).ui.rawui.ForegroundColor

if ($defaultcolor -eq '-1'){
    $defaultcolor = [ConsoleColor]"White"
}

$wordlist1 = @(
    "Drum"
    "Book"
    "Plug"
    "Score"
    "Excel"
    "Flow"
    "Lake"
    "Ruby"
    "Thor"
    "People"
    "Moment"
)

$wordlist2 = @(
    "Light"
    "Bless"
    "Flame"
    "Ground"
    "Charge"
    "Advice"
    "Soccer"
    "Bottle"
    "Cloud"
    "Polo"
    "Track"
    "Drink"
)

$wordlist3 = @(
    "Ocean"
    "Polo"
    "Below"
    "Track"
    "Drink"
    "Flight"
    "Notion"
    "Gusto"
    "Plug"
)

$languagelist = [ordered]@{
    "English"="en-US";
    "Chinese"="zh-CN";
    "Swedish"="sv-SE";
    "Spanish"="es-ES";
    "Japanese"="ja-JP";
    "Korean"="ko-KR";
    "Arabic"="ar-SA"
    "Hebrew"="he-IL";
}

$stdfilter = @(
    @{Name = "Network ID";Expression = {$_.SamAccountName}}
    "CN"
    "Name"
    "DisplayName"
    "UserPrincipalName"
    "EmailAddress"
    "EmployeeID"
    @{Name = "Country Code";Expression = {$_.Country}}
    @{Name = "Country Name";Expression = {$_.co}}
    "Company"
    "Office"
    "Department"
    "Title"
    @{Name = "Manager";Expression = {ForEach-Object{(Get-AdUser $_.Manager -Properties DisplayName).DisplayName}}}
    @{Name = "ManagerEmailAddress";Expression = {ForEach-Object{(Get-AdUser $_.Manager -Properties EmailAddress).EmailAddress}}}
    @{Name = "Mobile Phone(MFA)";Expression = {$_.MobilePhone}}
    "OfficePhone"
    @{Name = "Microsoft365License";Expression = {$_.extensionAttribute2}}
    "preferredLanguage"
    @{Name = "AD Enabled";Expression = {$_.Enabled}}
    @{Name = "UserComputer";Expression = {
        Search-ComputerName -list $_.directReports
    }}
    "whenCreated"
    "PasswordLastSet"
    @{Name="PasswordExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}
    "PasswordExpired"
    @{Name = "AccountExpiredinAD";Expression = {[Bool]$_.AccountExpirationDate}}
    "AccountExpirationDate"
    @{Name="Account LockedOut";Expression={[datetime]::FromFileTime($_."LockedOut")}}
    @{Name="Account LockoutTime";Expression={[datetime]::FromFileTime($_."LockoutTime")}}
    @{Name = "UDrive";Expression = {$_.HomeDrive}}
    @{Name = "UDrive Path";Expression = {$_.HomeDirectory}}
    @{Name = "AD Path";Expression = {$_.DistinguishedName}}
)

$idfilter = @(
    @{Name = "Network ID";Expression = {$_.SamAccountName}}
    "DisplayName"
    "EmailAddress"
    "Title"
    @{Name = "Manager";Expression = {ForEach-Object{(Get-AdUser $_.Manager -Properties DisplayName).DisplayName}}}
    @{Name = "Microsoft365License";Expression = {$_.extensionAttribute2}}
    @{Name = "Mobile Phone(MFA)";Expression = {$_.MobilePhone}}
    @{Name="PasswordExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}
    # "PasswordExpired"
    # @{Name = "AccountExpiredinAD";Expression = {[Bool]$_.AccountExpirationDate}}
    @{Name = "UserComputer";Expression = {
        Search-ComputerName -list $_.directReports
    }}
)


#####################################################################################
###################################### Main #########################################
#####################################################################################

Import-Module ActiveDirectory -ErrorAction Stop

try {
    Start-Menu
}
Catch {
    write-host ""
    Write-Warning "Oops,seems an error occurred, please try download the latest version of script.`n"
    write-host "You can download latest version from here: `nhttps://github.com/lischen2014/AD-User-Management-Tool`n"
    $decision = Prompt-Confirm -action 'show error details?'
    if ($decision -eq 'y'){
        $ErrLine = $_.ScriptStackTrace
        $ErrType =  $($Error[0]).Exception.GetType().FullName
        $ErrDetail = $_.Exception.Message
        Write-Warning "`nLocation: `n$ErrLine`nType: `n$ErrType`nDetail:`n$ErrDetail"
    }
}

