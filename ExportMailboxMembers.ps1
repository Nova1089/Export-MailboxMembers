<#
Allows you to search for a mailbox and export all members with access to that mailbox.
Does not work with Powershell 7.
#>

# Defining functions
Function UseModule([string]$moduleName)
{
    while ((TestModuleInstalled($moduleName)) -eq $false)
    {
        PromptToInstallModule($moduleName)
        TestSessionPrivileges
        Install-Module $moduleName
        if ((TestModuleInstalled($moduleName)) -eq $true)
        {
            Write-Host "Importing module..."
            Import-Module $moduleName
        }
        else
        {
            continue
        }
    }
}

Function TestModuleInstalled([string]$moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($module -ne $null)
}

Function PromptToInstallModule([string]$moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required."
        $confirmInstall = Read-Host -Prompt "Would you like to install it? (y/n)"
    }
    while ($confirmInstall -notmatch "\b[yY]\b")
}

Function TestSessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Throw "Please run script with admin privileges. 
            1. Open Powershell as admin.
            2. CD into script directory.
            3. Run .\scriptname.ps1"
    }
}

Function ConnectToExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($connectionStatus -eq $null)
    {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($connectionStatus -eq $null)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again."
        }
    }
}

Function ConnectToOffice365
{
    Get-MsolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null

    while ($errorConnecting -ne $null)
    {
        Write-Host "Connecting to Office 365..."
        Connect-MsolService -ErrorAction SilentlyContinue
        Get-MSolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null   

        if ($errorConnecting -ne $null)
        {
            Read-Host -Prompt "Failed to connect to Office 365. Press Enter to try again."
        }
    }
}

Function PromptAndExport
{
    do
    {
        $Table.Clear()
        $mailbox = PromptForMailbox
        CreateMemberTable($mailbox)
        ExportReport($mailbox)
        do
        {
            $goAgain = Read-Host -Prompt "Would you like to perform another export? (y/n)"
        }
        while ($goAgain -notmatch '\b[yYnN]\b')
    }
    while ($goAgain -match '\b[yY]\b')
}

Function PromptForMailbox
{
    do
    {
        $mailboxPrompt = Read-Host -Prompt "Enter name or email of mailbox."
        $mailbox = Get-EXOMailbox -Identity $mailboxPrompt -ErrorAction SilentlyContinue

        if ($mailbox -eq $null)
        {
            Write-Host "Mailbox not found."
        }
    }
    while ($mailbox -eq $null)
    
    Write-Host "Mailbox found. Display Name: $($mailbox.DisplayName), Email: $($mailbox.UserPrincipalName), Type: $($mailbox.RecipientTypeDetails) `n"
    Write-Host "Exporting members... `n"

    return $mailbox
}

Function CreateMemberTable([PSObject]$mailbox)
{
    $membersReadManage = GetMembersWithReadManage($mailbox)
    $membersSendAs = GetMembersWithSendAs($mailbox)

    foreach ($member in $membersReadManage)
    {
        $upn = $member.User
        if ($upn -match '[\w\.]+@[\w\.]+')
        {
            $userInfo = GetUserInfo($upn)
            $memberInfo = CreateMemberInfoObject -Mailbox $mailbox -Member $member -UserInfo $userInfo
        }
        else
        {
            $memberInfo = CreateMemberInfoObject -Mailbox $mailbox -Member $member
        }
        AddToMemberTable($memberInfo)
    }
    foreach ($member in $membersSendAs)
    {
        if(($Table.ContainsKey($member.Trustee)))
        {
            $Table[$member.Trustee].SendAsAccess = "Y"
        }
        else
        {
            $upn = $member.Trustee
            if ($upn -match '[\w\.]+@[\w\.]+')
            {
                $userInfo = GetUserInfo($upn)
                $memberInfo = CreateMemberInfoObject -Mailbox $mailbox -Member $member -UserInfo $userInfo
            }
            else
            {
                $memberInfo = CreateMemberInfoObject -Mailbox $mailbox -Member $member
            }
            AddToMemberTable($memberInfo)
        }
    }    
}

Function GetMembersWithReadManage([object]$mailbox)
{
    return (Get-EXOMailboxPermission -UserPrincipalName ($mailbox.UserPrincipalName) |
            Where-Object {$_.User -ne "NT AUTHORITY\SELF"})
}

Function GetMembersWithSendAs([object]$mailbox)
{
    return (Get-EXORecipientPermission -UserPrincipalName ($mailbox.UserPrincipalName) |
            Where-Object {$_.Trustee -ne "NT AUTHORITY\SELF"})
}

Function GetUserInfo($upn)
{
    return (Get-MsolUser -UserPrincipalName $upn)
}

Function CreateMemberInfoObject([PSObject]$mailbox, [object]$member, [object]$userInfo)
{
    if ($member -isnot [Microsoft.Exchange.Management.RestApiClient.ExoRecipientPermission])
    {
        $memberInfo = [PSCustomObject]@{
            MBDisplayName    = $mailbox.DisplayName
            MBEmail          = $mailbox.UserPrincipalName
            MBType           = $mailbox.RecipientTypeDetails
            UserEmail        = $member.User
            UserTitle        = $userInfo.Title
            UserDepartment   = $userInfo.Department
            ReadManageAccess = "Y"
            SendAsAccess     = "N"
        }        
    }
    else
    {
        $memberInfo = [PSCustomObject]@{
            MBDisplayName    = $mailbox.DisplayName
            MBEmail          = $mailbox.UserPrincipalName
            MBType           = $mailbox.RecipientTypeDetails
            UserEmail        = $member.Trustee
            UserTitle        = $userInfo.Title
            UserDepartment   = $userInfo.Department
            ReadManageAccess = "N"
            SendAsAccess     = "Y"
        }        
    }
    return $memberInfo
}

Function AddToMemberTable([PSObject]$memberInfo)
{
    if ($null -eq $memberInfo) {return}
    $Table.Add($memberInfo.UserEmail, $memberInfo)
}

Function ExportReport($mailbox)
{
    $path = NewPath($mailbox)
    $Table.Values | Export-CSV $path -NoTypeInformation
    Write-Host "Finished exporting to $path. `n"
}

Function NewPath($mailbox)
{
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $mailboxName = $mailbox.DisplayName
    $timeStamp = NewTimeStamp
    return "$desktopPath\$mailboxName Mailbox Members $timeStamp.csv"
}

Function NewTimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
}

# Main
$script:Table = @{}

UseModule("ExchangeOnlineManagement")
UseModule("MSOnline")
ConnectToExchangeOnline
ConnectToOffice365
PromptAndExport
Read-Host -Prompt "Press Enter to exit"