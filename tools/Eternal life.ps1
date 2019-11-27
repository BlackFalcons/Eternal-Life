using namespace System.Management.Automation.Host
clear # Cleaner terminal when run

function Program_Version
{
    # Program version
    Write-Host "$program_title"
}


function Secure_String_To_String
{
    Param(
    [String] $secString
    )
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secString))
}

function Message_Box
{
    Param(
        [String]$Header = "Not defined",
        [String]$Message = "Not defined",
        [String]$Buttons = "Ok",
        [String]$Type = "Information"
    )
    [System.Windows.MessageBox]::Show($Message, $Header, $Buttons, $Type) # Creates message box
}

function Get_User_Information
{
    Param(
        [string[]] $Object_Items = "SamAccountName, Name, ObjectClass, Enabled"
    )
    Write-Host ($AD_User | Format-Table | Out-String)

}


# Password changer
function Password_Changer
{
    $pwd = Read-Host "Password: " -AsSecureString # Request users password
    $pwd_confirm = Read-Host "Re-enter Password: " -AsSecureString # Request user to re-type password
    $pwd_txt = Secure_String_To_String($pwd)
    $pwd_txt_confirm = Secure_String_To_String($pwd_confirm)

    if ($pwd_txt -ceq $pwd_txt_confirm) # Check if both password's match
    {
        Get_User_Information("SamAccountName")
        $AD_User | Set-ADAccountPassword -NewPassword $pwd -Reset
        $AD_User | Set-ADUser -ChangePasswordAtLogon $True
        Message_Box -Message "Password was changed successfully!" -Header "Success" -Buttons "Ok" -Type "Information"
    }
    else
    {
        Message_Box -Message "Password's didn't match!", -Header"Error", -Buttons "Ok", -Type "Error"
    }
}


function Select_New_User
{
    clear
    return 0
}


function New-Menu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Question
    )

    $Change_Password = [ChoiceDescription]::new('&Change Password', "Option will change AD User's password to user selected password.")
    $Get_User_Info = [ChoiceDescription]::new('&User Info', "Option will display data about the selected user.")
    $Select_New_User = [ChoiceDescription]::new("&Select new user", "Option will allow user to select a new user to administrate.")
    $Clear_Terminal = [ChoiceDescription]::new("&Wipe terminal", "Option will ")
    
    
    $options = [ChoiceDescription[]]($Get_User_Info, $change_password, $Select_New_User, $Clear_Terminal)
    $result = $host.ui.PromptForChoice($Title, $Question, $options, 0)

    switch ($result) {
        0 { Get_User_Information }
        1 { Password_Changer }
        2 { return Select_New_User }
        3 { clear; Program_Version }
    }

}


# Program info
$program_version = "1.0"
$program_title = "Eternal life $program_version `n"


# Active Directory Domain and Organizational Unit data
$RYVS_Auto_Elever_OU = "OU=Auto Elever,OU=Brukere,OU=RYVS,OU=Virksomheter"
$BFK_DC = "DC=bfkskole,DC=top,DC=no"


while($True)
{
    Program_Version
    $username = Read-Host -Prompt "Username" # Active Directory username
    $AD_User = Get-ADUser -SearchBase "$RYVS_Auto_Elever_OU,$BFK_DC" -Filter {SamAccountName -eq $username} # AD user selected by $username
    while($True)
    {
        if(!$AD_User)
        {
            Message_Box -Message "No user with the name was found." -Header "Oh dear" -Buttons "Ok" -Type "Warning"
            clear; break;
        }

        Get_User_Information
        
        $Main_Menu = New-Menu -Title $program_title -Question 'Please select option'
        
        
        if($Main_Menu -eq 0)
        {
            break
        }
    }
}