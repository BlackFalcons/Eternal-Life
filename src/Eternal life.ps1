using namespace System.Management.Automation.Host
using namespace System.Reflection
using namespace System.Windows.Forms
using namespace System.Runtime.InteropServices


# Add support for message boxes to shells that does not have them supported.
[void][Assembly]::LoadWithPartialName("System.Drawing")
[void][Assembly]::LoadWithPartialName("System.Windows.Forms")


function Program_Version
{
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Program_Title
    )
    # Program version
    Write-Host $Program_Title
}


function Message_Box
{
    Param(
        [String]$Header = "Not defined",
        [String]$Message = "Not defined",
        [String]$Buttons = "Ok",
        [String]$Type = "Information"
    )
    [MessageBox]::Show($Message, $Header, $Buttons, $Type) # Creates message box
}


function Get_User_Information
{
    if($Dev_Mode)
    {
        Write-Host "Developer mode is active`n" -ForegroundColor Magenta
    }

    if(!$AD_User.Enabled)
    {
        Write-Host "Warning, this user is locked!" -ForegroundColor Red
    }


    if($AD_User.SamAccountName)
    {
        Write-Host "Selected user:" $AD_User.SamAccountName
    }

    if($AD_User.EmployeeID)
    {
        Write-Host "Employee ID: " $AD_User.EmployeeID
    } else {
        Write-Host "No employee ID found for the selected user."
    }
}


# Password changer
function Password_Changer
{
    $pwd = Read-Host "Password: " -AsSecureString # Request users password
    $pwd_confirm = Read-Host "Re-enter Password: " -AsSecureString # Request user to re-type password
    $pwd_txt = [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($pwd))
    $pwd_txt_confirm = [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($pwd_confirm))

    if ($pwd_txt -ceq $pwd_txt_confirm) # Check if both password's match
    {
        Get_User_Information
        $AD_User | Set-ADAccountPassword -NewPassword $pwd -Reset
        $AD_User | Set-ADUser -ChangePasswordAtLogon $True
        Message_Box -Message "Password was changed successfully!" -Header "Success" -Buttons "Ok" -Type "Information"
    }
    else
    {
        Message_Box -Message "Password's didn't match!" -Header "Error" -Buttons "Ok" -Type "Error"
    }
}


function Select_New_User
{
    Clear-Host
    return $False
}


function New_Menu_Item
{
	Param(
        [Parameter(Mandatory)]
		[String]$Option,
        
        [String]$Tutorial = "No information avilable."
	)
	[ChoiceDescription]::new($Option, $Tutorial)
}


function New_Menu {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$Title,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$Question
    )

    $Change_Password = New_Menu_Item -Option "&Change Password" -Tutorial "Option will change AD User's password to user selected password."
    $Get_Bitlocker_Info = New_Menu_Item -Option '&Bitlocker Recovery' -Tutorial "Option will display bitlocker recovery information about the selected AD Computer."
    $Select_New_User = New_Menu_Item -Option "&Select new user" -Tutorial "Option will allow user to select a new user to administrate."


    $options = [ChoiceDescription[]]($Get_Bitlocker_Info, $change_password, $Select_New_User)
    $result = $host.ui.PromptForChoice($Title, $Question, $options, 0)

    switch ($result) {
        0 {
            Clear-Host
            Program_Version $program_title
            $AD_Computer_Name = Read-Host -Prompt "Computer name: "
            Clear-Host
            Program_Version $program_title
            if($AD_Computer_Name)
            {
                $AD_Computer = Get-ADComputer -Identity $AD_Computer_Name -SearchBase $Computers_OU
                if($AD_Computer)
                {
                    write-host $AD_Computer
                } elseif (!$AD_Computer) 
                {
                    Clear-Host
                    Write-Host "No computer was found with the name: $AD_Computer_Name"
                    Write-Host "`nPress enter to continue..."; $Host.UI.ReadLine(); Clear-Host
                }
            }
        }
        1 { Password_Changer }
        2 { return Select_New_User }
    }
}


# Program info
$program_version = "1.0"
$program_title = "Eternal life $program_version `n"
$Dev_Mode = $False

# Active Directory Domain and Organizational Unit data
$BFK_DC = "DC=bfkskole,DC=top,DC=no"
$RYVS_Auto_Elever_OU = "OU=Auto Elever,OU=Brukere,OU=RYVS,OU=Virksomheter,$BFK_DC"
$Computers_OU = "OU=Maskiner Elev,$BFK_DC"


Clear-Host # Cleaner terminal when run
while($True)
{
    Program_Version $program_title
    # Active Directory user information
    $username = Read-Host -Prompt "Username" # Active Directory username
    $AD_User = Get-ADUser -SearchBase $RYVS_Auto_Elever_OU -Filter {SamAccountName -eq $username} -Properties * # AD user selected by $username

    while($True)
    {
        if(!$AD_User -And $Dev_Mode -eq $False)
        {
            Message_Box -Message "No user with the name '$username' was found." -Header "Oh dear" -Buttons "Ok" -Type "Warning"
            Clear-Host; break;
        }
        Clear-Host
        Get_User_Information

        $Main_Menu = New_Menu -Title $program_title -Question 'Please select option'


        if($Main_Menu -eq $False)
        {
            break
        }
    }
}
