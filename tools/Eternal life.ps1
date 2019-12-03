using namespace System.Management.Automation.Host

# Add support for message boxes to shells that does not have them supported.
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")


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
    [System.Windows.Forms.MessageBox]::Show($Message, $Header, $Buttons, $Type) # Creates message box
}


function Get_User_Information
{
    if(!$AD_User.Enabled)
    {
        Write-Host "Warning, this user is locked!" -ForegroundColor Red
    }
    
    
    Write-Host "Selected user:" $AD_User.SamAccountName
    
    
    if($AD_User.employeeID)
    {
        Write-Host "Employee ID: " $AD_User.employeeID
    }
}


# Password changer
function Password_Changer
{
    $pwd = Read-Host "Password: " -AsSecureString # Request users password
    $pwd_confirm = Read-Host "Re-enter Password: " -AsSecureString # Request user to re-type password
    $pwd_txt = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd))
    $pwd_txt_confirm = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd_confirm))

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
    return 0
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
    $Clear_Terminal = New_Menu_Item -Option "&Wipe terminal" -Tutorial "Option will clear the terminal"
    
    
    $options = [ChoiceDescription[]]($Get_Bitlocker_Info, $change_password, $Clear_Terminal, $Select_New_User)
    $result = $host.ui.PromptForChoice($Title, $Question, $options, 0)

    switch ($result) {
        0 
        { 
            if($BitLocker_Info)
            {
                $BitLocker_Info
            } else
            {
                Write-Host "You have not selected a computer, please go back and select a valid computer to get bitlocker information about it."
                Read-Host "Press enter to continue..."; Clear-Host
            }
        }
        1 { Password_Changer }
        2 { Clear-Host; Program_Version -Program_Title $program_title  }
        3 { return Select_New_User }
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
    Program_Version -Program_Title $program_title
    # Active Directory user information
    $username = Read-Host -Prompt "Username" # Active Directory username
    $AD_User = Get-ADUser -SearchBase $RYVS_Auto_Elever_OU -Filter {SamAccountName -eq $username} # AD user selected by $username
    
    # Active directory computer information
    $AD_Computer_Name = Read-Host -Prompt "[Optinal] Computer name: "
    if($AD_Computer_Name)
    {
        $AD_Computer = Get-ADComputer -Identity $AD_Computer_Name -SearchBase $Computers_OU
        if($AD_Computer)
        {
            $BitLocker_Info = $AD_Computer | Read-ADRecoveryInformation
        }
    }
    
    while($True)
    {
        if(!$AD_User -And $Dev_Mode -eq $False)
        {
            Message_Box -Message "No user with the name '$username' was found." -Header "Oh dear" -Buttons "Ok" -Type "Warning"
            Clear-Host; break;
        }

        Get_User_Information
        
        $Main_Menu = New_Menu -Title $program_title -Question 'Please select option'
        
        
        if($Main_Menu -eq 0)
        {
            break
        }
    }
}
