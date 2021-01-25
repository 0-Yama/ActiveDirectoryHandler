# version 0.4.7
# Auteur : OYama_

# Force le type d'execution
Set-ExecutionPolicy Unrestricted

# Importation des modules
Import-Module ActiveDirectory

# Variable Globale

$Language       = ''
$Selection      = 0
$SelectionColor = 'Yellow'
$Header         = "","
 `t`t`t'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''",
"`t`t`t''                                                                           ''",
"`t`t`t''                               Initialization                              ''",
"`t`t`t''                                                                           ''",
"`t`t`t'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''",
"`t`t`tCreated by OYama_                                                 Version 0.4.6`n`n"

# Fonction de Menu
function Menu-ShowList ($List)
{
    clear
    for($i=0; $i -le $Header.Length; $i++)
    {
        Write-Host -Object $Header[$i]
    }
    for($i=0; $i -le $List.Length; $i++)
    {
        if($Selection -eq $i)
        {
            Write-Host -Object $List[$i] -ForegroundColor $SelectionColor
        }else
        {
            Write-Host -Object $List[$i]
        }
        Write-Host -Object ""
    }

}

function Menu-Main
{
    $MenuList = Get-Content -Path .\language\$Language\Main_Menu.txt
    Menu-ShowList($MenuList)
    do
    {
        $keyPress = [Console]::ReadKey($true).Key
    } until($keyPress -eq [ConsoleKey]::DownArrow -or $keyPress -eq [ConsoleKey]::UpArrow-or $keyPress -eq [ConsoleKey]::Enter)

    if($keyPress -eq [ConsoleKey]::DownArrow)
    {
        $Selection = ($Selection + 1) % $MenuList.Length
    }elseif($keyPress -eq [ConsoleKey]::UpArrow)
    {
        $Selection = ($Selection + $MenuList.Length - 1) % $MenuList.Length
    }
    elseif($keyPress -eq [ConsoleKey]::Enter)
    {
        Switch($Selection)
        {
            0{$Selection=0; Menu-User}
            1{$Selection=0; Menu-Group}
            2{$Selection=0; Menu-OU}
            3{Close-App}
        }
    }
    Menu-Main
}

function Menu-User
{
   $MenuList = Get-Content -Path .\language\$Language\User_Menu.txt
    Menu-ShowList($MenuList)
    do
    {
        $keyPress = [Console]::ReadKey($true).Key
    } until($keyPress -eq [ConsoleKey]::DownArrow -or $keyPress -eq [ConsoleKey]::UpArrow-or $keyPress -eq [ConsoleKey]::Enter)

    if($keyPress -eq [ConsoleKey]::DownArrow)
    {
        $Selection = ($Selection + 1) % $MenuList.Length
    }elseif($keyPress -eq [ConsoleKey]::UpArrow)
    {
        $Selection = ($Selection + $MenuList.Length - 1) % $MenuList.Length
    }
    elseif($keyPress -eq [ConsoleKey]::Enter)
    {
        Switch($Selection)
        {
            0{$Selection=0; Add-User}
            1{$Selection=0; Rem-User}
            2{$Selection=0; Add-UserToGroup}
            3{$Selection=0; Rem-UserFromGroup}
            4{$Selection=0; Mov-User-OU}
            5{$Selection=0; Menu-Main}
        }
    }
    Menu-User
}

function Menu-Group
{
   $MenuList = Get-Content -Path .\language\$Language\Group_Menu.txt
    Menu-ShowList($MenuList)
    do
    {
        $keyPress = [Console]::ReadKey($true).Key
    } until($keyPress -eq [ConsoleKey]::DownArrow -or $keyPress -eq [ConsoleKey]::UpArrow-or $keyPress -eq [ConsoleKey]::Enter)

    if($keyPress -eq [ConsoleKey]::DownArrow)
    {
        $Selection = ($Selection + 1) % $MenuList.Length
    }elseif($keyPress -eq [ConsoleKey]::UpArrow)
    {
        $Selection = ($Selection + $MenuList.Length - 1) % $MenuList.Length
    }
    elseif($keyPress -eq [ConsoleKey]::Enter)
    {
        Switch($Selection)
        {
            0{$Selection=0; Add-Group}
            1{$Selection=0; Rem-Group}
            2{$Selection=0; Add-GrouptoGroup}
            3{$Selection=0; Rem-GroupfromGroup}
            4{$Selection=0; Mov-Group-OU}
            5{$Selection=0; Menu-Main}
        }
    }
    Menu-Group
}

function Menu-OU
{

}


# Fonction général
function Startup
{
    
    $Language = Sel-Language
    $Header = Get-Content -Path .\language\$Language\Header.txt
    Menu-Main
}

function Close-App
{
    clear
    Get-Content -Path .\language\$Language\Exit.txt
    exit
}

function Sel-Language
{
    $MenuList = Get-Content -Path .\language\list.txt
    Menu-ShowList($MenuList)

    do
    {
        $keyPress = [Console]::ReadKey($true).Key
    } until($keyPress -eq [ConsoleKey]::DownArrow -or $keyPress -eq [ConsoleKey]::UpArrow-or $keyPress -eq [ConsoleKey]::Enter)

    if($keyPress -eq [ConsoleKey]::DownArrow)
    {
        $Selection = ($Selection + 1) % $MenuList.Length
        Sel-Language
    }elseif($keyPress -eq [ConsoleKey]::UpArrow)
    {
        $Selection = ($Selection + $MenuList.Length - 1) % $MenuList.Length
        Sel-Language
    }
    elseif($keyPress -eq [ConsoleKey]::Enter)
    {
        Switch($Selection)
        {
            0{$Selection=0; return 'French'}
            1{$Selection=0; return 'English'}
        }
    }
    
}
# Fonction gestion utilisateurs

# Add-
function Add-User
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\User_Q.txt

    $FisrtName = Read-Host -Prompt $Question[0]
    $LastName = Read-Host -Prompt $Question[1]

    $FullName = $FisrtName + " " + $LastName

    $Username = $FisrtName.Substring(0,1) + "_" + $LastName
    $Password = (convertto-securestring "aaaaaaA1" -AsPlainText -Force)

    $Date = Get-Date
    $Description = "créé par " + $env:USERNAME + " le " + $Date

    New-ADUser `
        -SamAccountName $Username `
        -Name $FullName `
        -GivenName $Firstname `
        -Surname $Lastname `
        -DisplayName $FullName `
        -AccountPassword $Password `
        -UserPrincipalName "$Username@oyama.local" `
        -Enabled $true `
        -ChangePasswordAtLogon $true
}

function Add-UserToGroup
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\User_Q.txt

    $Username = Read-Host -Prompt $Question[2]
    $Group = Read-Host -Prompt $Question[3]
    Add-ADGroupMember  -Identity $Group -Members $Username 
    pause
}


# Rem-
function Rem-User
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\User_Q.txt

    $Username = Read-Host -Prompt $Question[2]
    Remove-ADUser -Identity $Username
}

function Rem-UserFromGroup
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\User_Q.txt

    $Username = Read-Host -Prompt $Question[2]
    $Group = Read-Host -Prompt $Question[3]
    Remove-ADGroupMember  -Identity $Group -Members $Username 
}


# Mov-
function Mov-User-OU
{

}

# Fonction gestion des groupes

# Add-
function Add-Group
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\Group_Q.txt

    $MenuList = @($Question[0..2])

    Menu-ShowList($MenuList)

    do
    {
        $keyPress = [Console]::ReadKey($true).Key
    } until($keyPress -eq [ConsoleKey]::DownArrow -or $keyPress -eq [ConsoleKey]::UpArrow-or $keyPress -eq [ConsoleKey]::Enter)

    if($keyPress -eq [ConsoleKey]::DownArrow)
    {
        $Selection = ($Selection + 1) % $MenuList.Length
    }elseif($keyPress -eq [ConsoleKey]::UpArrow)
    {
        $Selection = ($Selection + $MenuList.Length - 1) % $MenuList.Length
    }
    elseif($keyPress -eq [ConsoleKey]::Enter)
    {
        clear
        $Header


        $GroupName = Read-Host -Prompt $Question[3]

        New-ADGroup -Name $GroupName -GroupScope $Selection
        break
    }

    Add-Group
}

function Add-GrouptoGroup
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\Group_Q.txt

    $GroupToAdd = Read-Host -Prompt $Question[4]
    $GroupDest = Read-Host -Prompt $Question[5]
    Add-ADGroupMember  -Identity $GroupDest -Members $GroupToAdd 
    pause
}

# Rem-
function Rem-Group
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\Group_Q.txt

    $GroupName = Read-Host -Prompt $Question[3] 
    Remove-ADGroup -Identity $GroupName
}

function Rem-GroupfromGroup
{
    clear
    $Header


    $Question = Get-Content -Path .\language\$Language\Group_Q.txt

    $GroupToAdd = Read-Host -Prompt $Question[4]
    $GroupDest = Read-Host -Prompt $Question[6]
    Remove-ADGroupMember  -Identity $GroupDest -Members $GroupToAdd 
}

# Mov-
function Mov-Group-OU
{

}


# Startup
Startup