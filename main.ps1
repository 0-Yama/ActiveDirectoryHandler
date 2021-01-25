# version 0.4.5
# Auteur : OYama_

# Force le type d'execution
Set-ExecutionPolicy Unrestricted

# Importation des modules
Import-Module ActiveDirectory

# Variable Globale

$Selection      = 0
$SelectionColor = 'Yellow'
$Header         = "
`t`t`t'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
`t`t`t''                                                                           ''
`t`t`t''                 Gestionnaires de l'Active Directory                       ''
`t`t`t''                                                                           ''
`t`t`t'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
`t`t`tCréé par OYama_                                      Version 0.4.5 - 01/22/2021`n`n

"
# Fonction de Menu
function Menu-ShowList ($List)
{
    clear
    Write-Host -Object $Header
    for($i=0; $i -le $List.Length; $i++)
    {
        if($Selection -eq $i)
        {
            Write-Host -Object $List[$i] -ForegroundColor $SelectionColor
        }else
        {
            Write-Host -Object $List[$i]
        }
    }

}

function Menu-Main
{
    $MenuList = @(
                    "`t - Gérer les utilisateurs`n",
                    "`t - Gérer les groupes`n",
                    "`t - Gérer les unités d'organisations`n"
                    "`t - Sortir"
                  )
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
   $MenuList = @(
                    "`t - Créer un utilisateur`n"
                    "`t - Suprimer un utilisateur`n"
                    "`t - Ajouter un utilisateur a un groupe`n"
                    "`t - Suprimer un utilisateur d'un groupe`n"
                    "`t - Changer un utilisateur d'unité d'organisation WIP`n"
                    "`t - Retour"
                  )
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
   $MenuList = @(
                    "`t - Créer un groupe`n"
                    "`t - Suprimer un groupe`n"
                    "`t - Ajouter un gourpe à un groupe`n"
                    "`t - Suprimer un groupe d'un groupe`n"
                    "`t - Changer un groupe d'unité d'organisation WIP`n"
                    "`t - Retour"
                  )
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
    Menu-Main
}

function Close-App
{
    clear
    "
    `t`t`t'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    `t`t`t''                                                                           ''
    `t`t`t''                               Au Revoir                                   ''
    `t`t`t''                                                                           ''
    `t`t`t'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    "
    exit
}

# Fonction gestion utilisateurs

# Add-
function Add-User
{
    clear
    
    $FisrtName = Read-Host -Prompt "Prénom de l'utilisateur"
    $LastName = Read-Host -Prompt "Nom de l'utilisateur"
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
    $Username = Read-Host -Prompt "Nom d'utilisateur"
    $Group = Read-Host -Prompt "Nom du groupe cible"
    Add-ADGroupMember  -Identity $Group -Members $Username 
    pause
}


# Rem-
function Rem-User
{
    clear
    $Username = Read-Host -Prompt "Nom d'utilisateur"
    Remove-ADUser -Identity $Username
}

function Rem-UserFromGroup
{
    clear
    $Username = Read-Host -Prompt "Nom d'utilisateur"
    $Group = Read-Host -Prompt "Nom du groupe cible"
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
    $MenuList = @(
                "`t 0 - Domaine Locale`n",
                "`t 1 - Global`n",
                "`t 2 - Universel`n")

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
        $GroupName = Read-Host -Prompt "Nom du groupe"

        New-ADGroup -Name $GroupName -GroupScope $Selection
        break
    }

    Add-Group
}

function Add-GrouptoGroup
{
    clear
    $GroupToAdd = Read-Host -Prompt "Nom du groupe à ajouter"
    $GroupDest = Read-Host -Prompt "Nom du groupe cible"
    Add-ADGroupMember  -Identity $GroupDest -Members $GroupToAdd 
    pause
}

# Rem-
function Rem-Group
{
    clear
    $GroupName = Read-Host -Prompt "Nom du groupe" 
    Remove-ADGroup -Identity $GroupName
}

function Rem-GroupfromGroup
{
    clear
    $GroupToAdd = Read-Host -Prompt "Nom d'utilisateur"
    $GroupDest = Read-Host -Prompt "Nom du groupe cible"
    Remove-ADGroupMember  -Identity $GroupDest -Members $GroupToAdd 
}

# Mov-
function Mov-Group-OU
{

}



# Startup
Startup