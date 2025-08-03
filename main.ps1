Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-InputBox {
    param (
        [string]$message,
        [string]$title
    )

    # Create and configure the input form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = "CenterScreen"

    # Add a label with the prompt message
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10,10)
    $form.Controls.Add($label)

    # Add a textbox for user input
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Width = 260
    $form.Controls.Add($textBox)

    # Add an OK button to validate input
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(190,70)
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Display the dialog and return the text entered, or null if cancelled
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $null
    }
}

function Show-PasswordBox {
    param (
        [string]$Message = "Entrez un mot de passe",
        [string]$Title = "Mot de passe"
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Create and configure the password entry form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = "CenterScreen"

    # Add a label with instructions
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10,10)
    $form.Controls.Add($label)

    # Add a password-protected textbox
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Width = 260
    $textBox.UseSystemPasswordChar = $true
    $form.Controls.Add($textBox)

    # Add an OK button to validate input
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(190,70)
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Display the dialog and return the entered password as a SecureString
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return (ConvertTo-SecureString $textBox.Text -AsPlainText -Force)
    } else {
        return $null
    }
}

function Show-SelectionBox {
    param (
        [string]$Title = "Sélection",
        [string]$Message = "Faites un choix",
        [string[]]$Options
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Add "Create new" as the first option in the list
    $allOptions = @("Créer nouveau...") + $Options

    # Create and configure the selection form
    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size(400,250)
    $form.StartPosition = "CenterScreen"

    # Add a label with instructions
    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = New-Object Drawing.Point(10,10)
    $form.Controls.Add($label)

    # Add a list box to display options
    $listBox = New-Object Windows.Forms.ListBox
    $listBox.Location = New-Object Drawing.Point(10,40)
    $listBox.Size = New-Object Drawing.Size(360,120)
    $listBox.SelectionMode = "One"
    $listBox.Items.AddRange($allOptions)
    $form.Controls.Add($listBox)

    # Add an OK button to confirm the choice
    $okButton = New-Object Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object Drawing.Point(290,170)
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    # Show the dialog and return the selected option, or prompt for a new one
    if ($form.ShowDialog() -eq "OK" -and $listBox.SelectedItem) {
        if ($listBox.SelectedItem -eq "Créer nouveau...") {
            # Prompt the user to enter a new name
            return Show-InputBox "Saisissez le nom à créer :" $Title
        } else {
            return $listBox.SelectedItem
        }
    } else {
        return $null
    }
}

function Show-FolderPicker {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Sélectionnez l'emplacement du dossier personnel"
    $folderBrowser.ShowNewFolderButton = $true

    # Show folder picker and return selected path or null if cancelled
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Aucun dossier sélectionné. Opération annulée.", "Annulation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $null
    }
}

function Show-YesNo($message, $title = "Confirmation") {
    # Show a confirmation dialog with Yes/No buttons
    return [System.Windows.Forms.MessageBox]::Show($message, $title, "YesNo", "Question")
}

function Remove-Spaces {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputString
    )
    # Remove all whitespace characters from the input string
    return $InputString -replace '\s', ''
}


# ------------------- Step 1: User information --------------------
# Ask for user's first name, remove spaces, and convert to lowercase
$firstName = Remove-Spaces -InputString (Show-InputBox "Prénom de l'utilisateur" "Création d'utilisateur").ToLower()

# Ask for user's last name, remove spaces, and convert to lowercase
$lastName = Remove-Spaces -InputString (Show-InputBox "Nom de l'utilisateur" "Création d'utilisateur").ToLower()

# Generate login name in the format "f.lastname"
$AccountName = "$(${firstName}.Substring(0,1)).$lastName"

# Generate full user principal name (email-like format)
$userPrincipalName = "${AccountName}@tssr-crew.local"

# Create personal folder name by appending '_Perso'
$AccountName_Perso = "${AccountName}_Perso"

# Ask user to select the folder location for personal directory
$dossierParent = Show-FolderPicker

# If user cancels folder selection, terminate script
if (-not $dossierParent) { return }

# Generate full path for the personal folder
$dossierPath = Join-Path $dossierParent "$AccountName_Perso"

# ------------------- Step 2: Organizational Unit (OU) --------------------
# Ask whether to place the user in a specific OU
$placeInOU = Show-YesNo "Placer $firstName $lastName dans une unité d'organisation ?"

if ($placeInOU -eq "Yes") {
    # Get list of existing OUs
    $ouList = Get-ADOrganizationalUnit -Filter * | Sort-Object Name | Select-Object -ExpandProperty Name 

    # Let the user select or create an OU
    $selectedOU = Show-SelectionBox -Title "Choix de l'OU" -Message "Sélectionnez ou créez une OU :" -Options $ouList

    if (-not $selectedOU) {
        # No OU selected, use domain root as default
        [System.Windows.Forms.MessageBox]::Show("Aucune OU sélectionnée. L'utilisateur sera placé à la racine du domaine.", "Info") | Out-Null
        $ouPath = "DC=TSSR-CREW,DC=LOCAL"
    }
    else {
        # Check if selected OU exists
        $ouExists = Get-ADOrganizationalUnit -Filter "Name -eq '$selectedOU'"

        if (-not $ouExists) {
            # Ask to create new OU if it doesn't exist
            $createOU = Show-YesNo "L'OU '$selectedOU' n'existe pas. La créer ?"
            if ($createOU -eq "Yes") {
                # Create the OU and build the path
                New-ADOrganizationalUnit -Name $selectedOU -Path "DC=TSSR-CREW,DC=LOCAL" | Out-Null
                $ouPath = "OU=$selectedOU,DC=TSSR-CREW,DC=LOCAL"
                [System.Windows.Forms.MessageBox]::Show("L'unité d'organisation $selectedOU a été créée avec succès", "Succès") | Out-Null
                [System.Windows.Forms.MessageBox]::Show("L'utilisateur $firstName $lastName a bien été ajouté à l'unité d'organisation $selectedOU", "Succès") | Out-Null
            } else {
                # Use domain root if creation was declined
                $ouPath = "DC=TSSR-CREW,DC=LOCAL"
            }
        } else {
            # Use existing OU path
            $ouPath = "OU=$selectedOU,DC=TSSR-CREW,DC=LOCAL"
            [System.Windows.Forms.MessageBox]::Show("L'utilisateur $firstName $lastName a bien été ajouté à l'unité d'organisation $selectedOU", "Succès") | Out-Null
        }
    }
} else {
    # If user chose not to use an OU, place in domain root
    $ouPath = "DC=TSSR-CREW,DC=LOCAL"
}

# ------------------- Step 3: User creation --------------------
# Check if the user already exists
$user = Get-ADUser -Filter "SamAccountName -eq '$AccountName'"
if ($user) {
    # Show error and exit if user exists
    [System.Windows.Forms.MessageBox]::Show("L'utilisateur existe déjà", "Erreur", "OK", "Error") | Out-Null
    exit
}

# Ask for a password using the secure password prompt
$SecurePassword = Show-PasswordBox -Message "Entrez le mot de passe de l'utilisateur" -Title "Mot de passe"

# If no password is entered, cancel the operation
if (-not $SecurePassword) {
    [System.Windows.Forms.MessageBox]::Show("Aucun mot de passe saisi. Opération annulée.", "Erreur", "OK", "Error") | Out-Null
    exit
}

# Create the new Active Directory user with the provided information
[void](New-ADUser -Name "$firstName $lastName" `
           -DisplayName "$lastName $firstName" `
           -GivenName "$firstName" `
           -Surname "$lastName" `
           -SamAccountName "$AccountName" `
           -UserPrincipalName $userPrincipalName `
           -EmailAddress $userPrincipalName `
           -Path $ouPath `
           -AccountPassword $SecurePassword `
           -ChangePasswordAtLogon $true `
           -Enabled $true)

# Confirm that the user was created
[System.Windows.Forms.MessageBox]::Show("L'utilisateur $firstName $lastName a été créé avec succès", "Succès") | Out-Null

# ------------------- Step 4: Add user to group --------------------
# Ask whether to add the user to a group
$addToGroup = Show-YesNo "Ajouter $firstName $lastName à un groupe ?"
if ($addToGroup -eq "Yes") {
    # Retrieve list of existing groups
    $groupList = Get-ADGroup -Filter * | Sort-Object Name | Select-Object -ExpandProperty Name

    # Let user select or create a group
    $selectedGroup = Show-SelectionBox -Title "Choix du groupe" -Message "Sélectionnez ou créez un groupe :" -Options $groupList

    if ($selectedGroup) {
        # Check if the selected group exists
        $group = Get-ADGroup -Filter "Name -eq '$selectedGroup'"

        if (-not $group) {
            # Ask to create the group if it doesn't exist
            $createGroup = Show-YesNo "Le groupe '$selectedGroup' n'existe pas. Le créer ?"
            if ($createGroup -eq "Yes") {
                # Create the group in the same OU as the user
                New-ADGroup -Name $selectedGroup -GroupScope Global -Path $ouPath | Out-Null
                [System.Windows.Forms.MessageBox]::Show("Le groupe $selectedGroup a été créé avec succès", "Succès") | Out-Null
                [System.Windows.Forms.MessageBox]::Show("L'utilisateur $firstName $lastName a bien été ajouté au groupe $selectedGroup", "Succès") | Out-Null
            } else {
                $selectedGroup = $null
            }
        }

        if ($selectedGroup) {
            # Add the user to the group
            Add-ADGroupMember -Identity $selectedGroup -Members $AccountName | Out-Null
            [System.Windows.Forms.MessageBox]::Show("L'utilisateur $firstName $lastName a bien été ajouté au groupe $selectedGroup", "Succès") | Out-Null
        }
    } else {
        # No group was selected
        [System.Windows.Forms.MessageBox]::Show("Aucun groupe sélectionné.", "Information") | Out-Null
    }
}

# ------------------- Step 5: Create folder and share --------------------
# Create personal folder if it doesn't exist
if (-not (Test-Path $dossierPath)) {
    New-Item -Path $dossierPath -ItemType Directory | Out-Null
    [System.Windows.Forms.MessageBox]::Show("le dossier $AccountName_Perso a bien été créé sous $dossierParent", "Succès") | Out-Null
}

# Create a shared folder with full access for the user if not already shared
if (-not (Get-SmbShare -Name $AccountName_Perso -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name $AccountName_Perso -Path $dossierPath -FullAccess $userPrincipalName | Out-Null
    [System.Windows.Forms.MessageBox]::Show("le partage du dossier $AccountName_Perso a bien été effectué", "Succès") | Out-Null
}

# ------------------- Step 6: Set NTFS permissions --------------------
# Remove inherited permissions on the folder
icacls $dossierPath /inheritance:r | Out-Null

# Define domain path for the user
$domainSam = "TSSR-CREW\$AccountName"

# Grant the user full NTFS permissions on the folder
icacls $dossierPath /grant "${domainSam}:(OI)(CI)F" | Out-Null
[System.Windows.Forms.MessageBox]::Show("Les droit exclusifs ont bien été accordés à $firstName $lastName pour le dossier $AccountName_Perso", "Succès") | Out-Null