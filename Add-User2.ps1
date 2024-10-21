Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Определение класса User
class User {
    [string]$GivenName
    [string]$MiddleName
    [string]$Surname
    [string]$DisplayName
    [string]$UserPrincipalName
    [string]$SAMAccountName

    User([string]$givenName, [string]$middleName, [string]$surname, [string]$samAccountName) {
        $this.GivenName = $givenName
        $this.MiddleName = $middleName
        $this.Surname = $surname
        $this.SAMAccountName = $samAccountName
        $this.DisplayName = if ($middleName) { "$surname $givenName $middleName" } else { "$surname $givenName" }
        $this.UserPrincipalName = "$samAccountName@yourdomain.com" # Замените на ваш домен
    }

    [void]CreateUser([string]$ou) {
        $userParams = @{
            Name = $this.DisplayName
            GivenName = $this.GivenName
            Surname = $this.Surname
            DisplayName = $this.DisplayName
            UserPrincipalName = $this.UserPrincipalName
            SAMAccountName = $this.SAMAccountName
            Path = $ou
            AccountPassword = (ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force)
            Enabled = $true
        }

        if ($this.MiddleName) {
            $userParams.OtherAttributes = @{ middleName = $this.MiddleName }
        }

        New-ADUser @userParams

        # Установка флага "требовать смены пароля при следующем входе в систему"
        Set-ADUser -Identity $this.SAMAccountName -ChangePasswordAtLogon $true
    }

    [bool]UserExists() {
        return [bool](Get-ADUser -Filter { SamAccountName -eq $this.SAMAccountName } -ErrorAction SilentlyContinue)
    }
}

# Функция для получения списка OU
function Get-OUList {
    $ouList = Get-ADOrganizationalUnit -Filter * | Select-Object -Property DistinguishedName

    # Фильтрация списка OU, чтобы включить только те, в пути которых есть "OU=Users"
    $filteredOUList = $ouList | Where-Object { $_.DistinguishedName -match "OU=Users" }
    return $filteredOUList
}

# Создание формы
$form = New-Object System.Windows.Forms.Form
$form.Text = "Создание пользователя AD"
$form.Size = New-Object System.Drawing.Size(400, 400)
$form.StartPosition = "CenterScreen"

# Создание меток и текстовых полей для ввода данных пользователя
$labels = @("Фамилия", "Имя", "Отчество", "Имя входа пользователя")
$yPos = 20
$textBoxes = @() # Инициализация как массив

foreach ($labelText in $labels) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(10, $yPos)
    $label.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(170, $yPos)
    $textBox.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($textBox)
    $textBoxes += $textBox

    $yPos += 30
}

# Создание ComboBox для выбора OU
$ouLabel = New-Object System.Windows.Forms.Label
$ouLabel.Text = "Выберите OU"
$ouLabel.Location = New-Object System.Drawing.Point(10, $yPos)
$ouLabel.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($ouLabel)

$ouComboBox = New-Object System.Windows.Forms.ComboBox
$ouComboBox.Location = New-Object System.Drawing.Point(170, $yPos)
$ouComboBox.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($ouComboBox)

# Заполнение ComboBox списком OU
$ouList = Get-OUList
foreach ($ou in $ouList) {
    $null = $ouComboBox.Items.Add($ou.DistinguishedName)
}

$yPos += 40

# Создание кнопки для создания пользователя
$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = "Создать пользователя"
$createButton.Location = New-Object System.Drawing.Point(150, $yPos)
$createButton.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($createButton)

# Добавление обработчика события для нажатия кнопки
$createButton.Add_Click({
    $surname = $textBoxes[0].Text
    $givenName = $textBoxes[1].Text
    $middleName = $textBoxes[2].Text
    $samAccountName = $textBoxes[3].Text
    $selectedOU = $ouComboBox.SelectedItem

    # Отладочные сообщения
    Write-Host "Фамилия: $surname"
    Write-Host "Имя: $givenName"
    Write-Host "Отчество: $middleName"
    Write-Host "Имя входа пользователя: $samAccountName"
    Write-Host "Выбранный OU: $selectedOU"

    if ($surname -and $givenName -and $samAccountName -and $selectedOU) {
        $user = [User]::new($givenName, $middleName, $surname, $samAccountName)
        if ($user.UserExists()) {
            [System.Windows.Forms.MessageBox]::Show("Учетная запись с именем $samAccountName уже существует.")
        } else {
            try {
                $user.CreateUser($selectedOU)
                [System.Windows.Forms.MessageBox]::Show("Пользователь $($user.DisplayName) успешно создан в OU $selectedOU")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Ошибка при создании пользователя: $_")
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Пожалуйста, заполните все обязательные поля и выберите OU.")
    }
})

# Показ формы
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
