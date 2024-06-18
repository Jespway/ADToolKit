# Импорт библиотек для создания форм
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Создание формы
$form = New-Object System.Windows.Forms.Form
$form.Text = "Active Directory Scripts Interface"
$form.Size = New-Object System.Drawing.Size(400, 450)
$form.StartPosition = "CenterScreen"

# Функция сбора информации о ПК и пользователях из AD с выбором места сохранения
function Collect-ADInfo {
    # Создание диалога сохранения файла для компьютеров
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt"
    $saveFileDialog.Title = "Save Computers Information"
    $saveFileDialog.ShowDialog()
    $computersFilePath = $saveFileDialog.FileName

    if ($computersFilePath) {
        $computers = Get-ADComputer -Filter * -Properties * | Sort LastLogonDate | FT Name, ipv4*, OperatingSystem, LastLogonDate -Autosize | Out-String
        $computers | Out-File $computersFilePath -Append
        [System.Windows.Forms.MessageBox]::Show("Computers information saved to $computersFilePath")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Saving computers information was canceled.")
    }

    # Создание диалога сохранения файла для пользователей
    $saveFileDialog.Title = "Save Users Information"
    $saveFileDialog.ShowDialog()
    $usersFilePath = $saveFileDialog.FileName

    if ($usersFilePath) {
        $users = Get-ADUser -Filter * -Properties * | Sort LastLogonDate | FT SamAccountName, LastLogonDate | Out-String
        $users | Out-File $usersFilePath -Append
        [System.Windows.Forms.MessageBox]::Show("Users information saved to $usersFilePath")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Saving users information was canceled.")
    }
}

# Функция пинга IP-адресов и записи в файл с пользовательской маской
function Ping-IPAddresses {
    param (
        [string]$ipMask
    )

    function Set-IPAddress {
        param (
            [string]$baseIP,
            [int]$startRange,
            [int]$endRange
        )

        for ($c = $startRange; $c -le $endRange; $c++) {
            $ipadd = "$baseIP$c"
            if (Test-Connection $ipadd -Count 2 -Quiet) {
                try {
                    $hostEntry = [System.Net.Dns]::GetHostByAddress($ipadd)
                    $dnsname = $hostEntry.HostName
                } catch {
                    $dnsname = "Unknown"
                }
                $output = "$ipadd $dnsname"
                $output
            }
        }
    }

    # Разбор пользовательской маски
    if ($ipMask -match '^(\d{1,3}\.){3}(\d{1,3})-(\d{1,3})$') {
        $baseIP = $matches[1..3] -join ''
        $range = $matches[4].Split('-')
        $startRange = [int]$range[0]
        $endRange = [int]$range[1]

        # Диалог выбора места сохранения
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Text files (*.txt)|*.txt"
        $saveFileDialog.Title = "Save Ping Results"
        $saveFileDialog.ShowDialog()
        $filePath = $saveFileDialog.FileName

        if ($filePath) {
            Set-IPAddress -baseIP $baseIP -startRange $startRange -endRange $endRange | Out-File $filePath -Append
            [System.Windows.Forms.MessageBox]::Show("Ping results saved to $filePath")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Saving ping results was canceled.")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Invalid IP mask format. Please use format '192.168.1.1-254'")
    }
}

# Функция отключения старых компьютерных учетных записей с выбором периода
function Disable-OldComputers {
    param (
        [int]$monthsAgo
    )

    $Date = (Get-Date).AddMonths(-$monthsAgo)

    #List Computer Account
    $oldComputers = Get-ADComputer -Filter {LastLogonDate -lt $Date} -Properties LastLogonDate | Select-Object Name, LastLogonDate | Out-String
    $oldComputers | Out-File D:\old_computers.txt -Append

    #Disable Computer Account
    Get-ADComputer -Filter {LastLogonDate -lt $Date} -Properties LastLogonDate | Disable-ADAccount

    [System.Windows.Forms.MessageBox]::Show("Old computer accounts disabled and saved to D:\old_computers.txt")
}

# Создание элементов интерфейса
$lblIPMask = New-Object System.Windows.Forms.Label
$lblIPMask.Location = New-Object System.Drawing.Point(10, 10)
$lblIPMask.Size = New-Object System.Drawing.Size(360, 20)
$lblIPMask.Text = "Enter IP mask (e.g., 192.168.1.1-254):"

$txtIPMask = New-Object System.Windows.Forms.TextBox
$txtIPMask.Location = New-Object System.Drawing.Point(10, 35)
$txtIPMask.Size = New-Object System.Drawing.Size(360, 20)

$btnCollectADInfo = New-Object System.Windows.Forms.Button
$btnCollectADInfo.Location = New-Object System.Drawing.Point(10, 60)
$btnCollectADInfo.Size = New-Object System.Drawing.Size(360, 40)
$btnCollectADInfo.Text = "Collect AD Information"
$btnCollectADInfo.Add_Click({ Collect-ADInfo })

$btnPingIPAddresses = New-Object System.Windows.Forms.Button
$btnPingIPAddresses.Location = New-Object System.Drawing.Point(10, 110)
$btnPingIPAddresses.Size = New-Object System.Drawing.Size(360, 40)
$btnPingIPAddresses.Text = "Ping IP Addresses"
$btnPingIPAddresses.Add_Click({ 
    $ipMask = $txtIPMask.Text
    Ping-IPAddresses -ipMask $ipMask
})

$lblMonthsAgo = New-Object System.Windows.Forms.Label
$lblMonthsAgo.Location = New-Object System.Drawing.Point(10, 160)
$lblMonthsAgo.Size = New-Object System.Drawing.Size(360, 20)
$lblMonthsAgo.Text = "Select how many months ago to consider as old:"

$cmbMonthsAgo = New-Object System.Windows.Forms.ComboBox
$cmbMonthsAgo.Location = New-Object System.Drawing.Point(10, 185)
$cmbMonthsAgo.Size = New-Object System.Drawing.Size(360, 20)

# Заполнение выпадающего списка значениями от 1 до 24 месяцев
1..24 | ForEach-Object { $cmbMonthsAgo.Items.Add("$_ months") }
$cmbMonthsAgo.SelectedIndex = 5  # Default to 6 months

$btnDisableOldComputers = New-Object System.Windows.Forms.Button
$btnDisableOldComputers.Location = New-Object System.Drawing.Point(10, 210)
$btnDisableOldComputers.Size = New-Object System.Drawing.Size(360, 40)
$btnDisableOldComputers.Text = "Disable Old Computers"
$btnDisableOldComputers.Add_Click({ 
    $selectedValue = $cmbMonthsAgo.SelectedItem
    $monthsAgo = [int]$selectedValue.Split(' ')[0]
    Disable-OldComputers -monthsAgo $monthsAgo
})

# Добавление элементов на форму
$form.Controls.Add($lblIPMask)
$form.Controls.Add($txtIPMask)
$form.Controls.Add($btnCollectADInfo)
$form.Controls.Add($btnPingIPAddresses)
$form.Controls.Add($lblMonthsAgo)
$form.Controls.Add($cmbMonthsAgo)
$form.Controls.Add($btnDisableOldComputers)

# Отображение формы
$form.ShowDialog()
