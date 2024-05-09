##Author: Gevorg Minasyan
##Version 0.0.1
##Date: 2024-05-09
##Owned by: NCIA CSU
##Description: This script processes the data from a text file and writes it to an Excel file.
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Data Processor NCIA CSU"
$form.Size = New-Object System.Drawing.Size(500, 380) 
$form.StartPosition = 'CenterScreen' 
$form.BackColor = [System.Drawing.Color]::LightGray 

$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Text = "PLEASE CLOSE THE EXCEL FILE BEFORE PROCESSING THE DATA"
$instructionLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold) 
$instructionLabel.ForeColor = [System.Drawing.Color]::DarkRed 
$instructionLabel.Location = New-Object System.Drawing.Point(20, 20)
$instructionLabel.AutoSize = $true

$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Processing Controls"
$groupBox.Font = New-Object System.Drawing.Font("Arial", 10)
$groupBox.Size = New-Object System.Drawing.Size(460, 180)
$groupBox.Location = New-Object System.Drawing.Point(20, 50)

$clearContentCheckbox = New-Object System.Windows.Forms.CheckBox
$clearContentCheckbox.Text = "Clear content of source file after processing"
$clearContentCheckbox.Font = New-Object System.Drawing.Font("Arial", 9)
$clearContentCheckbox.AutoSize = $true
$clearContentCheckbox.Location = New-Object System.Drawing.Point(20, 40) 

$button = New-Object System.Windows.Forms.Button
$button.Text = "Process Data"
$button.Size = New-Object System.Drawing.Size(120, 40)
$button.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$button.Location = New-Object System.Drawing.Point(180, 100) 

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = 'Vertical'
$textBox.Location = New-Object System.Drawing.Point(20, 240)
$textBox.Size = New-Object System.Drawing.Size(460, 90) 

$button.Add_Click({
    try {
        $scannerData = Get-Content -Path 'C:\Users\gevor\Desktop\DM.txt'
        $rowsForExcel = @()
        $currentNCIATags = @()
        $currentOfficeNumber = $null
        $currentDivision = $null

        foreach ($line in $scannerData) {
            $parts = $line.Split(',')
            $identifier = $parts[2]

            if ($identifier -like "NCIA*") {
                $currentNCIATags += $identifier
            } elseif ($identifier -notlike "NCIA*" -and $currentOfficeNumber -eq $null) {
                $currentOfficeNumber = $identifier
            } elseif ($identifier -notlike "NCIA*" -and $currentOfficeNumber -ne $null) {
                $currentDivision = $identifier
                foreach ($tag in $currentNCIATags) {
                    $currentScanData = @($tag, $currentOfficeNumber, $currentDivision)
                    $rowsForExcel += ,$currentScanData
                }
                $currentNCIATags = @()
                $currentOfficeNumber = $null
                $currentDivision = $null
            }
        }

        $excelFilePath = "C:\Users\gevor\Desktop\data.xlsx"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($excelFilePath)
        $worksheet = $workbook.Worksheets.Item("Sheet1")

        $rowIndex = 1
        while ($worksheet.Cells.Item($rowIndex, 3).Value2 -ne $null) {
            $rowIndex += 1
        }

        foreach ($row in $rowsForExcel) {
            $colIndex = 3
            foreach ($item in $row) {
                $worksheet.Cells.Item($rowIndex, $colIndex).Value2 = $item
                $colIndex += 1
            }
            $rowIndex += 1
        }

        $workbook.Save()
        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

        if ($clearContentCheckbox.Checked) {
            Clear-Content -Path 'C:\Users\gevor\Desktop\DM.txt'
        }

        $textBox.Text = "Data processed successfully!"
    } catch {
        $textBox.Text = "An error occurred: " + $_.Exception.Message
    }
})

$form.Controls.Add($instructionLabel)
$form.Controls.Add($groupBox)
$groupBox.Controls.Add($button)
$groupBox.Controls.Add($clearContentCheckbox)
$form.Controls.Add($textBox)

$form.ShowDialog()
