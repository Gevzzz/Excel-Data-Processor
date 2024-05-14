##Author: Gevorg Minasyan
##Version 0.0.1
##Date: 2024-05-09
##Owned by: NCIA CSU
##Description: This script processes the data from a text file and writes it to an Excel file.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

class ScanData {
    [string]$SerialNumberOrTag
    [string]$OfficeLocation
    [string]$Division

    ScanData([string]$serialNumberOrTag, [string]$officeLocation, [string]$division) {
        $this.SerialNumberOrTag = $serialNumberOrTag
        $this.OfficeLocation = $officeLocation
        $this.Division = $division
    }

    [bool] Validate() {
        $isValidOfficeLocation = $this.OfficeLocation -match "^[SLIP]"
        $isValidDivision = $this.Division -match "^B\d+$"
        return $isValidOfficeLocation -and $isValidDivision
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Data Processor - NCIA CSU"
$form.Size = New-Object System.Drawing.Size(600, 430)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Text = "Please select a text file to process"
$instructionLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$instructionLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
$instructionLabel.AutoSize = $true
$instructionLabel.Location = New-Object System.Drawing.Point(30, 20)

$fileSelectionButton = New-Object System.Windows.Forms.Button
$fileSelectionButton.Text = "Select Text File"
$fileSelectionButton.Font = New-Object System.Drawing.Font("Arial", 10)
$fileSelectionButton.Location = New-Object System.Drawing.Point(30, 60)
$fileSelectionButton.Size = New-Object System.Drawing.Size(150, 40)
$fileSelectionButton.BackColor = [System.Drawing.Color]::SteelBlue
$fileSelectionButton.ForeColor = [System.Drawing.Color]::White
$fileSelectionButton.FlatStyle = 'Flat'
$fileSelectionButton.FlatAppearance.BorderSize = 0
$fileSelectionButton.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select the Text File"
        $fileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        $result = $fileDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:selectedFile = $fileDialog.FileName
            $textBox.AppendText([Environment]::NewLine + [Environment]::NewLine + "Selected file: $global:selectedFile" + [Environment]::NewLine)
        }
    })

$processButton = New-Object System.Windows.Forms.Button
$processButton.Text = "Process Data"
$processButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$processButton.Location = New-Object System.Drawing.Point(210, 60)
$processButton.Size = New-Object System.Drawing.Size(150, 40)
$processButton.BackColor = [System.Drawing.Color]::SteelBlue
$processButton.ForeColor = [System.Drawing.Color]::White
$processButton.FlatStyle = 'Flat'
$processButton.FlatAppearance.BorderSize = 0
$processButton.Add_Click({
    try {
        $scannerData = Get-Content -Path $global:selectedFile | Where-Object { $_ -ne "" }  # Skip empty lines
        $rowsForExcel = @()
        $currentNCIATags = @()
        $currentOfficeNumber = $null
        $currentDivision = $null

        foreach ($line in $scannerData) {
            $parts = $line.Split(',')
            if ($parts.Length -ne 3) {
                Write-Host "Skipping invalid line: $line" -ForegroundColor Yellow
                $textBox.AppendText([Environment]::NewLine + "Skipping invalid line: $line")
                continue
            }

            $identifier = $parts[2]

            if ($identifier -like "NCIA*") {
                $currentNCIATags += $identifier
            } elseif ($identifier -match "^[SLIP]") {
                $currentOfficeNumber = $identifier
            } elseif ($identifier -match "^B\d+$") {
                $currentDivision = $identifier
                if ($currentNCIATags.Count -gt 0) {
                    foreach ($tag in $currentNCIATags) {
                        $scanData = [ScanData]::new($tag, $currentOfficeNumber, $currentDivision)
                        if ($scanData.Validate()) {
                            $rowsForExcel += , @($scanData.SerialNumberOrTag, $scanData.OfficeLocation, $scanData.Division)
                        } else {
                            Write-Host "Invalid data detected: $($scanData | ConvertTo-Json)" -ForegroundColor Red
                            $textBox.AppendText([Environment]::NewLine + "Invalid data detected: $($scanData | ConvertTo-Json)")
                        }
                    }
                } else {
                    Write-Host "Invalid data detected: No NCIA tags for office number $currentOfficeNumber and division $currentDivision" -ForegroundColor Red
                    $textBox.AppendText([Environment]::NewLine + "Invalid data detected: No NCIA tags for office number $currentOfficeNumber and division $currentDivision")
                }
                $currentNCIATags = @()
                $currentOfficeNumber = $null
                $currentDivision = $null
            } else {
                Write-Host "Invalid identifier detected: $identifier" -ForegroundColor Yellow
                $textBox.AppendText([Environment]::NewLine + "Invalid identifier detected: $identifier")
            }
        }

        $rowsForExcel | ForEach-Object {
            Write-Host "Data to be written to Excel: $_" -ForegroundColor Cyan
            $textBox.AppendText([Environment]::NewLine + "Data to be written to Excel: $_")
        }

        $excelFilePath = "C:\Users\gevor\Desktop\EXCEL CSU\data.xlsx"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($excelFilePath)
        $worksheet = $workbook.Worksheets.Item("Sheet1")

        # Find the first empty row in the Excel sheet in column C, if you want A, change 3 to 1
        $rowIndex = 1
        while ($worksheet.Cells.Item($rowIndex, 3).Value2 -ne $null) {
            $rowIndex += 1
        }

        # Write the data to the Excel sheet starting from the first empty row in column C , if you want A, change 3 to 1
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

        Write-Host "Data processed successfully!" -ForegroundColor Green
        $textBox.AppendText([Environment]::NewLine + [Environment]::NewLine + "Data processed successfully!")
    }
    catch {
        Write-Host "An error occurred: " + $_.Exception.Message -ForegroundColor Red
        $textBox.AppendText([Environment]::NewLine + [Environment]::NewLine + "An error occurred: " + $_.Exception.Message)
    }
})

$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Processing Controls"
$groupBox.Font = New-Object System.Drawing.Font("Arial", 10)
$groupBox.Location = New-Object System.Drawing.Point(30, 120)
$groupBox.Size = New-Object System.Drawing.Size(540, 100)

$clearContentCheckbox = New-Object System.Windows.Forms.CheckBox
$clearContentCheckbox.Text = "Clear content of source file after processing"
$clearContentCheckbox.Font = New-Object System.Drawing.Font("Arial", 9)
$clearContentCheckbox.AutoSize = $true
$clearContentCheckbox.Location = New-Object System.Drawing.Point(20, 20)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = 'Vertical'
$textBox.Location = New-Object System.Drawing.Point(30, 240)
$textBox.Size = New-Object System.Drawing.Size(540, 120)

$warningLabel = New-Object System.Windows.Forms.Label
$warningLabel.Text = "Please ensure that the Excel file is closed."
$warningLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$warningLabel.ForeColor = [System.Drawing.Color]::Red
$warningLabel.AutoSize = $true
$warningLabel.Location = New-Object System.Drawing.Point(30, 370)

$form.Controls.Add($instructionLabel)
$form.Controls.Add($fileSelectionButton)
$form.Controls.Add($processButton)
$form.Controls.Add($groupBox)
$groupBox.Controls.Add($clearContentCheckbox)
$form.Controls.Add($textBox)
$form.Controls.Add($warningLabel)

$form.ShowDialog()
