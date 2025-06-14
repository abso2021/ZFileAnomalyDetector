Add-Type -AssemblyName System.Windows.Forms

# -------- EMAIL CONFIGURATION --------
$smtpServer = "smtp.gmail.com"
$smtpPort = 587
$emailFrom = "abbassobhi@gmail.com"
$emailDisplayName = "ZFile Anomaly Detector"
$emailPassword = "tzug seyi xucb smdl"

# ---------- GUI FUNCTIONS ----------

function Get-ValidEmail {
    param([string]$prompt)

    do {
        $inputBox = New-Object System.Windows.Forms.Form
        $inputBox.Text = $prompt
        $inputBox.Width = 400
        $inputBox.Height = 160
        $inputBox.StartPosition = "CenterScreen"

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Enter email address:"
        $label.Left = 10
        $label.Top = 20
        $label.Width = 360
        $inputBox.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Left = 10
        $textBox.Top = 45
        $textBox.Width = 360
        $textBox.TabIndex = 0
        $inputBox.Controls.Add($textBox)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Left = 100
        $okButton.Top = 85
        $okButton.Width = 80
        $okButton.TabIndex = 1
        $okButton.Add_Click({ $inputBox.DialogResult = [System.Windows.Forms.DialogResult]::OK; $inputBox.Close() })
        $inputBox.Controls.Add($okButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Left = 200
        $cancelButton.Top = 85
        $cancelButton.Width = 80
        $cancelButton.TabIndex = 2
        $cancelButton.Add_Click({ $inputBox.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $inputBox.Close() })
        $inputBox.Controls.Add($cancelButton)

        $inputBox.AcceptButton = $okButton
        $inputBox.CancelButton = $cancelButton
        $textBox.Select()
        $dialogResult = $inputBox.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
            Write-Host "Operation cancelled by user."
            exit
        }

        $email = $textBox.Text.Trim()

        $pattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
        if ($email -match $pattern) {
            return $email
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid email address. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } while ($true)
}

function Get-PositiveInt {
    param([string]$prompt)

    do {
        $inputBox = New-Object System.Windows.Forms.Form
        $inputBox.Text = $prompt
        $inputBox.Width = 400
        $inputBox.Height = 160
        $inputBox.StartPosition = "CenterScreen"

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Enter number of email addresses to receive alerts:"
        $label.Left = 10
        $label.Top = 20
        $label.Width = 360
        $inputBox.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Left = 10
        $textBox.Top = 45
        $textBox.Width = 360
        $textBox.TabIndex = 0
        $inputBox.Controls.Add($textBox)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Left = 100
        $okButton.Top = 85
        $okButton.Width = 80
        $okButton.TabIndex = 1
        $okButton.Add_Click({ $inputBox.DialogResult = [System.Windows.Forms.DialogResult]::OK; $inputBox.Close() })
        $inputBox.Controls.Add($okButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Left = 200
        $cancelButton.Top = 85
        $cancelButton.Width = 80
        $cancelButton.TabIndex = 2
        $cancelButton.Add_Click({ $inputBox.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $inputBox.Close() })
        $inputBox.Controls.Add($cancelButton)

        $inputBox.AcceptButton = $okButton
        $inputBox.CancelButton = $cancelButton
        $textBox.Select()
        $dialogResult = $inputBox.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
            Write-Host "Operation cancelled by user."
            exit
        }

        if ([int]::TryParse($textBox.Text.Trim(), [ref]$null) -and [int]$textBox.Text.Trim() -gt 0) {
            return [int]$textBox.Text.Trim()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid number. Please enter a positive integer.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } while ($true)
}

# ---------- START SCRIPT ----------

# 1) Get the number of recipient emails
$emailCount = Get-PositiveInt -prompt "How many email addresses should receive alerts?"

# 2) Get the email addresses
$recipients = @()
for ($i = 1; $i -le $emailCount; $i++) {
    $email = Get-ValidEmail -prompt "Recipient $i of $emailCount"
    $recipients += $email
}

# 3) Select the source folder
$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = "Select the source folder containing .z files"
if ($folderDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No folder selected. Exiting script."
    exit
}
$sourceFolder = $folderDialog.SelectedPath

# 4) Create the destination folder inside the source folder
$destinationFolder = Join-Path -Path $sourceFolder -ChildPath ("$([System.IO.Path]::GetFileName($sourceFolder))_converted")
if (-not (Test-Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder | Out-Null
}

Write-Host "Monitoring started. Press Ctrl+C to stop."

# ---------- MAIN LOOP ----------
while ($true) {
    $faultyFiles = @()

    # Get .z files in source folder
    $zFiles = Get-ChildItem -Path $sourceFolder -Filter *.z -ErrorAction SilentlyContinue
    foreach ($file in $zFiles) {
        try {
            $lines = Get-Content -Path $file.FullName
            if ($lines.Count -lt 143) { continue }

            # Process lines from line 141 to end (0-based index: line 140)
            $processed = $lines[140..($lines.Count - 1)]
            # Remove lines equal to second line of processed (index 1)
            $processed = $processed | Where-Object { $_ -ne $processed[1] }

            $processedCsvLines = foreach ($line in $processed) {
                ($line -replace '\s+', ',')
            }

            # Shift header cells one position to the left
            if ($processedCsvLines.Count -gt 0) {
                $headerParts = $processedCsvLines[0] -split ','
                if ($headerParts.Count -gt 1) {
                    $processedCsvLines[0] = ($headerParts[1..($headerParts.Count - 1)]) -join ','
                }
            }

            # Output CSV file path
            $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".csv"
            $outputPath = Join-Path -Path $destinationFolder -ChildPath $outputFileName

            # Delete existing CSV if not locked
            if (Test-Path $outputPath) {
                try {
                    Remove-Item -Path $outputPath -Force
                } catch {
                    Write-Host "Could not delete locked file: $outputPath. Skipping..."
                    continue
                }
            }

            # Save CSV
            $processedCsvLines | Set-Content -Path $outputPath

            # Check CSV for column 5 and 6 values
            $csvLines = Get-Content -Path $outputPath
            if ($csvLines.Count -lt 2) { continue }

            $dataRows = $csvLines[1..($csvLines.Count - 1)]
            $countZprime = 0
            $countZdoubleprime = 0

            foreach ($row in $dataRows) {
                $cols = $row -split ','
                if ($cols.Count -ge 6) {
                    $zPrimeVal = $cols[4]
                    $zDoublePrimeVal = $cols[5]

                    if ([double]::TryParse($zPrimeVal, [ref]$null) -and [double]$zPrimeVal -gt 5) {
                        $countZprime++
                    }

                    if ([double]::TryParse($zDoublePrimeVal, [ref]$null) -and [double]$zDoublePrimeVal -lt -3) {
                        $countZdoubleprime++
                    }
                }
            }

            if ($countZprime -gt 5 -or $countZdoubleprime -gt 5) {
                $faultyFiles += $outputPath
            }

        } catch {
            Write-Host "Error processing file $($file.FullName): $_"
            continue
        }
    }

    # If there are faulty files, send a summary email
    if ($faultyFiles.Count -gt 0) {
        $subject = "ZFile Anomaly Detect Alert - Anomaly Detection/Data Validation Issues Found"
        $body = "Dear User,`n`nThe following files have anomaly detection/data validation issues:`n`n"

        foreach ($f in $faultyFiles) {
            $body += "$f`n"
        }
        $body += "`nRegards,`nZFile Anomaly Detector"

        try {
            Send-MailMessage -From "$emailDisplayName <$emailFrom>" `
                             -To $recipients `
                             -Subject $subject `
                             -Body $body `
                             -SmtpServer $smtpServer `
                             -Port $smtpPort `
                             -UseSsl `
                             -Credential (New-Object System.Management.Automation.PSCredential -ArgumentList $emailFrom, (ConvertTo-SecureString $emailPassword -AsPlainText -Force))
            Write-Host "Alert email sent."
        } catch {
            Write-Host "Failed to send email: $_"
        }
    } else {
        Write-Host "No data validation issues found."
    }

    Start-Sleep -Seconds 3600  # Sleep for 1 hour
}
