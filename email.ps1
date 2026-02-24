# ==============================
# CONFIGURATION
# ==============================

# Folder to search (local OneDrive folder)
$oneDriveFolder = "$env:C:\Users\drew\OneDrive\Attachments\excel"

# File pattern for Excel files
$filePattern = "*.xlsx"

# Email details
$toEmail = ""
$subject = "Request for Data Extract - EAM Bin Vehicle Stock Statement Report"
$body = "Hi Elyza,`n`nSee attached requested extract.`n`nRegards."

# ==============================
# FIND THE LATEST EXCEL FILE
# ==============================

$latestFile = Get-ChildItem -Path $oneDriveFolder -Filter $filePattern |
    Where-Object { -not $_.PSIsContainer } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $latestFile) {
    Write-Host "No Excel files found in $oneDriveFolder"
    exit
}

Write-Host "Latest Excel file found: $($latestFile.FullName)"

# ==============================
# CREATE OUTLOOK EMAIL
# ==============================

# Create Outlook COM object
$outlook = New-Object -ComObject Outlook.Application

# Create a new mail item
$mail = $outlook.CreateItem(0)  # 0 = olMailItem

# Set email properties
$mail.To = $toEmail
$mail.Subject = $subject
$mail.Body = $body
$mail.Attachments.Add($latestFile.FullName)

# Send the email
$mail.Send()


Write-Host "Email sent successfully via Outlook!"
