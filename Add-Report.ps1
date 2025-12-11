# ==========================================
# Script: Update SharePoint Document Library Metadata from CSV
# ==========================================

# مرحله 1: نصب و Import ماژول PnP PowerShell (اگر نصب نیست)
# فقط یک بار نیاز است
# Install-Module -Name PnP.PowerShell -Force -AllowClobber

Import-Module PnP.PowerShell

# مرحله 2: تنظیمات اولیه
$SiteUrl = "https://extranet.pedc.ir/pogp/PRD"
$LibraryName = "Production Engineering Report"
$CSVPath = "D:\Sepher_Pasargad\works\Maintenace\PythonDataAnalysis\PythonPractice\Weekly.csv"  # مسیر فایل CSV خود را اینجا بگذارید

# مرحله 3: اتصال به SharePoint
Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive
    Write-Host "Connected successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Error connecting to SharePoint: $_" -ForegroundColor Red
    exit
}

# مرحله 4: بررسی وجود Library
Write-Host "Checking library existence..." -ForegroundColor Cyan
$library = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue
if ($null -eq $library) {
    Write-Host "Library '$LibraryName' not found!" -ForegroundColor Red
    exit
}

# مرحله 5: خواندن فایل CSV
Write-Host "Reading CSV file..." -ForegroundColor Cyan
try {
    $csvData = Import-Csv -Path $CSVPath -Encoding UTF8
    Write-Host "Found $($csvData.Count) rows in CSV" -ForegroundColor Green
}
catch {
    Write-Host "Error reading CSV file: $_" -ForegroundColor Red
    exit
}

# مرحله 6: دریافت تمام فایل‌های موجود در Library
Write-Host "Getting all files from library..." -ForegroundColor Cyan
$allFiles = Get-PnPListItem -List $LibraryName -PageSize 500

# مرحله 7: آپدیت هر فایل
$successCount = 0
$errorCount = 0
$notFoundCount = 0

Write-Host "`nStarting update process..." -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan

foreach ($row in $csvData) {
    # ساخت نام فایل - فرض می‌کنیم فایل‌ها با Report No شروع می‌شوند
    # می‌توانید الگوی مختلفی داشته باشید
    $reportNo = $row.'Report No'.Trim()
    
    # جستجو برای فایل‌های مرتبط (docx یا pdf)
    $matchingFiles = $allFiles | Where-Object { 
        $_.FieldValues.FileLeafRef -like "*$reportNo*"
    }
    
    if ($matchingFiles.Count -eq 0) {
        Write-Host "File not found for Report No: $reportNo" -ForegroundColor Yellow
        $notFoundCount++
        continue
    }
    
    # آپدیت هر فایل مرتبط
    foreach ($file in $matchingFiles) {
        try {
            $fileName = $file.FieldValues.FileLeafRef
            Write-Host "Updating: $fileName" -ForegroundColor White
            
            # آماده‌سازی داده‌ها برای آپدیت
            # نام ستون‌ها را با نام Internal Name های SharePoint تطبیق دهید
            $updateValues = @{}
            
            # اضافه کردن فیلدها فقط اگر مقدار داشته باشند
            if (![string]::IsNullOrWhiteSpace($row.ReportDate)) {
                # تبدیل تاریخ به فرمت مناسب
                try {
                    $date = [DateTime]::Parse($row.ReportDate)
                    $updateValues["ReportDate"] = $date.ToString("yyyy-MM-dd")
                } catch {
                    Write-Host "  Warning: Invalid date format for $reportNo" -ForegroundColor Yellow
                }
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Pttern)) {
                $updateValues["Pttern"] = $row.Pttern
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Rev)) {
                $updateValues["Rev"] = $row.Rev
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Process)) {
                $updateValues["Process"] = $row.Process
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Subprocess)) {
                $updateValues["Subprocess"] = $row.Subprocess
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Location)) {
                $updateValues["Location"] = $row.Location
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Subject)) {
                $updateValues["Subject"] = $row.Subject
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Type)) {
                $updateValues["Type"] = $row.Type
            }
            
            if (![string]::IsNullOrWhiteSpace($row.Contractor)) {
                $updateValues["Contractor"] = $row.Contractor
            }
            
            if (![string]::IsNullOrWhiteSpace($row.MainGroup)) {
                $updateValues["MainGroup"] = $row.MainGroup
            }
            
            # آپدیت فایل
            if ($updateValues.Count -gt 0) {
                Set-PnPListItem -List $LibraryName -Identity $file.Id -Values $updateValues -ErrorAction Stop
                Write-Host "  ✓ Updated successfully" -ForegroundColor Green
                $successCount++
            } else {
                Write-Host "  - No values to update" -ForegroundColor Gray
            }
            
        }
        catch {
            Write-Host "  ✗ Error updating file: $_" -ForegroundColor Red
            $errorCount++
        }
    }
}

# مرحله 8: نمایش خلاصه نتایج
Write-Host "`n================================" -ForegroundColor Cyan
Write-Host "Update Summary:" -ForegroundColor Cyan
Write-Host "  Total rows in CSV: $($csvData.Count)" -ForegroundColor White
Write-Host "  Successfully updated: $successCount" -ForegroundColor Green
Write-Host "  Files not found: $notFoundCount" -ForegroundColor Yellow
Write-Host "  Errors: $errorCount" -ForegroundColor Red
Write-Host "================================" -ForegroundColor Cyan

# مرحله 9: قطع اتصال
Disconnect-PnPOnline
Write-Host "`nDisconnected from SharePoint." -ForegroundColor Cyan