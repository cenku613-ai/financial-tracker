<#
.VERSION 2.0
.AUTHOR Hermes Agent
.DESCRIPTION Financial/Procurement Tracker - PowerShell HTTP server + Excel COM backend.
             Supports 32 fields, configurable categories/types, and a settings page.
.NOTES
  Run: .\server.ps1
  Open: http://localhost:8080
  Requirements: Windows + PowerShell 5.1+ + Excel (pre-installed)
#>

# -- Configuration ----------------------------------------------
$PORT      = 8080
$ServerUrl  = "http://localhost:$PORT/"
$SCRIPT_DIR = if ($PSScriptRoot -and $PSScriptRoot -ne '') { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Definition }
$DATA_FILE = Join-Path $SCRIPT_DIR 'data.xlsx'
$CONFIG_FILE = Join-Path $SCRIPT_DIR 'config.json'

# -- Field Schema (column index - field name) ------------------
# 1=ID, 2=Date, 3=Platform, 4=PRFNumber, 5=ProjectDescription,
# 6=SOWEffectiveDate, 7=SOWEndDate, 8=ContractNo, 9=Status,
# 10=YearofPurchasing, 11=PurchasedBy, 12=Vendor, 13=RequestTask,
# 14=PRNumber, 15=PONumber, 16=RenewForPONumber, 17=DeliveryNote,
# 18=InvoiceNo, 19=PartNumber, 20=PartNoDescription,
# 21=LicenseSubscriptionStatus, 22=Category, 23=Environment,
# 24=Type, 25=Subtype, 26=Qty, 27=UnitPriceUSD, 28=ServiceYears,
# 29=TotalAmountUSD, 30=ServicePeriodStart, 31=ServicePeriodEnd, 32=Remark
$FIELD_MAP = @{
    1  = 'ID'
    2  = 'Date'
    3  = 'Platform'
    4  = 'PRFNumber'
    5  = 'ProjectDescription'
    6  = 'SOWEffectiveDate'
    7  = 'SOWEndDate'
    8  = 'ContractNo'
    9  = 'Status'
    10 = 'YearofPurchasing'
    11 = 'PurchasedBy'
    12 = 'Vendor'
    13 = 'RequestTask'
    14 = 'PRNumber'
    15 = 'PONumber'
    16 = 'RenewForPONumber'
    17 = 'DeliveryNote'
    18 = 'InvoiceNo'
    19 = 'PartNumber'
    20 = 'PartNoDescription'
    21 = 'LicenseSubscriptionStatus'
    22 = 'Category'
    23 = 'Environment'
    24 = 'Type'
    25 = 'Subtype'
    26 = 'Qty'
    27 = 'UnitPriceUSD'
    28 = 'ServiceYears'
    29 = 'TotalAmountUSD'
    30 = 'ServicePeriodStart'
    31 = 'ServicePeriodEnd'
    32 = 'Remark'
}

# -- Helpers ----------------------------------------------------

function Send-JsonResponse {
    param($Context, $StatusCode = 200, $Body = @{})
    try {
        $json = $Body | ConvertTo-Json -Depth 4
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
        $Context.Response.StatusCode = $StatusCode
        $Context.Response.ContentType = 'application/json; charset=utf-8'
        $Context.Response.ContentLength64 = $buffer.Length
        $Context.Response.OutputStream.Write($buffer, 0, $buffer.Length) | Out-Null
        $Context.Response.Close()
    } catch { }
}

function Send-HtmlResponse {
    param($Context, $Path)
    try {
        $file = Join-Path $SCRIPT_DIR $Path
        if (Test-Path $file) {
            $buffer = [System.IO.File]::ReadAllBytes($file)
            $contentType = switch -Regex ($Path) {
                '\.html$' { 'text/html; charset=utf-8' }
                '\.css$'  { 'text/css; charset=utf-8' }
                '\.js$'   { 'application/javascript; charset=utf-8' }
                '\.png$'  { 'image/png' }
                '\.svg$'  { 'image/svg+xml' }
                '\.json$' { 'application/json' }
                default   { 'application/octet-stream' }
            }
            $Context.Response.StatusCode = 200
            $Context.Response.ContentType = $contentType
            $Context.Response.ContentLength64 = $buffer.Length
            $Context.Response.OutputStream.Write($buffer, 0, $buffer.Length) | Out-Null
            $Context.Response.Close()
        } else {
            $Context.Response.StatusCode = 404
            $Context.Response.Close()
        }
    } catch {
        $Context.Response.StatusCode = 500
        $Context.Response.Close()
    }
}

function Get-RequestBody {
    param($Context)
    $stream = $Context.Request.InputStream
    $bytes = New-Object byte[] $Context.Request.ContentLength64
    $stream.Read($bytes, 0, $bytes.Length) | Out-Null
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

# -- Config Management (config.json) ---------------------------

function Get-DefaultConfig {
    return @{
        categories = @('Hardware','Software','Service','Subscription','License','Cloud','Consulting','Maintenance')
        types      = @('Purchase','Rental','Subscription','Contract','License','Renewal')
        subtypes   = @('One-time','Recurring','Annual','Monthly','Per-project')
        statuses   = @('Draft','Submitted','Approved','In Progress','Completed','Cancelled','On Hold')
        environments = @('Production','Development','Staging','Testing','Disaster Recovery')
        licenseStatuses = @('Active','Expired','Pending','Renewed','Trial','Grace Period')
        platforms  = @('Windows','Linux','macOS','Cloud','On-premise','Hybrid','Network','Storage')
    }
}

function Get-Config {
    if (Test-Path $CONFIG_FILE) {
        try {
            return Get-Content $CONFIG_FILE -Raw | ConvertFrom-Json -ErrorAction Stop
        } catch {
            return (Get-DefaultConfig | ConvertTo-Json -Depth 2 | ConvertFrom-Json)
        }
    }
    $cfg = Get-DefaultConfig
    $cfg | ConvertTo-Json -Depth 2 | Set-Content $CONFIG_FILE
    return $cfg
}

function Save-Config {
    param($Config)
    $Config | ConvertTo-Json -Depth 2 | Set-Content $CONFIG_FILE
}

# -- Excel COM Operations ---------------------------------------

function Initialize-ExcelFile {
    if (!(Test-Path $DATA_FILE)) {
        $excel = New-Object -ComObject Excel.Application
        try {
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $wb = $excel.Workbooks.Add()
            $ws = $wb.Sheets.Item(1)
            $ws.Name = 'Transactions'
            $headers = @('ID','Date','Platform','PRFNumber','ProjectDescription',
                'SOWEffectiveDate','SOWEndDate','ContractNo','Status','YearofPurchasing',
                'PurchasedBy','Vendor','RequestTask','PRNumber','PONumber',
                'RenewForPONumber','DeliveryNote','InvoiceNo','PartNumber','PartNoDescription',
                'LicenseSubscriptionStatus','Category','Environment','Type','Subtype',
                'Qty','UnitPriceUSD','ServiceYears','TotalAmountUSD',
                'ServicePeriodStart','ServicePeriodEnd','Remark')
            for ($i = 1; $i -le $headers.Count; $i++) {
                $ws.Cells.Item(1, $i).Value2 = $headers[$i - 1]
                $ws.Cells.Item(1, $i).Font.Bold = $true
            }
            $wb.SaveAs($DATA_FILE, 51)
        } finally {
            $wb.Close($true)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

function Get-NextId {
    $maxId = 0
    $excel = New-Object -ComObject Excel.Application
    try {
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($DATA_FILE)
        $ws = $wb.Sheets.Item('Transactions')
        for ($r = 2; $r -le 10000; $r++) {
            $val = $ws.Cells.Item($r, 1).Value2
            if ($val -eq $null -or $val -eq '') { break }
            $numId = [double]::Parse($val)
            if ($numId -gt $maxId) { $maxId = $numId }
        }
        $wb.Close($false)
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    return ($maxId + 1)
}

function Row-To-Object {
    param($ws, $row)
    return [PSCustomObject]@{
        ID                      = $ws.Cells.Item($row, 1).Value2
        Date                    = $ws.Cells.Item($row, 2).Value2
        Platform                = $ws.Cells.Item($row, 3).Value2
        PRFNumber               = $ws.Cells.Item($row, 4).Value2
        ProjectDescription      = $ws.Cells.Item($row, 5).Value2
        SOWEffectiveDate        = $ws.Cells.Item($row, 6).Value2
        SOWEndDate              = $ws.Cells.Item($row, 7).Value2
        ContractNo              = $ws.Cells.Item($row, 8).Value2
        Status                  = $ws.Cells.Item($row, 9).Value2
        YearofPurchasing        = $ws.Cells.Item($row, 10).Value2
        PurchasedBy             = $ws.Cells.Item($row, 11).Value2
        Vendor                  = $ws.Cells.Item($row, 12).Value2
        RequestTask             = $ws.Cells.Item($row, 13).Value2
        PRNumber                = $ws.Cells.Item($row, 14).Value2
        PONumber                = $ws.Cells.Item($row, 15).Value2
        RenewForPONumber        = $ws.Cells.Item($row, 16).Value2
        DeliveryNote            = $ws.Cells.Item($row, 17).Value2
        InvoiceNo               = $ws.Cells.Item($row, 18).Value2
        PartNumber              = $ws.Cells.Item($row, 19).Value2
        PartNoDescription       = $ws.Cells.Item($row, 20).Value2
        LicenseSubscriptionStatus = $ws.Cells.Item($row, 21).Value2
        Category                = $ws.Cells.Item($row, 22).Value2
        Environment             = $ws.Cells.Item($row, 23).Value2
        Type                    = $ws.Cells.Item($row, 24).Value2
        Subtype                 = $ws.Cells.Item($row, 25).Value2
        Qty                     = $ws.Cells.Item($row, 26).Value2
        UnitPriceUSD            = $ws.Cells.Item($row, 27).Value2
        ServiceYears            = $ws.Cells.Item($row, 28).Value2
        TotalAmountUSD          = $ws.Cells.Item($row, 29).Value2
        ServicePeriodStart      = $ws.Cells.Item($row, 30).Value2
        ServicePeriodEnd        = $ws.Cells.Item($row, 31).Value2
        Remark                  = $ws.Cells.Item($row, 32).Value2
    }
}

function Object-To-Row {
    param($ws, $row, $obj)
    $fields = @(
        $obj.ID,$obj.Date,$obj.Platform,$obj.PRFNumber,$obj.ProjectDescription,
        $obj.SOWEffectiveDate,$obj.SOWEndDate,$obj.ContractNo,$obj.Status,$obj.YearofPurchasing,
        $obj.PurchasedBy,$obj.Vendor,$obj.RequestTask,$obj.PRNumber,$obj.PONumber,
        $obj.RenewForPONumber,$obj.DeliveryNote,$obj.InvoiceNo,$obj.PartNumber,$obj.PartNoDescription,
        $obj.LicenseSubscriptionStatus,$obj.Category,$obj.Environment,$obj.Type,$obj.Subtype,
        $obj.Qty,$obj.UnitPriceUSD,$obj.ServiceYears,$obj.TotalAmountUSD,
        $obj.ServicePeriodStart,$obj.ServicePeriodEnd,$obj.Remark
    )
    for ($i = 1; $i -le 32; $i++) {
        $val = $fields[$i - 1]
        if ($val -is [String]) { $ws.Cells.Item($row, $i).Value2 = $val }
        elseif ($val -ne $null) {
            if (($i -eq 26) -or ($i -eq 27) -or ($i -eq 28) -or ($i -eq 29)) {
                $ws.Cells.Item($row, $i).Value2 = [double]::Parse($val)
            } else {
                $ws.Cells.Item($row, $i).Value2 = "$val"
            }
        } else {
            $ws.Cells.Item($row, $i).Value2 = ''
        }
    }
}

function Open-Excel {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($DATA_FILE)
    $ws = $wb.Sheets.Item('Transactions')
    return @{ Excel = $excel; Wb = $wb; Ws = $ws }
}

function Close-Excel {
    param($handle)
    $handle.Wb.Close($false)
    $handle.Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($handle.Ws) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($handle.Wb) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($handle.Excel) | Out-Null
}

function Read-Transactions {
    $handle = Open-Excel
    $transactions = @()
    try {
        $ws = $handle.Ws
        for ($r = 2; $r -le 10000; $r++) {
            $id = $ws.Cells.Item($r, 1).Value2
            if ($id -eq $null -or $id -eq '') { break }
            $transactions += Row-To-Object -ws $ws -row $r
        }
    } finally { Close-Excel -handle $handle }
    return $transactions
}

function Add-Transaction {
    param($Body)
    $nextId = Get-NextId
    $handle = Open-Excel
    try {
        $ws = $handle.Ws
        $nextRow = 2
        while ($true) {
            $val = $ws.Cells.Item($nextRow, 1).Value2
            if ($val -eq $null -or $val -eq '') { break }
            $nextRow++
        }
        $obj = [PSCustomObject]@{
            ID                      = $nextId
            Date                    = $Body.Date
            Platform                = $Body.Platform
            PRFNumber               = $Body.PRFNumber
            ProjectDescription      = $Body.ProjectDescription
            SOWEffectiveDate        = $Body.SOWEffectiveDate
            SOWEndDate              = $Body.SOWEndDate
            ContractNo              = $Body.ContractNo
            Status                  = $Body.Status
            YearofPurchasing        = $Body.YearofPurchasing
            PurchasedBy             = $Body.PurchasedBy
            Vendor                  = $Body.Vendor
            RequestTask             = $Body.RequestTask
            PRNumber                = $Body.PRNumber
            PONumber                = $Body.PONumber
            RenewForPONumber        = $Body.RenewForPONumber
            DeliveryNote            = $Body.DeliveryNote
            InvoiceNo               = $Body.InvoiceNo
            PartNumber              = $Body.PartNumber
            PartNoDescription       = $Body.PartNoDescription
            LicenseSubscriptionStatus = $Body.LicenseSubscriptionStatus
            Category                = $Body.Category
            Environment             = $Body.Environment
            Type                    = $Body.Type
            Subtype                 = $Body.Subtype
            Qty                     = $Body.Qty
            UnitPriceUSD            = $Body.UnitPriceUSD
            ServiceYears            = $Body.ServiceYears
            TotalAmountUSD          = $Body.TotalAmountUSD
            ServicePeriodStart      = $Body.ServicePeriodStart
            ServicePeriodEnd        = $Body.ServicePeriodEnd
            Remark                  = $Body.Remark
        }
        Object-To-Row -ws $ws -row $nextRow -obj $obj
        $handle.Wb.Save()
    } finally { Close-Excel -handle $handle }
    return @{ id = $nextId; message = 'Transaction added' }
}

function Update-Transaction {
    param($Id, $Body)
    $handle = Open-Excel
    $found = $false
    try {
        $ws = $handle.Ws
        for ($r = 2; $r -le 10000; $r++) {
            $val = $ws.Cells.Item($r, 1).Value2
            if ($val -eq $null -or $val -eq '') { break }
            if ([string]$val -eq [string]$Id) {
                $obj = [PSCustomObject]@{
                    ID                      = $Id
                    Date                    = $Body.Date
                    Platform                = $Body.Platform
                    PRFNumber               = $Body.PRFNumber
                    ProjectDescription      = $Body.ProjectDescription
                    SOWEffectiveDate        = $Body.SOWEffectiveDate
                    SOWEndDate              = $Body.SOWEndDate
                    ContractNo              = $Body.ContractNo
                    Status                  = $Body.Status
                    YearofPurchasing        = $Body.YearofPurchasing
                    PurchasedBy             = $Body.PurchasedBy
                    Vendor                  = $Body.Vendor
                    RequestTask             = $Body.RequestTask
                    PRNumber                = $Body.PRNumber
                    PONumber                = $Body.PONumber
                    RenewForPONumber        = $Body.RenewForPONumber
                    DeliveryNote            = $Body.DeliveryNote
                    InvoiceNo               = $Body.InvoiceNo
                    PartNumber              = $Body.PartNumber
                    PartNoDescription       = $Body.PartNoDescription
                    LicenseSubscriptionStatus = $Body.LicenseSubscriptionStatus
                    Category                = $Body.Category
                    Environment             = $Body.Environment
                    Type                    = $Body.Type
                    Subtype                 = $Body.Subtype
                    Qty                     = $Body.Qty
                    UnitPriceUSD            = $Body.UnitPriceUSD
                    ServiceYears            = $Body.ServiceYears
                    TotalAmountUSD          = $Body.TotalAmountUSD
                    ServicePeriodStart      = $Body.ServicePeriodStart
                    ServicePeriodEnd        = $Body.ServicePeriodEnd
                    Remark                  = $Body.Remark
                }
                Object-To-Row -ws $ws -row $r -obj $obj
                $handle.Wb.Save()
                $found = $true
                break
            }
        }
    } finally { Close-Excel -handle $handle }
    return $found
}

function Remove-Transaction {
    param($Id)
    $handle = Open-Excel
    $found = $false
    try {
        $ws = $handle.Ws
        for ($r = 2; $r -le 10000; $r++) {
            $val = $ws.Cells.Item($r, 1).Value2
            if ($val -eq $null -or $val -eq '') { break }
            if ([string]$val -eq [string]$Id) {
                $ws.Rows.Item("${r}:${r}").Delete()
                $handle.Wb.Save()
                $found = $true
                break
            }
        }
    } finally { Close-Excel -handle $handle }
    return $found
}

function Get-Transaction {
    param($Id)
    $handle = Open-Excel
    $tx = $null
    try {
        $ws = $handle.Ws
        for ($r = 2; $r -le 10000; $r++) {
            $val = $ws.Cells.Item($r, 1).Value2
            if ($val -eq $null -or $val -eq '') { break }
            if ([string]$val -eq [string]$Id) {
                $tx = Row-To-Object -ws $ws -row $r
                break
            }
        }
    } finally { Close-Excel -handle $handle }
    return $tx
}

# -- Initialize -------------------------------------------------
Initialize-ExcelFile
Get-Config | Out-Null  # Creates config.json if missing

# -- HTTP Server ------------------------------------------------
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($ServerUrl)
$listener.Start()
Write-Host ''
Write-Host '+--------------------------------------------------+' -ForegroundColor Cyan
Write-Host '|   Financial Tracker v2.0 - Server Started        |' -ForegroundColor Cyan
Write-Host "|   http://localhost:$PORT                         |" -ForegroundColor Cyan
Write-Host "|   Data: $DATA_FILE --" -ForegroundColor Cyan
Write-Host "|   Config: $CONFIG_FILE --" -ForegroundColor Cyan
Write-Host '|   Press Ctrl+C to stop                           |' -ForegroundColor Cyan
Write-Host '+--------------------------------------------------+' -ForegroundColor Cyan

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $path = $context.Request.Url.AbsolutePath.TrimEnd('/')
        $method = $context.Request.HttpMethod
        $queryString = $context.Request.Url.Query

        if ($method -eq 'OPTIONS') {
            $context.Response.StatusCode = 200
            $context.Response.Close()
            continue
        }

        # -- Config API --
        if ($path -eq '/api/config') {
            if ($method -eq 'GET') {
                Send-JsonResponse -Context $context -Body (Get-Config)
            } elseif ($method -eq 'PUT') {
                $raw = Get-RequestBody -Context $context
                $cfg = $raw | ConvertFrom-Json
                Save-Config -Config $cfg
                Send-JsonResponse -Context $context -Body @{ message = 'Config saved' }
            }
            continue
        }

        # -- Transactions List --
        if ($path -eq '/api/transactions') {
            if ($method -eq 'GET') {
                $all = Read-Transactions
                if ($queryString) {
                    $params = [System.Web.HttpUtility]::ParseQueryString($queryString)
                    if ($params['category']) { $all = $all | Where-Object { $_.Category -eq $params['category'] } }
                    if ($params['type']) { $all = $all | Where-Object { $_.Type -eq $params['type'] } }
                    if ($params['status']) { $all = $all | Where-Object { $_.Status -eq $params['status'] } }
                    if ($params['vendor']) { $all = $all | Where-Object { $_.Vendor -like "*$($params['vendor'])*" } }
                    if ($params['search']) {
                        $s = $params['search']
                        $all = $all | Where-Object {
                            ($_.ProjectDescription -and $_.ProjectDescription -like "*$s*") -or
                            ($_.PONumber -and $_.PONumber -like "*$s*") -or
                            ($_.Vendor -and $_.Vendor -like "*$s*") -or
                            ($_.Category -and $_.Category -like "*$s*") -or
                            ($_.PRNumber -and $_.PRNumber -like "*$s*")
                        }
                    }
                }
                Send-JsonResponse -Context $context -Body @{ transactions = $all }
            } elseif ($method -eq 'POST') {
                $raw = Get-RequestBody -Context $context
                $body = $raw | ConvertFrom-Json
                $result = Add-Transaction -Body $body
                Send-JsonResponse -Context $context -StatusCode 201 -Body $result
            }
            continue
        }

        # -- Single Transaction --
        if ($path -match '^/api/transactions/(\d+)$') {
            $id = $Matches[1]
            if ($method -eq 'GET') {
                $tx = Get-Transaction -Id $id
                if ($tx) { Send-JsonResponse -Context $context -Body $tx }
                else { Send-JsonResponse -Context $context -StatusCode 404 -Body @{ message = 'Not found' } }
            } elseif ($method -eq 'PUT') {
                $raw = Get-RequestBody -Context $context
                $body = $raw | ConvertFrom-Json
                $result = Update-Transaction -Id $id -Body $body
                if ($result) { Send-JsonResponse -Context $context -Body @{ message = 'Updated' } }
                else { Send-JsonResponse -Context $context -StatusCode 404 -Body @{ message = 'Not found' } }
            } elseif ($method -eq 'DELETE') {
                $result = Remove-Transaction -Id $id
                if ($result) { Send-JsonResponse -Context $context -Body @{ message = 'Deleted' } }
                else { Send-JsonResponse -Context $context -StatusCode 404 -Body @{ message = 'Not found' } }
            }
            continue
        }

        # -- Summary --
        if ($path -eq '/api/summary') {
            if ($method -eq 'GET') {
                $all = Read-Transactions
                $totalAmt = ($all | Where-Object { $_.TotalAmountUSD } | Measure-Object -Property TotalAmountUSD -Sum).Sum
                $totalAmt = if ($totalAmt) { [math]::Round($totalAmt, 2) } else { 0 }
                $categories = $all | Where-Object { $_.Category } | Group-Object -Property Category | ForEach-Object {
                    $cat = $_
                    [PSCustomObject]@{
                        Category = $cat.Name
                        Count = $cat.Count
                        TotalAmount = [math]::Round(($cat.Group | Measure-Object -Property TotalAmountUSD -Sum).Sum, 2)
                    }
                }
                $statuses = $all | Where-Object { $_.Status } | Group-Object -Property Status | ForEach-Object {
                    @{ Name = $_.Name; Count = $_.Count }
                }
                $vendors = $all | Where-Object { $_.Vendor } | Group-Object -Property Vendor | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {
                    @{ Name = $_.Name; Count = $_.Count }
                }
                Send-JsonResponse -Context $context -Body @{
                    totalTransactions = $all.Count
                    totalAmountUSD    = $totalAmt
                    categories        = $categories
                    statuses          = $statuses
                    topVendors        = $vendors
                }
            }
            continue
        }

        # -- Static Files --
        if ($path -eq '') { Send-HtmlResponse -Context $context -Path 'index.html' }
        else {
            $fileName = $path.Substring(1)
            if ($fileName -eq '') { $fileName = 'index.html' }
            Send-HtmlResponse -Context $context -Path $fileName
        }
    }
} finally {
    $listener.Stop()
    Write-Host 'Server stopped.' -ForegroundColor Yellow
}
