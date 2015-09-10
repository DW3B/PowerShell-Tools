param(
    [Parameter(Mandatory=$True)][string]$FilePath,
    [int]$HeaderRow=1,
    [string]$Worksheet,
    [switch]$ShowProgress,
    [string]$Password
)

$StartTime = Get-Date # TIMER FOR TESTING. REMOVE!

$Missing = [System.Reflection.Missing]::Value

try {
    $Excel = New-Object -ComObject Excel.Application
    if ($Password) {
        $WorkBook = $Excel.Workbooks.Open($FilePath,$Missing,$Missing,$Missing,$Password,$Password)
    }
    else {
        $WorkBook = $Excel.Workbooks.Open($FilePath)
    }
}
catch {
    $Error[0]
    break
}

# If a worksheet name was specified, use only that worksheet. Otherwise, use 'em all!
if ($Worksheet) {
    $Worksheets = $WorkBook.Sheets.Item($Worksheet)
}
else {
    $Worksheets = $WorkBook.Worksheets
}

$ScriptBlock = {
    param($Excel, $WorkBook, $WorkSheets, $WS, $HeaderRow, [int]$Start, [int]$End, $MaxCols)
    
    for ($Row = $Start; $Row -le $End; $Row++) {
        $RowObj = New-Object PSObject
        for ($Col = 1; $Col -le $MaxCols; $Col++) {
            $RowObj | Add-Member -MemberType NoteProperty -Name $($WS.Cells.Item.Invoke($HeaderRow, $Col).Value2) -Value $($WS.Cells.Item.Invoke($Row, $Col).Value2)
        }
        [array]$AllObjects += $RowObj
    }
    return $AllObjects
}

function Get-NumberRanges {
    param([int]$Start, [int]$Size, [int]$Max)
    
    while ($Start -lt $Max) {
        $RStart = $Start
        $REnd   = $Start + ($Size - 1)
        if ($REnd -gt $Max) {
            $REnd = $Max
        }
        $RangeObj = New-Object PSObject -Property @{
            "Start" = $RStart
            "End"   = $REnd    
        }
        [array]$AllRangeObj += $RangeObj
        $Start += $Size
    }
    
    return $AllRangeObj
}

function MultiThread-Excel {
    param($XLDoc, $WorkBook, $Worksheets, $WS, [int]$Rows, [int]$Columns, [int]$MaxThreads, [int]$HeaderRow)

    $ThreadSize = [Math]::Ceiling($Rows / $MaxThreads)
    $RowRanges  = Get-NumberRanges -Start ($HeaderRow + 1) -Size $ThreadSize -Max $Rows

    $ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads, $ISS, $Host)
    $RunspacePool.Open()

    foreach ($Range in $RowRanges) {
        $PSThread = [PowerShell]::Create().AddScript($ScriptBlock)
        
        $PSThread.AddArgument($XLDoc)
        $PSThread.AddArgument($WorkBook)
        $PSThread.AddArgument($Worksheets)
        $PSThread.AddArgument($WS)
        $PSThread.AddArgument($HeaderRow) 
        $PSThread.AddArgument($Range.Start)
        $PSThread.AddArgument($Range.End)
        $PSThread.AddArgument($Columns)

        [array]$Results += $PSThread.BeginInvoke()
    }
    return $Results
}

foreach ($WS in $Worksheets) {
    $MaxCols = $WS.UsedRange.Columns.Count
    $MaxRows = $WS.UsedRange.Rows.Count
    MultiThread-Excel -XLDoc $Excel -WorkBook $WorkBook -Worksheets $Worksheets -WS $WS -Rows $MaxCols -Columns $MaxCols -MaxThreads 20 -HeaderRow $HeaderRow
    <#for ($Row = ($HeaderRow + 1); $Row -le $MaxRows; $Row++) {
        if ($ShowProgress) {
            $Percent = [Math]::Round(($Row / $MaxRows) * 100, 2)
            Write-Progress -Activity "Creating Row Objects" -Status "PercentComplete: $Percent%" -PercentComplete $Percent
        }
        $RowObj = New-Object PSObject
        for ($Col = $HeaderRow; $Col -le $MaxCols; $Col++) {
            $RowObj | Add-Member -MemberType NoteProperty -Name $($WS.Cells.Item.Invoke($HeaderRow, $Col).Value2) -Value $($WS.Cells.Item.Invoke($Row, $Col).Value2)
        }
        [array]$AllObjects += $RowObj
    }#>
}
# TIMER FOR TESTING. REMOVE!
$End = Get-Date 
$TimeDiff = New-TimeSpan -Start $Start -End $End
Write-Host "$($TimeDiff.Minutes)min $($TimeDiff.Seconds)s $($TimeDiff.Milliseconds)msec"
#return $AllObjects