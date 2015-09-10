param(
    [String]$FilePath,
    [int]$HeaderRow=1,
    [string]$Worksheet,
    [string]$ExcludeWS,
    [string]$Password
)

$Missing = [System.Reflection.Missing]::Value

# Try to open the Excel doc. Use a password if necessary
# Return an error if encountered and stop the script
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

#
if ($Worksheet) {
    $Worksheets = $WorkBook.Sheets.Item($Worksheet)
}
elseif ($ExcludeWS) {
    $Worksheets = $WorkBook.Worksheets | where {$_.Name -ne $ExcludeWS}
}
else {
    $Worksheets = $WorkBook.Worksheets
}

foreach ($WS in $Worksheets) {
    $Sheet = $WorkBook.Sheets.Item($WS.Name)
    $MaxCols = $Sheet.UsedRange.Columns.Count
    $MaxRows = $Sheet.UsedRange.Rows.Count
    for ($Col = 1; $Col -le $MaxCols; $Col++) {
        $Header = $Sheet.Cells.Item.Invoke($HeaderRow, $Col).Value2
        $Header = $Header.Replace(" ", "")
        [array]$Headers += $Header
    }
    for ($Row = ($HeaderRow + 1); $Row -le $MaxRows; $Row++) {
        $RowObj = New-Object PSObject
        for ($Col = $HeaderRow; $Col -le $MaxCols; $Col++) {
            $RowObj | Add-Member -MemberType NoteProperty -Name $Headers[$Col - 1] -Value $($Sheet.Cells.Item.Invoke($Row, $Col).Value2)
        }
        [array]$AllObjects += $RowObj
    }
}

return $AllObjects