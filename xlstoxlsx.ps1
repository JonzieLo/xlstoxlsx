if ($PSVersionTable.PSVersion -lt 5.1.19041.3570 -and -not (Get-Command -ErrorAction Ignore -Type Cmdlet Start-ThreadJob)) {
    Write-Verbose "Installing module 'ThreadJob' on demand..."
    Install-Module -ErrorAction Stop -Scope CurrentUser ThreadJob
  }
  
  # Getting here means that Start-ThreadJob is now available.
  Get-Command -Type Cmdlet Start-ThreadJob
# Specify the folder containing the .xls files
$folderPath = ""

# Get a list of .xls files in the folder
$xlsFiles = Get-ChildItem -Path $folderPath -File -Recurse | Where-Object { $_.Extension -eq ".xls" }

$scriptBlock = {
    Param($xlsFile)
    
    #Open Excel Application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false #Hides Excel Dialog Box
    $xlsxFile = [System.IO.Path]::ChangeExtension($xlsFile.FullName, ".xlsx")
    $workbook = $excel.Workbooks.Open($xlsFile.FullName)

    $format = 51
    # https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat

    $workbook.SaveAs($xlsxFile, $format)
    $workbook.Close($true)

    #Clean up COM objects to prevent memory leaks
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
}

$jobs = @()

foreach ($xlsFile in $xlsFiles) {
    $jobs += Start-ThreadJob -ScriptBlock $scriptBlock -ArgumentList $xlsFile, $excel.Workbooks -ThrottleLimit 8
}

Wait-Job -Job $jobs

foreach ($job in $jobs) {
    Receive-Job -Job $job
}

$results | ForEach-Object {Write-Host $_ }

Remove-Job -Job $jobs