$path = "C:\Users\Timothy.MacLatchy\Journe Brands\Journe Finance - Documents"
$excelFiles = Get-ChildItem -Path $path -Include *.xls, *.xlsx -Recurse
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

foreach ($file in $excelFiles) {
  $wb = $excel.Workbooks.Open($file.FullName, 0, $true)  # don't update links
  $links = $wb.LinkSources(1)  # xlExcelLinks = external workbook links
  if ($links) {
    Write-Host "File:" $file.FullName
    foreach ($link in $links) {
      $status = $wb.LinkInfo($link, 1)  # get status code
      Write-Host "  Link:" $link "Status:" $status
    }
  }
  $wb.Close($false)
}

$excel.Quit()
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
