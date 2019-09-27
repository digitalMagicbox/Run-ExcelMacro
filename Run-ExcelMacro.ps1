Param(
    [string]$excel_file_name # マクロを実行するエクセルファイル名
    ,[string]$macro_name     # 実行するマクロ名
)

if( [String]::IsNullOrEmpty($excel_file_name) -or 
    [String]::IsNullOrEmpty($macro_name) ) {
    # 空のときは終わり
    Write-Host $excel_file_name is not found! 
    exit 1
}

if(Test-Path $excel_file_name) {
    $excel = New-Object -ComObject Excel.Application
    # Excelを表示しない
    $excel.Visible = $false
    $full_name = Convert-Path $excel_file_name
    $book = $excel.Workbooks.Open($full_name)
    $excel.run($macro_name)
    $book.close()
    $book = $null
    $excel.quit()
    $excel = $null
    [GC]::Collect()

    exit 0
}