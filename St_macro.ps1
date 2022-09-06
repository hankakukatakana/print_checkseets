# Excelを操作する為の宣言
$excel = New-Object -ComObject Excel.Application
# 可視化しない
$excel.Visible = $false
# 既存のワークブックを開く場合
$book = $excel.Workbooks.Open("F:\usr\PowerShell\sample_07\Test\PS-Test05.xlsm")
# ワークシートを番号で指定し、接続する
$sheet = $excel.Worksheets.Item(1)
# Macro2 を実行する。
$excel.Run( "Macro2" , "PowerShell からマクロを実行しています。" )
# Excelを閉じる
$excel.Quit()
# プロセスを解放する
$excel = $null
[GC]::Collect()
