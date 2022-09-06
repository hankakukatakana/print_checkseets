$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$book = $excel.Workbooks.Open("C:\...\Book1.xlsx")
$sheet = $excel.Worksheets.Item("Sheet1")
 
#開始ページと終了ページ、部数を指定して印刷（シートを選択する方法が不明）
$From = 1 #開始ページ
$To = 1 #終了ページ
$Copies = 1 #部数
$book.PrintOut.Invoke(@($From, $To, $Copies))
 
# 閉じる
$book.Close()
$excel.Quit()
#↓これ忘れずに。$excel.Quit()だけではプロセスは落ちない。
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)