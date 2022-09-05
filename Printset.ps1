$Printers = Get -WmiObject Win32_Printer
$Printer = $Printers | Where-Object Name -eq "DocuCentre-V C5575 T2(1)"

$Printer.SetDefaultPrinter()
$Printer.SetDefaultPaperType = "A4"

$Printer.Put