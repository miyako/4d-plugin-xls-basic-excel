//%attributes = {}
$vl_workbook:=XLS Create(1)

XLS Set real value($vl_workbook; 1; 1; 1; 31650.6)

$vt_path:=System folder:C487(Desktop:K41:16)+"test.xls"

$vl_result:=XLS Save as($vl_workbook; $vt_path)

XLS CLOSE($vl_workbook)
