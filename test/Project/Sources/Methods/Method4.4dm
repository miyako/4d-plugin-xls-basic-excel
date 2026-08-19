//%attributes = {}
//create a blank book with n sheets
$numberOfSheets:=3
$wb:=XLS Create($numberOfSheets)  //min=1, max=10

$sheet:=1
$row:=1
$col:=1
$width:=120*20  //in twips (20th of a point)
XLS Set text value($wb; $sheet; $row; $col; "abcde"*100)
XLS SET WRAPPING($wb; $sheet; $row; $col; 1)
XLS Set sheet name($wb; $sheet; "##WRAPPING##")

$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"
$success:=XLS Save as($wb; $filePath)

XLS CLOSE($wb)