//%attributes = {}
//create a blank book with n sheets
$numberOfSheets:=3
$wb:=XLS Create($numberOfSheets)  //min=1, max=10

$sheet:=1
$row:=1
$col:=1
$width:=240*20  //in twips (20th of a point)

$alignment:=XLS H Align General
$rotation:=0
$properties:=XLS Text Default
$top:=XLS Border None
$topColor:=XLS Color Red
$left:=XLS Border None
$leftColor:=XLS Color Red
$right:=XLS Border None
$rightColor:=XLS Color Red
$bottom:=XLS Border None
$bottomColor:=XLS Color Red
$pattern:=XLS Pattern Solid
$patternColor:=XLS Color Red
$patternAltColor:=XLS Color Yellow

XLS SET FORMAT PROPERTY($wb; $sheet; 2; 2; \
$alignment; $rotation; $properties; \
$top; $topColor; \
$left; $leftColor; \
$right; $rightColor; \
$bottom; $bottomColor; \
$pattern; $patternColor; $patternAltColor)

XLS Set real value($wb; $sheet; 2; 2; 1234.5678)
$success:=XLS Set format string($wb; $sheet; 2; 2; "\"$\"#,##0_);(\"$\"#,##0)")
$success:=XLS Get format string($wb; $sheet; 2; 2; $format)

$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"
$success:=XLS Save as($wb; $filePath)

XLS CLOSE($wb)