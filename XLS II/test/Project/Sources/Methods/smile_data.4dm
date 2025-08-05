//%attributes = {}
$numberOfSheets:=3
$wb:=XLS Create($numberOfSheets)

XLS SET FORMAT PROPERTY($wb; 1; 1; 1; \
$alignment; $rotation; $properties; \
$top; $topColor; \
$left; $leftColor; \
$right; $rightColor; \
$bottom; $bottomColor; \
$pattern; $patternColor; $patternAltColor)

$success:=XLS Set real value($wb; 1; 1; 1; 1234)
$success:=XLS Set format string($wb; 1; 1; 1; "\"$\"#,##0_);(\"$\"#,##0)")

$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"
$success:=XLS Save as($wb; $filePath)

XLS CLOSE($wb)

SHOW ON DISK:C922($filePath)