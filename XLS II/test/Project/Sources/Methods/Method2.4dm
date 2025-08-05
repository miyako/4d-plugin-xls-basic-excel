//%attributes = {}
$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"

$wb:=XLS Load($filePath)

$sheet:=1
$row:=1
$col:=1
$text:=XLS Get text value($wb; $sheet; $row; $col)
$real:=XLS Get real value($wb; $sheet; $row; $col)
$long:=XLS Get long value($wb; $sheet; $row; $col)
$width:=XLS Get column width($wb; $sheet; $row; $col)
$type:=XLS Get value type($wb; $sheet; $row; $col)
XLS Get sheet name($wb; $sheet; $name)
$cols:=XLS Get total columns($wb; $sheet)
$rows:=XLS Get total rows($wb; $sheet)
$sheets:=XLS Get total sheets($wb)

$font:=""
$height:=0
$color:=0
$weight:=0
$option:=0
$underline:=0
$family:=0
$escape:=0
XLS GET FONT PROPERTY($wb; $sheet; $row; $col; \
$font; $height; $color; $weight; $option; $underline; $family; $escape)

$alignment:=0
$rotation:=0
$properties:=0
$top:=0
$topColor:=0
$left:=0
$leftColor:=0
$right:=0
$rightColor:=0
$bottom:=0
$bottomColor:=0
$pattern:=0
$patternColor:=0
$patternAltColor:=0
XLS GET FORMAT PROPERTY($wb; $sheet; 2; 2; \
$alignment; $rotation; $properties; \
$top; $topColor; \
$left; $leftColor; \
$right; $rightColor; \
$bottom; $bottomColor; \
$pattern; $patternColor; $patternAltColor)

$success:=XLS Save as($wb; $filePath)

XLS CLOSE($wb)

