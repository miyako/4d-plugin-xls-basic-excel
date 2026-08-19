//%attributes = {}
//create a blank book with n sheets
$numberOfSheets:=3
$wb:=XLS Create($numberOfSheets)  //min=1, max=10

$sheet:=1
$row:=1
$col:=1
$width:=120*20  //in twips (20th of a point)
XLS Set text value($wb; $sheet; $row; $col; "1,1:abcde")
XLS Set column width($wb; $sheet; $col; $width)
XLS Set sheet name($wb; $sheet; "##SHEET1##")

$font:="Lucida Grande"
$height:=12*20  //in twips (20th of a point)
$color:=XLS Color Cyan
$weight:=XLS Font Weight Bold
$option:=XLS Font Italic
$underline:=XLS Underline Single
$family:=XLS Font Family Roman
$escape:=XLS Escapement Normal
XLS SET FONT PROPERTY($wb; $sheet; $row; $col; \
$font; $height; $color; $weight; $option; $underline; $family; $escape)

XLS Set text value($wb; $sheet; 2; 2; "2,2:ghijk")

$alignment:=XLS H Align Centered | XLS V Align Centered
$rotation:=45
$properties:=XLS Text Default
$top:=XLS Border Dashed
$topColor:=XLS Color Red
$left:=XLS Border Dashed
$leftColor:=XLS Color Red
$right:=XLS Border Dashed
$rightColor:=XLS Color Red
$bottom:=XLS Border Dashed
$bottomColor:=XLS Color Red
$pattern:=XLS Pattern Solid
$patternColor:=XLS Color Blue
$patternAltColor:=XLS Color Yellow
XLS SET FORMAT PROPERTY($wb; $sheet; 2; 2; \
$alignment; $rotation; $properties; \
$top; $topColor; \
$left; $leftColor; \
$right; $rightColor; \
$bottom; $bottomColor; \
$pattern; $patternColor; $patternAltColor)

$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"
$success:=XLS Save as($wb; $filePath)

XLS CLOSE($wb)
