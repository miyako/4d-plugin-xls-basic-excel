//%attributes = {}
//
//
//$wb:=XLS Create (1)
//
//XLS Set sheet name ($wb;1;"A")
//XLS Set column width ($wb;1;1;0x00AA)
//
//$filePath:=System folder(Desktop)+"test.xls"
//$success:=XLS Save as ($wb;$filePath)
//XLS CLOSE ($wb)

$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"
DOCUMENT TO BLOB:C525($filePath; $fileData)

$t:=""
For ($i; 0; BLOB size:C605($fileData)-1)
	If (($i%16)=0)
		$t:=$t+"\r"
	End if 
	
	$t:=$t+Substring:C12(String:C10($fileData{$i}; "&x"); 5; 2)+" "
	
End for 

SET TEXT TO PASTEBOARD:C523($t)

