Dim inputStr, hexStr
inputStr = InputBox("16�i���ɕϊ���������������͂��Ă��������B")

For i = 1 To Len(inputStr)
	hexStr = hexStr & Right("0" & Hex(Asc(Mid(inputStr, i, 1))), 2)
Next

MsgBox("����: " & hexStr)