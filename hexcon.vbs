Dim inputStr, hexStr
inputStr = InputBox("16進数に変換したい文字を入力してください。")

For i = 1 To Len(inputStr)
	hexStr = hexStr & Right("0" & Hex(Asc(Mid(inputStr, i, 1))), 2)
Next

MsgBox("結果: " & hexStr)