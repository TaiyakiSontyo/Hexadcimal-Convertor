Dim inputHex, outputStr
inputHex = InputBox("変換したい16進数を入力してください。")

For i = 1 To Len(inputHex) Step 2
	outputStr = outputStr & Chr("&H" & Mid(inputHex, i, 2))
Next

Msgbox("結果: " & outputStr)
