Dim inputHex, outputStr
inputHex = InputBox("�ϊ�������16�i������͂��Ă��������B")

For i = 1 To Len(inputHex) Step 2
	outputStr = outputStr & Chr("&H" & Mid(inputHex, i, 2))
Next

Msgbox("����: " & outputStr)
