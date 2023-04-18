Dim inputStr, hexStr
inputStr = InputBox("16i”‚É•ÏŠ·‚µ‚½‚¢•¶š‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B")

For i = 1 To Len(inputStr)
	hexStr = hexStr & Right("0" & Hex(Asc(Mid(inputStr, i, 1))), 2)
Next

MsgBox("Œ‹‰Ê: " & hexStr)