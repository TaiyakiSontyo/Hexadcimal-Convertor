Dim inputHex, outputStr
inputHex = InputBox("•ÏŠ·‚µ‚½‚¢16i”‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B")

For i = 1 To Len(inputHex) Step 2
	outputStr = outputStr & Chr("&H" & Mid(inputHex, i, 2))
Next

Msgbox("Œ‹‰Ê: " & outputStr)
