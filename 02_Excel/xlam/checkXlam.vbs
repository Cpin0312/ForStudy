Call callCheck

Sub callCheck()
	Dim xbk As Workbook
	Set xbk = Workbooks("addOn.xlam")
	If xbk Is Nothing Then
	MsgBox "ŠJ‚©‚ê‚Ä‚¢‚Ü‚¹‚ñ"
	Else
	MsgBox "ŠJ‚©‚ê‚Ä‚¢‚Ü‚·"
	End If

End Sub