Call callCheck

Sub callCheck()
	Dim xbk As Workbook
	Set xbk = Workbooks("addOn.xlam")
	If xbk Is Nothing Then
	MsgBox "�J����Ă��܂���"
	Else
	MsgBox "�J����Ă��܂�"
	End If

End Sub