Call callCheck

Sub callCheck()
	Dim xbk As Workbook
	Set xbk = Workbooks("addOn.xlam")
	If xbk Is Nothing Then
	MsgBox "開かれていません"
	Else
	MsgBox "開かれています"
	End If

End Sub