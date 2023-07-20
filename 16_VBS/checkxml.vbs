Option Explicit
Dim objXML, rPath, i, strResult, objErr

If wscript.arguments.count > 0 Then
	For i = 0 to wscript.arguments.length - 1
		rPath = wscript.arguments.item(i)

		If LCASE(RIGHT(rPath, 4)) = ".xml" Then
			Set objXML = CreateObject("MSXML2.DOMDocument.6.0")

			objXML.setProperty "ProhibitDTD", False
			objXML.setProperty "ResolveExternals", True 
			objXML.validateOnParse = True

			objXML.async = False
			On Error Resume Next

			objXML.load(rPath)
			If objXML.parseError.errorCode <> 0 Then
				Set objErr = objXML.parseError
				strResult = objErr.reason
				strResult = strResult & vbNewLine
				strResult = strResult & objErr.line & "行目に記載されている次の記述を確認してください。"
				strResult = strResult & vbNewLine
				strResult = strResult & " " & objErr.srcText
				strResult = strResult & vbNewLine
				WScript.Echo strResult
			Else
				wscript.echo "true"
			End If

			On Error Goto 0
			Set objXML = Nothing
		End If

	Next
End If



