Const FILLE_NAME="addOn.xlam"

Call Exec

Sub Exec()
    Dim objExcel
    Dim strAdPath
    Dim strMyPath
    Dim strAdCp
    Dim strMyCp
    Dim objFileSys
    Dim oAdd

    Set objExcel   = CreateObject("Excel.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    strAdPath = objExcel.Application.UserLibraryPath
    strMyPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    strAdCp   = objFileSys.BuildPath(strAdPath, FILLE_NAME)
    strMyCp   = objFileSys.BuildPath(strMyPath, FILLE_NAME)

    objFileSys.CopyFile strMyCp, strAdCp

    objExcel.Workbooks.Add
    Set oAdd = objExcel.AddIns.Add(strAdCp,True)
    oAdd.Installed = True
    objExcel.Quit

    Set objExcel   = Nothing
    Set objFileSys = Nothing

    MsgBox "Complete!"
End Sub