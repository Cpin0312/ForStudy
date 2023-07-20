'Create Progress Window and bring progress bar to front


'Set ObjProgWindow = new ProgressWindow
'ObjProgWindow.Initialise "Processing...", 4, 500, 100
'ObjProgWindow.UpdateRegionProgress 1, 50,30
'WScript.Echo "Done"
'ObjProgWindow.UpdateRegionProgress 1, 60,30
'WScript.Echo "Done"
'ObjProgWindow.UpdateRegionProgress 1, 80,30
'WScript.Echo "Done"
'ObjProgWindow.UpdateRegionProgress 1, 100,30
'WScript.Echo "Done"
'ObjProgWindow.Destroy

'################
'#### CLASSES ###
'################
'Create an instance of this class to create a progress window based on an HTML page
'displayed in an Internet Explorer window.
Class ProgressWindow
 
    'Progress Window Properties
    Dim intWidth
    Dim intHeight
    Dim intRegions
    Dim strTitle
    Dim objProgress
 
    'Create the progress window
    'The first parameter is the title of the progress window.
    'The second parameter is the number of regions on the page you wish to be able to update independently.
    'The third parameter is the width of the window in pixels.
    'The fourth parameter is the height of the window in pixels.
    Sub Initialise(p_strTitle, p_intRegions, p_intWidth, p_intHeight)
        'Set class properties
        strTitle = p_strTitle
        intRegions = p_intRegions
        intWidth = p_intWidth
        intHeight = p_intHeight
 
        'Create the progress window
        Set objProgress = CreateMSIEWindow()
    End Sub
 
    'Display the basic window
    Function CreateMSIEWindow()
    	Dim p_objMSIE
        Dim strHTML
        Dim intCounter
 
        strHTML = ""
        For intCounter = 1 to intRegions
            strHTML = strHTML & "<div id='region" & intCounter & "' align='' left=''></div>"
        Next 
        strHTML = "" & strHTML & ""
 
        Set p_objMSIE = CreateObject("InternetExplorer.Application")
        With p_objMSIE
  
            .Navigate2 "about:blank"
            Do While .readyState <> 4
                Wscript.sleep 10
            Loop
            .Document.Title = strTitle
            .Document.Body.InnerHTML = strHTML
            .Document.Body.Scroll = "no"
            .Toolbar = False
            .StatusBar = False
            .Resizable = False
            .Width = intWidth
            .Height = intHeight
            .Left = 700
            .Top = 400
            .Visible = True
        End With
        Set CreateMSIEWindow = p_objMSIE
    End Function
 
    'This should be called to close the window
    Sub Destroy()
        objProgress.Quit
    End Sub
 
    'This method can be used to update the content of a specified region
    Sub UpdateRegion(p_intRegion, p_strContent)
        Set obj = objProgress.Document.GetElementByID("region" & Cstr(p_IntRegion))
        obj.innerHTML = p_strContent
    End Sub
 
    'This method clears a specified region.
    Sub ClearRegion(p_intRegion)
        UpdateRegion p_intRegion, ""
    End Sub
 
    'This method clears all of the regions at once.
    Sub ClearAllRegions()
        Dim intRegion
 
        For intRegion = 1 to intRegions
            UpdateRegion intRegion, ""
        Next
    End Sub
 
    'Set a progress bar as a region...
    'The first parameter is which region.
    'The second parameter is the percentage completion of the bar.
    'The third parameter is the number of characters that make up the bar.
    Sub UpdateRegionProgress(p_intRegion, p_intPrecentageProgress, p_intCharacterRange)
        'These constants are used to create the progress bar
        Const SOLID_BLOCK_CHARACTER = "Å°"
        Const EMPTY_BLOCK_CHARACTER = "Å†"
        Const SOLID_BLOCK_COLOUR = "#ffcc33;"
        Const EMPTY_BLOCK_COLOUR = "#666666;"
 
        Dim intSolidBlocks, intEmptyBlocks, intCounter
        Dim strProgress
 
        'Calculate how many blocks we need to create
        intSolidBlocks = round(p_intPrecentageProgress / 100 * p_intCharacterRange)
        intEmptyBlocks = p_intCharacterRange - intSolidBlocks
  
        intCounter = 0
        'Build the progress so far blocks
        strProgress = "<span style='font-family:courier; color:'" & SOLID_BLOCK_COLOUR & "'>" 
        While intCounter < intSolidBlocks
            strProgress = strProgress & SOLID_BLOCK_CHARACTER
            intCounter = intCounter + 1
        Wend
  
        strProgress = strProgress &  "</span><span style='font-family:courier;color:" & EMPTY_BLOCK_COLOUR & ";'>"
         
        'Build the progress to be met blocks
        strProgress = strProgress & ""
        intCounter = 0
        While intCounter < intEmptyBlocks
            strProgress = strProgress & EMPTY_BLOCK_CHARACTER
            intCounter = intCounter + 1
        Wend
        strProgress = strProgress &  "</span>"
  
        'Set the specified region to be the blocks
        UpdateRegion p_intRegion, strProgress
    End Sub
 
End Class