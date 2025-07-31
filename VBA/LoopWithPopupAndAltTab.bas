Option Explicit

Public Sub LoopWithPopupAndAltTab()
    Dim sh   As Object
    Dim resp As Long
    
    Set sh = CreateObject("WScript.Shell")
    
    Do
        ' 10-second popup (OK only)
        resp = sh.Popup( _
                 "Script running… click OK to stop.", _
                 10, _
                 "Script Running", _
                 vbOKOnly + vbInformation _
               )
               
        ' if user clicked OK → exit
        If resp = vbOK Then Exit Sub
        
        ' -- HERE'S THE ONLY CHANGE --
        ' make sure the popup is gone before we send Alt+Tab
        DoEvents
        
        ' now switch windows
        Application.SendKeys "%{TAB}", True
    Loop
End Sub
