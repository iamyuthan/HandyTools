Public Sub live_SilentWait()

    Dim fso         As Object
    Dim sh          As Object
    Dim ts          As Object
    Dim basePath    As String
    Dim folderPath  As String
    Dim filePath    As String
    Dim counter     As Long
    Dim resp        As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh  = CreateObject("WScript.Shell")

    ' --- Base path: C:\Users\<YourName>\Documents\ ---
    basePath = Environ("USERPROFILE") & "\Documents\"
    counter = 1

    Do
        ' ========================================
        '  STEP 1 : Create Folder & File
        ' ========================================
        folderPath = basePath & "Test_" & counter
        filePath = folderPath & "\Text_" & counter & ".txt"
        
        fso.CreateFolder folderPath
        Set ts = fso.CreateTextFile(filePath, True)
        ts.WriteLine "Created at : " & Format(Now, "dd-mmm-yyyy hh:nn:ss AM/PM")
        ts.Close
        Set ts = Nothing

        ' ========================================
        '  STEP 2 : 10-SECOND Notification (Creation)
        '           (Click OK to stop script)
        ' ========================================
        resp = sh.Popup( _
                "Cycle #" & counter & " — Files created." & vbCrLf & _
                "Waiting 1 min before deletion...", _
                10, _
                "live  |  Running", _
                vbOKOnly + vbInformation)
        
        ' ========================================
        '  STEP 3 : WAIT 1 MINUTE (Silently)
        ' ========================================
        Application.Wait Now + TimeValue("00:01:00")
        DoEvents

        ' ========================================
        '  STEP 4 : Delete File & Folder
        ' ========================================
        fso.DeleteFile filePath
        fso.DeleteFolder folderPath
        
        ' ========================================
        '  STEP 5 : 10-SECOND Notification (Deletion)
        '           (Click OK to stop script)
        ' ========================================
        resp = sh.Popup( _
                "Cycle #" & counter & " — Cleanup done." & vbCrLf & _
                "Waiting 2 min before next cycle...", _
                10, _
                "live  |  Running", _
                vbOKOnly + vbInformation)

        If resp = vbOK Then Exit Do
        
        ' ========================================
        '  STEP 6 : WAIT 2 MINUTES (Silently)
        ' ========================================
        Application.Wait Now + TimeValue("00:02:00")
        DoEvents

        ' ========================================
        '  STEP 7 : Increment counter & Repeat
        ' ========================================
        counter = counter + 1

    Loop

    ' --- Cleanup objects ---
    Set ts  = Nothing
    Set fso = Nothing
    Set sh  = Nothing

    MsgBox "Script stopped at Cycle #" & counter & ".", _
           vbInformation, "live  |  Stopped"

End Sub
