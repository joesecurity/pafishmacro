Private Sub Document_Open()

   On Error Resume Next

   ActiveDocument.Range.Text = "Pafish for Office Macro v2, by Joe Security" & vbCrLf & vbCrLf & "Includes latest in-the-wild evasion checks. Developed to test and improve sandboxes!" & vbCrLf
   
   checkRecentDocs
   
   checkNbrOfTask
   
   checkTasks
   
   checkZoneIdentifier
   
   checkPartOfDomain
   
   checkBios
   
   checkPnP
   
   checkUsername
   
   checkFilenameHash
   
   checkFilenameBad
   
   checkPreciseFileName
   
   checkCores
   
   checkAppCount
   
   checkApps
   
   mark
   
End Sub


Public Sub checkApps()

    printMsg "[*] WordBasic.AppGetNames ..."
    
    d = False
    tns = Array("vmware", "vmtools", "vbox", "process explorer", "processhacker", "procmon", "visual basic", "fiddler", "wireshark")
    Set ws = GetObject("winmgmts:\\.\root\cimv2")
    
    Dim names() As String
    ReDim names(WordBasic.AppCount())
    
    WordBasic.AppGetNames names
    
    For Each n In names
        For Each tn In tns
            If InStr(LCase(n), tn) > 0 Then
                d = True
            End If
        Next
    Next

    If d Then
    
        printMsg "DETECTED"
        
    Else
    
        printMsg "OK"
    End If
    
End Sub

Public Sub checkAppCount()

    printMsg "[*] Checking WordBasic.AppCount() ..."

    If WordBasic.AppCount() < 50 Then
    
        printMsg "DETECTED"
        
    Else
    
        printMsg "OK"
    End If
    
End Sub

Public Sub checkPreciseFileName()

    printMsg "[*] Checking Precise Filename ..."
    
    badName = False

  
    If ActiveDocument.Name <> "Pafish.docm" Then
            badName = True
    End If
 
    If badName Then
        
        printMsg "DETECTED"
    Else
        
        printMsg "OK"
    End If
    
End Sub

Public Sub checkFilenameHash()

    printMsg "[*] Checking Filename Hashname ..."
    
    hexchars = "0123456789abcdef"
    
    c = 0
    
    For i = 1 To Len(ThisDocument.Name)
        s = Mid(LCase(ThisDocument.Name), i, 1)
        
        If InStr(s, hexchars) > 0 Then
            c = c + 1
        End If
        
    Next
    
    If c >= (Len(ThisDocument.Name) - 5) Then
        printMsg "DETECTED"
        
    Else
    
    
        printMsg "OK"
    End If
    
End Sub

Public Sub checkFilenameBad()

    printMsg "[*] Checking Bad Filename ..."
    
    badName = False
    badNames = Array("malware", "myapp", "sample", ".bin", "mlwr_", "Desktop")

    
    For Each n In badNames
        If InStr(LCase(ActiveDocument.FullName), n) > 0 Then
            badName = True
        End If
    Next
 

    If badName Then
        
        printMsg "DETECTED"
    Else
        
        printMsg "OK"
    End If
    
End Sub

Public Sub checkTasks()

    printMsg "[*] Checking Application.Tasks.Name ..."

    badTask = False
    badTaskNames = Array("vbox", "vmware", "vxstream", "autoit", "vmtools", "tcpview", "wireshark", "process explorer", "visual basic", "fiddler")
    
    For Each Task In Application.Tasks
    
        For Each badTaskName In badTaskNames
            If InStr(LCase(Task.Name), badTaskName) > 0 Then
                badTask = True
            End If
        Next
        
    Next

    If badTask Then
        
        printMsg "DETECTED"
    Else
        
        printMsg "OK"
    End If
    
End Sub

Public Sub checkCores()

    printMsg "[*] Checking Win32_Processor.NumberOfCores ..."

    badCores = 0

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
    
    For Each objItem In colItems
    
            If objItem.NumberOfCores < 3 Then
                badCores = True
            End If
        
    Next

    If badCores Then
        
        printMsg "DETECTED"
    Else
        
        printMsg "OK"
    End If
    
End Sub

Public Sub checkBios()

    printMsg "[*] Checking Win32_Bios.SMBIOSBIOSVersion & SerialNumber ..."

    badBios = False
    badBiosNames = Array("virtualbox", "vmware", "kvm")
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Bios", , 48)
    
    For Each objItem In colItems
    
        For Each badName In badBiosNames
            If InStr(LCase(objItem.SMBIOSBIOSVersion), badName) > 0 Then
                badBios = True
            End If
            If InStr(LCase(objItem.SerialNumber), badName) > 0 Then
                badBios = True
            End If
        Next
        
    Next

    If badBios Then
        
        printMsg "DETECTED"
    Else
        
        printMsg "OK"
    End If
    
End Sub

Public Sub checkPnP()

    printMsg "[*] Checking Win32_PnPEntity.DeviceId ..."

    badPNP = False
    badPNPNames = Array("VEN_80EE", "VEN_15AD")
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity", , 48)
    
    For Each objItem In colItems
    
        For Each badName In badPNPNames
            If InStr(LCase(objItem.DeviceId), badName) > 0 Then
                badPNP = True
            End If
        Next
        
    Next

    If badPNP Then
        
        printMsg "DETECTED"
    Else
        
        printMsg "OK"
    End If
    
End Sub

Public Sub checkUsername()

    printMsg "[*] Checking Win32_ComputerSystem.Username ..."

    badUsername = False
    badUsernames = Array("admin", "malfind", "sandbox", "test")
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem", , 48)
    
    For Each objItem In colItems
    
        For Each badName In badUsernames
            If InStr(LCase(objItem.UserName), badName) > 0 Then
                badUsername = True
            End If
        Next
        
    Next

    If badUsername Then
        
        printMsg "DETECTED"
    Else
        
        printMsg "OK"
    End If
    
End Sub

Public Sub checkPartOfDomain()

    printMsg "[*] Checking Win32_ComputerSystem.PartOfDomain ..."

    partOfDomain = False
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem", , 48)
    
    For Each objItem In colItems
        If objItem.partOfDomain Then
            partOfDomain = True
        End If
    Next

    If partOfDomain Then
        printMsg "OK"
        
    Else
        printMsg "DETECTED"
    End If
    
End Sub

Public Sub checkZoneIdentifier()

    printMsg "[*] Checking Zone.Identifier ..."

    If CreateObject("Scripting.FileSystemObject").FileExists(ThisDocument.Path & Application.PathSeparator & ThisDocument.Name & ":Zone.Identifier") Then
    
        printMsg "OK"
        
    Else
    
        printMsg "DETECTED"
    End If
    
End Sub

Public Sub checkNbrOfTask()

    printMsg "[*] Checking Application.Tasks.Count ..."

    If Application.Tasks.Count < 3 Then
    
        printMsg "DETECTED"
        
    Else
    
        printMsg "OK"
    End If
    
End Sub

Public Sub checkRecentDocs()

    printMsg "[*] Checking Application.RecentFiles.Count ..."

    If Application.RecentFiles.Count < 3 Then
    
        printMsg "DETECTED"
        
    Else
    
        printMsg "OK"
    End If
    
End Sub

Public Function printMsg(msg)

   ActiveDocument.Range.Text = ActiveDocument.Range.Text & msg
   
   Set objFSO = CreateObject("Scripting.FileSystemObject")
 
    outFile = "pafish.log"
    Set objFile = objFSO.CreateTextFile(outFile, True)
    objFile.Write ActiveDocument.Range.Text & msg
    objFile.Close
    
End Function

Public Sub mark()

  Text = ActiveDocument.Range.Text
 
    toks = Split(Text, vbCr)
    
    c = 0
    
    For Each tok In toks
        
        l = Len(tok)
        
        If tok = "OK" Then

            ActiveDocument.Range(c, c + l).Font.color = vbGreen
     
        End If
  
        If tok = "DETECTED" Then

            ActiveDocument.Range(c, c + l).Font.color = vbRed
     
        End If
        
        
        c = c + l + 1
    Next
    
    ActiveDocument.Range.ParagraphFormat.SpaceBefore = 0
    ActiveDocument.Range.ParagraphFormat.SpaceAfter = 0
    ActiveDocument.Range.Font.Size = 8
  
End Sub

