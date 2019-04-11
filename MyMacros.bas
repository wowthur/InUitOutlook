Attribute VB_Name = "MyMacros"
Public Sub AddInOut()
    frmInOut.Show
End Sub

Public Sub AddAppointment(ByVal subject As String, ByVal appTime As Date)
    Dim klokFolder As Folder
    
    Set klokFolder = GetKlokFolder
    
    'Debug.Print "Found folder " & klokFolder.FolderPath
    
    Dim app As AppointmentItem
    
    Set app = klokFolder.Items.Add
    
    With app
        .subject = subject
        .Start = appTime
        .End = appTime
        .Categories = "Categorie Groen"
        .BusyStatus = olFree
    End With
    
    If subject = "Uit" Then
        app.ReminderMinutesBeforeStart = 5
    Else
        app.ReminderSet = False
    End If
    
    app.Save
End Sub

Public Function GetKlokFolder() As Folder
    Dim f As Folder
    
    For Each f In Application.Session.Folders
        'Debug.Print f.FolderPath
        
        If Right(f.Name, 15) = "@dkgservices.nl" Then
            For Each f2 In f.Folders
                'Debug.Print f2.FolderPath
                
                If f2.Name = "Klok" Then
                    Set GetKlokFolder = f2
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Public Sub TestMacro()
    'frmThread.Show vbModeless
    Dim pane As OutlookBarPane
    Dim grp As OutlookBarGroup
    Set pane = Application.ActiveExplorer.Panes.Item("OutlookBar")
    For Each grp In pane.Contents.Groups
        Debug.Print grp.Name
    Next
End Sub
