VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmThread 
   Caption         =   "UserForm1"
   ClientHeight    =   345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "frmThread.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    Me.Hide
    
    Dim t As Single
    Dim count As Integer
    Dim running As Boolean
    
    t = Timer
    count = 0
    running = True
    
    While running
        If Timer > t + 1# Then
            Debug.Print "Tick! " & count & " " & Timer
            'Application.ActiveExplorer.
            t = Timer
            count = count + 1
        End If
        
        If count > 10 Then
            running = False
        End If
        
        DoEvents
    Wend
    
    Unload Me
End Sub
