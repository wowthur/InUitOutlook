VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInOut 
   Caption         =   "In Uit"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2070
   OleObjectBlob   =   "frmInOut.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private InDate As Date
Private OutDate As Date

Private Sub btnOk_Click()
    MyMacros.AddAppointment "In", InDate
    MyMacros.AddAppointment "Uit", OutDate
    
    Unload Me
End Sub

Private Sub tbIn_Change()
    Dim dt As Date
    
    If Len(tbIn.Text) <> 5 Then
        Exit Sub
    End If
    
    Dim nowDate As String
    
    nowDate = DatePart("yyyy", Now) & "-" & DatePart("m", Now) & "-" & DatePart("d", Now)
    
    InDate = CDate(nowDate & " " & tbIn.Text)
    
    OutDate = DateAdd("h", 8, InDate)
    OutDate = DateAdd("n", 15, OutDate)
    
    tbOut.Text = FormatDateTime(OutDate, vbShortTime)
End Sub
