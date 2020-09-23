VERSION 5.00
Begin VB.Form frmListen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listening..."
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "frmListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strListenFolder As String
Dim strExeFolder As String
Dim iShow As VbAppWinStyle

Private Sub Form_Load()
    
    strListenFolder = "c:\temp\share\"
    strExeFolder = "c:\temp\exe\"
    '   You can change this to vbHide hide the "kicked off" application.
    iShow = vbMaximizedFocus
End Sub

Private Sub Timer1_Timer()
Dim strFileName As String

    '   Check the shared folder for a new .exe file
    strFileName = Dir(strListenFolder & "*.exe")
        
        
    If strFileName > "" Then
        Me.Caption = "Executing " & strFileName & "..."
           '    Move the file so it won't kick off each time the timer fires
        FileCopy strListenFolder & strFileName, "c:\temp\exe\" & strFileName
        Kill strListenFolder & strFileName
        
        '   Run the application
        Shell strExeFolder & strFileName, iShow
        
        'comment the next line out for testing
        'Msgbox "File was executed"
        
        Me.Caption = "Listening..."
    
    End If
    
End Sub
