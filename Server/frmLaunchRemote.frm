VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLaunchRemote 
   Caption         =   "Run Remote Application"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtTriggerFolder 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtMachineName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtApplication 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTriggerFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Trigger Folder:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblMachine 
      BackStyle       =   0  'Transparent
      Caption         =   "Machine Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblLocalFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Local EXE"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmLaunchRemote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    With CommonDialog1
        .DialogTitle = "Select an Application"
        .Filter = "*.exe"
        .FileName = "*.exe"
        .ShowOpen
        
        If .FileName > "*.exe" Then
            txtApplication.Text = .FileName
        End If
        
    End With
       
    
End Sub

Private Sub cmdSend_Click()

Dim strDestination As String

    strDestination = "\\" & txtMachineName.Text & "\" & txtTriggerFolder.Text & "\"
    If Dir(txtApplication.Text) = "" Then
        MsgBox "The File " & txtApplication.Text & " does not exist.", vbExclamation
    ElseIf Dir(strDestination, vbDirectory) = "" Then
        MsgBox "The selected trigger folder does not exist.", vbExclamation
    Else
        FileCopy txtApplication.Text, strDestination & "run.exe"
        MsgBox "Command to run " & txtApplication.Text & " was sent.", vbInformation
    End If
End Sub
