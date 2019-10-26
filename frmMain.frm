VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB Code Lines"
   ClientHeight    =   1716
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   6252
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1716
   ScaleWidth      =   6252
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFile 
      Height          =   288
      Left            =   1092
      TabIndex        =   3
      Top             =   840
      Width           =   4716
   End
   Begin VB.CommandButton cmdEllipses 
      Caption         =   "..."
      Height          =   288
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   264
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   348
      Left            =   3696
      TabIndex        =   1
      Top             =   1260
      Width           =   1020
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   348
      Left            =   4788
      TabIndex        =   0
      Top             =   1260
      Width           =   1020
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   420
      Top             =   1176
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "VBP"
      DialogTitle     =   "Select Project File"
      Filter          =   "VBP Project Files (*.vbp)|*.vbp|All files (*.*)|*.*"
      Flags           =   4
   End
   Begin VB.Label Label2 
      Caption         =   "Note! Pass project filename as command line argument to automate the procedure of line numbering."
      Height          =   432
      Left            =   168
      TabIndex        =   5
      Top             =   84
      Width           =   5976
   End
   Begin VB.Label Label1 
      Caption         =   "VBP file:"
      Height          =   264
      Left            =   168
      TabIndex        =   4
      Top             =   840
      Width           =   852
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEllipses_Click()
    On Error GoTo EH_Cancel
    comDlg.ShowOpen
    If comDlg.FileName <> "" Then
        txtFile = comDlg.FileName
    End If
EH_Cancel:
End Sub

Private Sub cmdProcess_Click()
    Dim lNumFiles       As Long
    Dim lNumLines       As Long
    
    Screen.MousePointer = vbHourglass
    lNumFiles = ProcessProject(txtFile, lNumLines)
    Screen.MousePointer = vbDefault
    MsgAlert "Successfully put " & lNumLines & " line numbers on " & lNumFiles & " files in " & txtFile
End Sub
