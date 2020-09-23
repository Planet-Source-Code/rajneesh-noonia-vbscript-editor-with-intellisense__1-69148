VERSION 5.00
Object = "*\A..\VbScriptEditor.vbp"
Begin VB.Form frmMain 
   Caption         =   "Wrapper"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VbScriptEditor.Editor Editor1 
      Height          =   7125
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   12568
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Script As String


Private Sub cmdClear_Click()
    If Len(Editor1.Script) > 0 Then
        m_Script = Editor1.Script
        Editor1.Script = ""
    End If
    
End Sub

Private Sub cmdLoad_Click()
     
End Sub

Private Sub Command1_Click()
    Dim pXML As String
    
    
    MsgBox Editor1.XMLProject
End Sub


Private Sub Form_Resize()
    
    On Error Resume Next
    With Editor1
       ' .Left = 0
       ' .Top = 0
       ' .Width = Me.Width - 100
       ' .Height = Me.Height - 555
    End With
End Sub
