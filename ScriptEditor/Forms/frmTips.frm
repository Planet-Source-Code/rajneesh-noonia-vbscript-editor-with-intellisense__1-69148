VERSION 5.00
Begin VB.Form frmTips 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   315
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3480
      Top             =   60
   End
   Begin VbScriptEditor.FunctionTips FunctionTips1 
      Height          =   270
      Left            =   0
      Top             =   0
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   476
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Private Const GWL_STYLE As Long = (-16&)
  Private Const GWL_EXSTYLE As Long = (-20&)
  Private Const WS_THICKFRAME As Long = &H40000
  Private Const WS_MINIMIZEBOX As Long = &H20000
  Private Const WS_MAXIMIZEBOX As Long = &H10000
  Private Const SW_SHOWNOACTIVATE = 4
  Private Const WS_EX_TOOLWINDOW As Long = &H80&
  Private Const GWL_HINSTANCE As Long = -6

  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd&, ByVal nIndex&) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
  Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
        

  ' SetWindowPos Flags
  Private Const SWP_NOSIZE = &H1
  Private Const SWP_NOMOVE = &H2
  Private Const SWP_NOZORDER = &H4
  Private Const SWP_NOREDRAW = &H8
  Private Const SWP_NOACTIVATE = &H10
  Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
  Private Const SWP_SHOWWINDOW = &H40
  Private Const SWP_HIDEWINDOW = &H80
  Private Const SWP_NOCOPYBITS = &H100
  Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
  
  Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
  Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
  
  ' SetWindowPos() hwndInsertAfter values
  Private Const HWND_TOP = 0
  Private Const HWND_BOTTOM = 1
  Private Const HWND_TOPMOST = -1
  Private Const HWND_NOTOPMOST = -2
  Public m_ParentCtrl As Editor
  
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Property Get TipText() As String
    TipText = FunctionTips1.TipText
End Property

Public Property Let TipText(ByVal pTipText As String)
    FunctionTips1.TipText = pTipText
End Property

Public Property Get TextFont()
    Set TextFont = FunctionTips1.Font
End Property

Public Property Set TextFont(ByVal pTextFont As Font)
    Set FunctionTips1.Font = pTextFont
End Property


Public Sub ShowTip(ByVal pBoolean As Boolean)
    If pBoolean Then
        'Me.Visible = True
        If Len(Trim(FunctionTips1.TipText)) > 0 Then
            SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
                                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
            ShowWindow Me.hwnd, SW_SHOWNOACTIVATE
            Timer1.Enabled = True
        End If
    Else
        Timer1.Enabled = False
        Me.Visible = False
    End If
End Sub

Private Sub Form_Deactivate()
    Me.ShowTip False
End Sub

Private Sub Form_LostFocus()
    Me.ShowTip False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ShowTip False
End Sub

Private Sub Form_Resize()
    FunctionTips1.Left = 0
    FunctionTips1.Top = 0
    With Me
        .Height = FunctionTips1.Height
        .Width = FunctionTips1.Width
    End With
    
End Sub

Private Sub FunctionTips1_Resize()
    Call Form_Resize
End Sub

Private Sub Timer1_Timer()
    If Not m_ParentCtrl Is Nothing Then
        m_ParentCtrl.PositionTip
    End If
End Sub
