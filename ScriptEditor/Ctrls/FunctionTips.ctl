VERSION 5.00
Begin VB.UserControl FunctionTips 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   FillColor       =   &H80000018&
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000013&
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   300
   ScaleWidth      =   4140
   Begin VB.Label lblFunctionReturn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   ") As Object"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1890
      TabIndex        =   2
      Top             =   0
      Width           =   870
   End
   Begin VB.Label lblParam 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Parameter1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblFunctionName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Function("
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   765
   End
End
Attribute VB_Name = "FunctionTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TipText As String
Private m_boldIndex As Long
Public Event Resize()

Public Property Get BoldIndex() As Long
    BoldIndex = m_boldIndex
End Property

Public Property Let BoldIndex(ByVal pBoldIndex As Long)
    m_boldIndex = pBoldIndex
    Call FormatTip
End Property

Public Property Get TipText() As String
    TipText = m_TipText
End Property

Public Property Let TipText(ByVal pTipText As String)
    m_TipText = pTipText
    m_boldIndex = 0
    Call FormatTip
End Property

Public Property Get Font()
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal pFont As Font)
    Set UserControl.Font = pFont
    Call UserControl_AmbientChanged("Font")
End Property

Private Sub FormatTip()
    Dim pParamText As String
    Dim pParams() As String
    Dim pParamCount As Long
    Dim pPos As Long
    Dim pFunctionName As String
    Dim pFunctionEnd As String
    Dim pCounter As Long
    Dim pLeft As Long
    
    lblFunctionName.Visible = False
    For pCounter = lblParam.Count - 1 To 0 Step -1
        If pCounter > 0 Then
            lblParam(pCounter).Visible = False
            Unload lblParam(pCounter)
        Else
            lblParam(pCounter).Visible = False
        End If
    Next
    lblFunctionReturn.Visible = False
    
    
    If Len(m_TipText) = 0 Then Exit Sub
    pPos = InStr(1, m_TipText, "(")
    
    pFunctionName = Left(m_TipText, pPos)
    pParamText = Mid(m_TipText, pPos + 1)
    pPos = InStr(1, StrReverse(pParamText), ")")
    
    If pPos > 0 Then
        pPos = Len(pParamText) - pPos
        pFunctionEnd = Mid(pParamText, pPos + 1)
    End If
    
    pParamText = Left(pParamText, pPos)
    pParamText = Replace(pParamText, ",", ",`")
    pParams = Split(pParamText, "`")
    
    
    lblFunctionName.Caption = pFunctionName
    lblFunctionName.Visible = True
    pLeft = lblFunctionName.Width + lblFunctionName.Left
    For pCounter = 0 To UBound(pParams)
        
        If pCounter > 0 Then
            Load lblParam(pCounter)
            lblParam(pCounter).Left = pLeft
            lblParam(pCounter).Caption = pParams(pCounter)
            lblParam(pCounter).Visible = True
            
        Else
            lblParam(pCounter).Caption = pParams(pCounter)
            lblParam(pCounter).Left = pLeft
            lblParam(pCounter).Visible = True
            
        End If
        lblParam(pCounter).FontBold = False
        
        If pCounter = m_boldIndex Then
            lblParam(pCounter).FontBold = True
        End If
        pLeft = lblParam(pCounter).Width + pLeft
    Next
    
    lblFunctionReturn.Left = pLeft
    lblFunctionReturn.Caption = pFunctionEnd
    lblFunctionReturn.Visible = True
    Call UserControl_Resize
    
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Dim pWidth As Long
    Dim pCounter As Single

    If PropertyName = "Font" Then
        
        Set lblFunctionName.Font = UserControl.Font
        For pCounter = 0 To lblParam.Count - 1
            Set lblParam.Item(pCounter).Font = UserControl.Font
        Next
        Set lblFunctionReturn.Font = UserControl.Font
    End If
End Sub

Private Sub UserControl_Resize()
    Dim pWidth As Long
    Dim pCounter As Single
    
    
    pWidth = lblFunctionReturn.Left + lblFunctionReturn.Width + 100
    With UserControl
        .Height = lblFunctionReturn.Height + 50
        .Width = pWidth
    End With
    RaiseEvent Resize
End Sub
