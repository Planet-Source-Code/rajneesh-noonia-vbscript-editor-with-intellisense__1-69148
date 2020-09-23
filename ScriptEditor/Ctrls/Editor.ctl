VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl Editor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   6660
   ScaleWidth      =   9570
   Begin MSComctlLib.ListView lvwVariables 
      Height          =   795
      Left            =   7260
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1402
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilsResources"
      SmallIcons      =   "ilsResources"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "X"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwVBA 
      Height          =   675
      Left            =   7320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1191
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilsResources"
      SmallIcons      =   "ilsResources"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwAllClasses 
      Height          =   585
      Left            =   7320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1032
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilsResources"
      SmallIcons      =   "ilsResources"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsResources 
      Left            =   1620
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   13642488
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":0000
            Key             =   "UNKNOWN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":05F2
            Key             =   "CLASS"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":0BE4
            Key             =   "CONSTANT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":11D6
            Key             =   "ENUM"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":17C8
            Key             =   "EVENT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":1DBA
            Key             =   "FUNCTION"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":23AC
            Key             =   "LIBRARY"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":299E
            Key             =   "MODULE"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":2F90
            Key             =   "PROPERTY"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":3582
            Key             =   "TYPE"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":3B74
            Key             =   "VAR"
         EndProperty
      EndProperty
   End
   Begin VB.Timer LoadInBack 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   660
      Top             =   4200
   End
   Begin MSComctlLib.ListView lvwCreatables 
      Height          =   585
      Left            =   7350
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1032
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilsResources"
      SmallIcons      =   "ilsResources"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtEditor 
      Height          =   3090
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   5450
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Editor.ctx":3F86
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   6345
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2752
            MinWidth        =   1764
            Picture         =   "Editor.ctx":4006
            Text            =   "Refrences"
            TextSave        =   "Refrences"
            Key             =   "REF"
            Object.ToolTipText     =   "Add refrences for intellisense"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11483
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   6000
      Y1              =   6000
      Y2              =   6000
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Private IsLoaded As Boolean
Private isDirty As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long

Private WithEvents m_ActiveListViewControl As ListView
Attribute m_ActiveListViewControl.VB_VarHelpID = -1
Private m_TipPos As POINTAPI

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SCROLL = (LVM_FIRST + 20)
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Private m_OperatorType As OperatorType
'wMSG For Find Line Position
Const EM_LINEINDEX = &HBB
Const WM_SETREDRAW = &HB

Dim bKey As Boolean
' True If The RTF Is Change
Dim bChange As Boolean
' Last Line Of RTF
Dim LastLine As Integer

' Color
Dim K_COLOR(1 To 2) As Long
Dim C_COLOR As Long
Dim Q_COlOR As Long
Dim N_Color As Long
Private ctrlTips As frmTips
Dim strDelimiter As String
Dim Delimiter(27) As String
Private m_LoadingScript As Boolean
Private m_Intellisense As cIntellisense
Dim m_PartialText As String
Dim m_PartialIniLen As Long
Dim m_LastSearchedItem As ListItem
Dim LastStart As Long
Private m_blnIgnoreKey As Boolean
Private m_XMLProject As String
Private m_frmRefrences As frmRefrences
Private m_PreviousLine As String
Private Const GLOBAL_MODULE = "GLOBAL_Module"
Private m_ReturnKeyPressed As Boolean
Private m_TabsInLastLine As String

Public Property Let XMLProject(ByVal pXMLProject As String)
    g_Project.XML = pXMLProject
End Property

Public Property Get XMLProject() As String
    XMLProject = g_Project.XML
End Property

Private Sub LoadRefrences()
    m_Intellisense.IsBusyLoading = True
    Dim pRefrence As Refrence
    SB1.Panels(2).Text = "Loading Class Info..."
    For Each pRefrence In g_Project.Refrences
        Call m_Intellisense.AddRefToIntellsense(pRefrence)
    Next
    Call PopulateVBAFunctions
    Call PopulateCreatableClasses
    Call PopulateAllClasses
    m_Intellisense.IsBusyLoading = False
    SB1.Panels(2).Text = ""
End Sub

Public Sub AddIntelRefrenceFromObject(ByVal pObject As Object)
    Dim pRefrence As Refrence
    On Error Resume Next
    Set pRefrence = m_Intellisense.GetRefrenceFromObject(pObject)
    g_Project.Refrences.Add pRefrence.GUID, pRefrence.Verion, pRefrence.BinaryPath, pRefrence.Name
    Call LoadRefrences
    
End Sub


Public Sub InsertText(ByVal pText As String)
    UserControl.rtEditor.SelText = pText
End Sub

Private Sub LoadInBack_Timer()
    Dim pRefrence As Refrence
    Dim pRefExisting As Refrence
    LoadInBack.Enabled = False
    
    Set pRefrence = m_Intellisense.GetVBARefrence()
    Call m_Intellisense.AddRefToIntellsense(pRefrence)
    On Error Resume Next
    Set pRefExisting = g_Project.Refrences.Item(pRefrence.GUID & "#" & pRefrence.Verion)
    If pRefExisting Is Nothing Then
        Call g_Project.Refrences.Add(pRefrence.GUID, pRefrence.Verion, pRefrence.BinaryPath, pRefrence.Name)
    End If
    On Error GoTo 0
    Call LoadRefrences
    
    m_Intellisense.IsBusyLoading = True
    LoadInBack.Enabled = False
    'lblWait.Caption = "Please wait ..."
    'lblWait.Refresh
   
    m_Intellisense.IsBusyLoading = False
    rtEditor.Enabled = True
    SB1.Enabled = True
    'UserControl.MousePointer = vbNormal
    'lblWait.Caption = "Done"
    'fraMessage.Visible = False
    UserControl.Enabled = True
    'rtEditor.SetFocus
    Call rtEditor_SelChange
    Call ParseForModulesAndVars
End Sub


Private Sub InitTabTrap(ByVal bState As Boolean)
   'Dim pFont As Font
   If bState Then
      
      If UserControl.Ambient.UserMode Then
        Set ctrlTips = New frmTips
        Set ctrlTips.m_ParentCtrl = Me
        Call SetParent(lvwAllClasses.hwnd, 0) 'UserControl.Parent.hwnd)
        Call SetParent(lvwCreatables.hwnd, 0) 'UserControl.Parent.hwnd)
        Call SetParent(lvwVariables.hwnd, 0) 'UserControl.Parent.hwnd)
        Call SetParent(lvwVBA.hwnd, 0)
        'Set pFont = rtEditor.Font
    '    pFont.Size = pFont.Size - 1
     '   Set ctrlTips.TextFont = pFont
        ctrlTips.ShowTip False
        LoadInBack.Enabled = True
      End If
   End If
End Sub


Private Sub lvwAllClasses_LostFocus()
    If m_ActiveListViewControl Is Nothing Then
        lvwAllClasses.Visible = False
    End If
End Sub

Private Sub lvwCreatables_LostFocus()
    If m_ActiveListViewControl Is Nothing Then
        lvwCreatables.Visible = False
    End If
End Sub

Private Sub lvwVariables_LostFocus()
    If m_ActiveListViewControl Is Nothing Then
        lvwVariables.Visible = False
    End If
End Sub

Private Sub lvwVBA_LostFocus()
    If m_ActiveListViewControl Is Nothing Then
        lvwVBA.Visible = False
    End If
End Sub

Private Sub m_ActiveListViewControl_DblClick()
    Call HideListBox
End Sub

Private Sub m_ActiveListViewControl_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set m_LastSearchedItem = Item
End Sub

Private Sub m_ActiveListViewControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyEscape Then
            m_ActiveListViewControl.Visible = False
            Set m_ActiveListViewControl = Nothing
        Else
            Call HideListBox
        End If
    End If
End Sub


Private Sub rtEditor_Change()
    On Error GoTo ErrorTrap
    If m_LoadingScript Then Exit Sub
    isDirty = True
    bChange = True

    ' Update Color
    Dim OStart As Long
    Dim OLen As Long
    Dim pLenDelta As Long
    Dim StartPos As Long
    Dim EndPos As Long

    Dim EndLine As Integer
    Dim StartLine As Integer
    Dim X As Long
    Dim pTextBefore As String
    Dim Text As String
    Dim pUnformatedText As String
    Dim pOnlyCreatableClasses As Boolean
    Dim pItemTextToSearch As String
    Dim pPos As Long
    Dim pModName As String
    Dim pOperatorType As OperatorType
    Dim pShowAfterEqual As Boolean
    
    
    
    If Not m_ActiveListViewControl Is Nothing Then
        m_ActiveListViewControl.ZOrder 0
    End If

    With rtEditor
    'If .Text = "" Then Exit Sub

    'x = SendMessage(.hwnd, WM_SETREDRAW, 0, 0)
    Call LockWindowUpdate(.hwnd)
    If LastStart > .SelStart Then
        EndLine = .GetLineFromChar(LastStart)
        StartLine = .GetLineFromChar(.SelStart)
    Else
        StartLine = .GetLineFromChar(LastStart)
        EndLine = .GetLineFromChar(.SelStart)
    End If

    StartPos = SendMessage(.hwnd, EM_LINEINDEX, StartLine, 0&)
    EndPos = SendMessage(.hwnd, EM_LINEINDEX, EndLine + 1, 0&)

    If EndPos <= 0 Then EndPos = Len(.Text)

    OStart = .SelStart
    OLen = .SelLength

    .SelStart = StartPos
    .SelLength = EndPos - StartPos
     Text = .SelText
     pUnformatedText = Text
     If Len(Text) > 0 Then
      
      'If Right(Text, 2) = vbCrLf Then
      '  If Left(Right(Text, 4), 2) <> vbCrLf Then
      '      Text = Replace(Left(Text, Len(Text) - 2), vbCrLf, "") & Right(Text, 2)
      '  End If
      'End If
      
      Dim pEqualFound As Single
      pEqualFound = 1

        If InStr(pEqualFound, Text, "=") > 0 And m_ReturnKeyPressed Then
            Do
                 If InStr(pEqualFound, Text, " =") > 0 Then
                    pEqualFound = InStr(pEqualFound, Text, " =")
                    If Left(Mid(Text, pEqualFound + 2), 1) <> " " Then
                        Text = Left(Text, pEqualFound + 1) & " " & Mid(Text, pEqualFound + 2)
                        pLenDelta = pLenDelta + 1

                    End If
                ElseIf InStr(pEqualFound, Text, "= ") > 0 Then
                    pEqualFound = InStr(pEqualFound, Text, "= ")
                    If Left(Mid(Text, pEqualFound - 1), 1) <> " " Then
                        Text = Left(Text, pEqualFound - 1) & " =" & Mid(Text, pEqualFound + 1)
                        pLenDelta = pLenDelta + 1
                    End If

                ElseIf InStr(pEqualFound, Text, " = ") > 0 Then
                '
                Else
                    pEqualFound = InStr(pEqualFound, Text, "=")
                    If pEqualFound > 1 Then
                        Text = Left(Text, pEqualFound - 1) & " = " & Mid(Text, pEqualFound + 1)
                        pLenDelta = pLenDelta + 2
                    End If
                End If
                pEqualFound = InStr(pEqualFound, Text, "=") + 1

           Loop While InStr(pEqualFound, Text, "=") > 0
      End If

        .SelRTF = ColorIt(Text, pLenDelta)
    End If
        .SelStart = OStart + pLenDelta
        .SelLength = OLen

     LockWindowUpdate (0)
    LastStart = .SelStart
    End With

    If m_blnIgnoreKey Then
        If m_ReturnKeyPressed Then
            rtEditor.SelText = m_TabsInLastLine
            Text = Mid(rtEditor.Text, Abs(rtEditor.SelStart - 16) + 1)
            Text = Replace(Text, vbTab, " ")
            Text = Replace(Text, vbCrLf, "")
            pPos = InStr(1, Text, "Sub ")
            If pPos > 0 Then
                    If InStr(pPos, Text, "End Sub") <= 0 Then
                        pPos = InStr(pPos + 3, Text, "Sub ")
                        If pPos <= 0 Then
                            rtEditor.SelText = ColorIt("End Sub", 0)
                        End If
                    End If
            ElseIf InStr(1, Text, "Function ") > 0 Then
                pPos = InStr(1, Text, "Function ")
                If InStr(pPos, Text, "End Function") <= 0 Then
                    pPos = InStr(pPos + 8, Text, "Function ")
                    If pPos <= 0 Then
                        rtEditor.SelText = ColorIt("End Function", 0)
                    End If
                End If
            End If
        End If
        m_blnIgnoreKey = False
        Exit Sub
    End If

    
    pItemTextToSearch = Replace(pUnformatedText, vbCrLf, "")
    pItemTextToSearch = Trim(Replace(pItemTextToSearch, vbTab, " "))
    If Len(pItemTextToSearch) > 0 Then
        If Left(pItemTextToSearch, 1) = "." Then
            pPos = InStr(Len(rtEditor.Text) - OStart, StrReverse(rtEditor.Text), "htiW")
            If pPos > 0 Then
                pPos = Len(rtEditor.Text) - pPos - 4
                pItemTextToSearch = Mid(rtEditor.Text, pPos)
                pItemTextToSearch = Replace(pItemTextToSearch, vbCrLf, " ")
                pItemTextToSearch = Trim(Replace(pItemTextToSearch, vbTab, " "))
                pPos = InStr(1, pItemTextToSearch, "With")
                If pPos > 0 Then
                    pItemTextToSearch = Trim(Mid(pItemTextToSearch, pPos + 4))
                End If
                pPos = InStr(1, pItemTextToSearch, " ")
                If pPos > 0 Then
                    pItemTextToSearch = Trim(Left(pItemTextToSearch, pPos))
                    pUnformatedText = pItemTextToSearch & Replace(pUnformatedText, vbTab, "")
                End If
            End If
        End If
    End If
    If ctrlTips.Visible Then
        ctrlTips.ShowTip False
    End If
        
    pOperatorType = m_Intellisense.ParseOperator(rtEditor, pUnformatedText, m_PartialText)
    If m_OperatorType <> pOperatorType Then
        If Not m_ActiveListViewControl Is Nothing Then
            If m_ActiveListViewControl.Visible Then
                m_ActiveListViewControl.Visible = False
            End If
            Set m_LastSearchedItem = Nothing
            Set m_ActiveListViewControl = Nothing
            
        End If
    End If
    m_OperatorType = pOperatorType
    If (pOperatorType = otDimAs Or pOperatorType = otCreateObject) And m_ActiveListViewControl Is Nothing Then
            If m_ActiveListViewControl Is Nothing Then
                m_PartialIniLen = Len(Text)
                'm_PartialText = ""
                Set m_LastSearchedItem = Nothing
            End If

            If m_ActiveListViewControl Is Nothing Then
                pOnlyCreatableClasses = False
                If pOperatorType = otCreateObject Then
                    pOnlyCreatableClasses = True
                End If
                If pOnlyCreatableClasses Then
                    Set m_ActiveListViewControl = lvwCreatables
                Else
                    Set m_ActiveListViewControl = lvwAllClasses
                End If
            End If
        If Not m_ActiveListViewControl.Visible Then
            Call ShowListView
            Call rtEditor_SelChange
        End If
    ElseIf (pOperatorType = otDot Or pOperatorType = otEqualTo Or pOperatorType = otShowTip) And m_ActiveListViewControl Is Nothing Then
        
        If pOperatorType = otEqualTo Then
            pShowAfterEqual = True
        End If
        If m_ActiveListViewControl Is Nothing Then
            Set m_ActiveListViewControl = lvwVariables
        Else
            If m_ActiveListViewControl.hwnd <> lvwVariables.hwnd Then
                Set m_ActiveListViewControl = lvwVariables
            End If
        End If
        pItemTextToSearch = Text
        
        Dim pTempText As String
        Dim pShowTip As Boolean
        Dim Y As Long
        
        Y = GetCurrentColumn(rtEditor)
        pItemTextToSearch = Replace(pItemTextToSearch, vbTab, "     ")
        If Y - 2 > 1 Then
            pItemTextToSearch = Left(pItemTextToSearch, Y - 2)
        End If
        
        pItemTextToSearch = Replace(pItemTextToSearch, vbCrLf, "")
        pItemTextToSearch = Trim(pItemTextToSearch)
        Y = Len(pItemTextToSearch) - 1
        If pShowAfterEqual Then
            pPos = InStr(1, StrReverse(pItemTextToSearch), "=")
            If pPos > 0 Then
                pPos = Len(pItemTextToSearch) - pPos
                pItemTextToSearch = Trim(Left(Text, pPos))
            End If
            
                
            
            pPos = InStr(1, StrReverse(pItemTextToSearch), " ")
            If pPos > 0 Then
                pPos = Len(pItemTextToSearch) - pPos
                pItemTextToSearch = Trim(Mid(pItemTextToSearch, pPos + 1))
            End If
            
        Else
            If m_OperatorType = otShowTip Then pShowTip = True
            If Y = 0 Then Y = 1
            pPos = InStr(Y, pItemTextToSearch, "(")
            If pPos > 0 Then
                pShowTip = True
                pTempText = Left(pItemTextToSearch, pPos - 1)
                pPos = InStr(pPos, pItemTextToSearch, ")")
                
                If pPos > 0 Then
                    pShowTip = False
                    pTempText = pTempText & Mid(pItemTextToSearch, pPos + 1)
                End If
                pItemTextToSearch = pTempText
            End If
            pPos = InStr(1, StrReverse(pItemTextToSearch), ".")
            If pPos > 0 Then
                pPos = Len(pItemTextToSearch) - pPos
            End If
            m_PartialText = Trim(Mid(pItemTextToSearch, pPos + 2))
            'pItemTextToSearch = Trim(Left(Text, pPos))
            'pPos = InStr(1, StrReverse(pItemTextToSearch), " ")
            'If pPos > 0 Then
            '    pPos = Len(pItemTextToSearch) - pPos
                
            'End If
            'pItemTextToSearch = Trim(Mid(pItemTextToSearch, pPos + 1))
        End If
        pModName = FindCurrentModuleName(EndPos)
        
        pItemTextToSearch = Replace(pItemTextToSearch, vbTab, " ")
        If Left(Trim(pItemTextToSearch), 1) = "." Then
            pItemTextToSearch = Replace(pUnformatedText, "=", "")
            pItemTextToSearch = Replace(pItemTextToSearch, vbCrLf, "")
            pItemTextToSearch = Replace(pItemTextToSearch, vbTab, " ")
        Else
        End If
        
        Dim pFinalText As String
        Dim pPos2 As Long
        
        Do
            pPos = InStr(1, pItemTextToSearch, "(")
            If pPos > 0 Then
                pPos2 = InStr(pPos, pItemTextToSearch, ")")
            Else
                pPos2 = 0
            End If
            
            If pPos2 > 0 And pPos > 0 Then
                pFinalText = Left(pItemTextToSearch, pPos - 1)
                pFinalText = pFinalText & Mid(pItemTextToSearch, pPos2 + 1)
            Else
                If pPos > 0 Then
                    pFinalText = Left(pItemTextToSearch, pPos - 1)
                Else
                    If pPos2 > 0 Then
                        pFinalText = Left(pItemTextToSearch, pPos - 1)
                    Else
                        pFinalText = pItemTextToSearch
                    End If
                End If
            End If
            pItemTextToSearch = pFinalText
            
        Loop While InStr(1, pItemTextToSearch, "(") > 0 Or InStr(1, pItemTextToSearch, ")") > 0
        
         Call PopulateVariableProprs(pShowAfterEqual, pModName, pItemTextToSearch)
        
        
        If Not m_ActiveListViewControl.Visible Then
            X = GetNormalisedColumn(Me, rtEditor)
            Y = GetNormalisedLine(Me, rtEditor)
            
            pPos = InStr(1, Replace(Text, vbTab, "      "), m_PartialText)
            If pPos > 0 Then
                If pShowTip Then
                    X = (pPos) * UserControl.TextWidth("W")
                Else
                    If InStr(1, Text, vbTab) > 0 Then
                        X = Len(Replace(Text, vbTab, "      ")) * UserControl.TextWidth("W")
                    Else
                        X = (X) * UserControl.TextWidth("W")
                    End If

                End If
            Else
                If InStr(1, Text, vbTab) > 0 Then
                    X = Len(Replace(Text, vbTab, "      ")) * UserControl.TextWidth("W")
                Else
                    X = (X) * UserControl.TextWidth("W")
                End If
            End If
            Y = (Y + 0.5) * UserControl.TextHeight("y")
                
            If Not pShowTip Then
                ctrlTips.ShowTip False
                If pOperatorType = otDot Or pOperatorType = otEqualTo Then
                    Call ShowListView(X, Y)
                    If m_OperatorType = otUnknown Then
                        pItemTextToSearch = pUnformatedText
                        GoTo lblTryVBAFunctionTips
                    End If
                Else
                    If Not m_ActiveListViewControl Is Nothing Then
                        m_ActiveListViewControl.Visible = False
                    End If
                    Set m_LastSearchedItem = Nothing
                    Set m_ActiveListViewControl = Nothing
                End If
            Else
                
                m_TipPos.X = X / Screen.TwipsPerPixelX
                m_TipPos.Y = Y / Screen.TwipsPerPixelY
                Call ctrlTips.ShowTip(True)
                Call PositionTip
                Set m_ActiveListViewControl = Nothing
                Set m_LastSearchedItem = Nothing
            End If
        Else
            If Not m_ActiveListViewControl Is Nothing Then
                If m_ActiveListViewControl.Visible Then
                    m_ActiveListViewControl.Visible = False
                End If
                Set m_LastSearchedItem = Nothing
                Set m_ActiveListViewControl = Nothing
            
            End If
        End If
    
    Else
lblTryVBAFunctionTips:
        Dim pItem As ListItem
         X = GetNormalisedColumn(Me, rtEditor) - 1
         If X = 0 Then X = 1
        pPos = InStr(X, pItemTextToSearch, "(")
        If pPos = X Then
            pItemTextToSearch = Left(pItemTextToSearch, pPos - 1)
            pPos = InStr(1, StrReverse(pItemTextToSearch), " ", vbTextCompare)
            If pPos > 0 Then
                pPos = Len(pItemTextToSearch) - pPos
               pItemTextToSearch = Trim(Mid(pItemTextToSearch, pPos + 2))
            End If
            
            pPos = InStr(1, StrReverse(pItemTextToSearch), "(", vbTextCompare)
            If pPos > 0 Then
                pPos = Len(pItemTextToSearch) - pPos
                pItemTextToSearch = Trim(Mid(pItemTextToSearch, pPos + 2))
            End If
                
            
            GoTo lblItemFound
        End If
        If pPos <= 0 Then pPos = 1
        pPos = InStr(pPos, pItemTextToSearch, ")")
        If pPos <= 0 Then
            pPos = InStr(1, pItemTextToSearch, "(")
            
            If pPos > 0 Then
                pItemTextToSearch = Left(pItemTextToSearch, pPos - 1)
                
                pPos = InStr(1, StrReverse(pItemTextToSearch), " ")
                If pPos > 0 Then
                    pPos = Len(pItemTextToSearch) - pPos
                    pItemTextToSearch = Trim(Mid(pItemTextToSearch, pPos + 1))
                End If
lblItemFound:
                
                Set pItem = lvwVBA.FindItem(pItemTextToSearch, lvwText, , lvwPartial)
                If Not pItem Is Nothing Then
                    If InStr(1, pItemTextToSearch, pItem.Text, vbTextCompare) <= 0 Then
                        Set pItem = Nothing
                    End If
                End If
                If Not pItem Is Nothing Then
                    If m_OperatorType = otUnknown Then
                        'If Not m_ActiveListViewControl Is Nothing Then
                        '    m_ActiveListViewControl.Visible = False
                        'End If
                        'Set m_ActiveListViewControl = Nothing
                        'Set m_LastSearchedItem = Nothing
                        If Len(pItem.Tag) > 0 Then
                            X = GetNormalisedColumn(Me, rtEditor)
                            Y = GetNormalisedLine(Me, rtEditor)
                            Debug.Print " X= ", X
                            pPos = InStr(1, Replace(Text, vbTab, "      "), m_PartialText)
                            If pPos > 0 Then
                                If pShowTip Then
                                    If pPos > X Then
                                        pPos = X
                                    End If
                                    X = (pPos) * UserControl.TextWidth("W")
                                Else
                                    If InStr(1, Text, vbTab) > 0 Then
                                        pPos = InStr(1, Replace(Text, vbTab, "      "), m_PartialText)
                                        If pPos > 0 Then
                                            pPos = pPos - 1
                                        End If
                                        
                                        If pPos > X Then
                                            pPos = X
                                            pPos = pPos - Len(m_PartialText)
                                        End If
                                        X = (pPos) * UserControl.TextWidth("W")
                                        
                                    Else
                                        If X > Len(m_PartialText) Then
                                            pPos = Abs(X - Len(m_PartialText) - 1)
                                        End If
                                        If pPos > X Then
                                            pPos = X
                                            pPos = pPos - Len(m_PartialText)
                                        End If
                                        X = (pPos) * UserControl.TextWidth("W")
                                    End If
                
                                End If
                            Else
                                If InStr(1, Text, vbTab) > 0 Then
                                    pPos = (Len(Replace(Text, vbTab, "      ")) - Len(m_PartialText)) * UserControl.TextWidth("W")
                                    If pPos > X Then
                                        pPos = X
                                        pPos = pPos - Len(m_PartialText)
                                    Else
                                        X = pPos
                                    End If
                                Else
                                    X = (X) * UserControl.TextWidth("W")
                                End If
                            End If
                            Y = (Y + 0.5) * UserControl.TextHeight("y")
                            
                            ctrlTips.TipText = pItem.Tag
                            m_TipPos.X = X / Screen.TwipsPerPixelX
                            m_TipPos.Y = Y / Screen.TwipsPerPixelY
                            Call ctrlTips.ShowTip(True)
                            Call PositionTip
                         '   m_OperatorType = otShowTip
                        End If
                        
                    End If
                End If
           
            End If
        End If
    End If
    
    Dim pSeps() As String
    If ctrlTips.Visible Then
        pPos = InStr(1, StrReverse(Text), ")")
        If pPos <= 0 Then
            pPos = InStr(1, StrReverse(Text), "(")
            If pPos > 0 Then
                pPos = Len(Text) - pPos
                If InStr(pPos, Text, ",") > 0 Then
                  pSeps = Split(Mid(Text, pPos), ",")
                  ctrlTips.FunctionTips1.BoldIndex = UBound(pSeps)
                End If
            End If
        End If
    End If
    Exit Sub
ErrorTrap:
    MsgBox "rtEditor:Change:" & Err.Description, vbCritical
    Resume
End Sub

Friend Sub PositionTip()
    Dim pTipPos As POINTAPI
    Dim pPosChanged As Boolean
    Dim pActiveWindow As Long
    On Error Resume Next
    pActiveWindow = GetActiveWindow()
    
    If Not UserControl.Parent Is Nothing Then
        If UserControl.Parent.hwnd <> pActiveWindow Then
            If Err.Number = 0 Then
                ctrlTips.ShowTip False
                Exit Sub
            End If
        End If
    End If
    
    LSet pTipPos = m_TipPos
    Call ClientToScreen(rtEditor.hwnd, pTipPos)
    If ctrlTips.Left <> pTipPos.X * Screen.TwipsPerPixelX Then
        ctrlTips.Left = pTipPos.X * Screen.TwipsPerPixelX
        pPosChanged = True
    End If
    If ctrlTips.Top <> pTipPos.Y * Screen.TwipsPerPixelY Then
        ctrlTips.Top = pTipPos.Y * Screen.TwipsPerPixelY
        pPosChanged = True
    End If
    If pPosChanged Then
        ctrlTips.ShowTip True
    End If
    Exit Sub
ErrorTrap:
End Sub


Private Function FindCurrentModuleName(ByVal pCurrentpos As Long) As String
    Dim pFullText As String
    Dim pNarrowText As String
    Dim pPos As Long
    FindCurrentModuleName = GLOBAL_MODULE
    pFullText = Left(rtEditor.Text, pCurrentpos)
    pFullText = Replace(pFullText, vbTab, " ")
    pFullText = Replace(pFullText, vbCrLf, " ")
    
    pPos = InStr(1, StrReverse(pFullText), " etavirP", vbTextCompare)
    If pPos > 0 Then
        pPos = Len(pFullText) - pPos
        pFullText = Mid(pFullText, pPos + 1)
    Else
        'pFullText = Mid(pFullText, pPos + 1)
        pPos = InStr(1, StrReverse(pFullText), " buS")
        If pPos > 0 Then
            pPos = Len(pFullText) - pPos
            pFullText = Mid(pFullText, pPos - 3)
        Else
            pPos = InStr(1, StrReverse(pFullText), " noitcnuF")
            If pPos > 0 Then
                pPos = Len(pFullText) - pPos
                pFullText = Mid(pFullText, pPos - 8)
            Else
                FindCurrentModuleName = GLOBAL_MODULE
                Exit Function
            End If
        End If
    End If
    pPos = InStr(1, pFullText, "Sub ")
    If pPos <= 0 Then
        pPos = InStr(1, pFullText, "Function ")
        pNarrowText = Mid(pFullText, pPos + 9)
        pPos = InStr(1, pNarrowText, " ")
        If pPos > 0 Then
            pNarrowText = Trim(Left(pNarrowText, pPos)) & "("
        End If
    Else
        pNarrowText = Mid(pFullText, pPos + 4)
        
        pPos = InStr(1, pNarrowText, " ")
        If pPos > 0 Then
            pNarrowText = Trim(Left(pNarrowText, pPos)) & "("
        End If
    End If
    
    pPos = InStr(1, pNarrowText, "(", vbTextCompare)
    If pPos > 0 Then
        FindCurrentModuleName = Left(pNarrowText, pPos - 1)
    Else
        pPos = InStr(1, pNarrowText, vbCrLf, vbTextCompare)
        If pPos > 0 Then
            FindCurrentModuleName = Left(pNarrowText, pPos - 1)
        End If
    End If
    
End Function



Private Sub PopulateVariableProprs(ByVal pShowAfterEqual As Boolean, ByVal pModName As String, ByVal pVar As String)
    Dim pModule As Module
    Dim pVariable As Variable
    Dim pColl As Collection
    Dim pTypeInfo As cTypeLibInfo
    Dim pImageKey As String
    Dim pItem As ListItem
    Dim pVarName As String
    Dim pExtVars As String
    Dim pPos As Long
    Dim pReturnString As String
    
    pVar = Replace(pVar, vbCrLf, "")
    pVar = Replace(pVar, vbTab, " ")
    pVar = Trim(pVar)
    pPos = InStr(1, pVar, ".", vbTextCompare)
    If pPos > 0 Then
        pVarName = Trim(Left(pVar, pPos - 1))
        pExtVars = Mid(pVar, pPos + 1)
    Else
        pVarName = Trim(pVar)
    End If
    
    pVarName = Trim(pVarName)
    
    pPos = InStr(1, StrReverse(pVarName), " ")
    If pPos > 0 Then
        pPos = Len(pVarName) - pPos
        pVarName = Mid(pVarName, pPos + 1)
    End If
    pVarName = Trim(pVarName)
    lvwVariables.ListItems.Clear
    Set pModule = g_Project.Modules.Item(pModName)
    If Not pModule Is Nothing Then
        Set pVariable = pModule.Variables(pVarName)
        If pVariable Is Nothing Then
            Set pModule = g_Project.Modules.Item(GLOBAL_MODULE)
            Set pVariable = pModule.Variables(pVarName)
        ElseIf pVariable.ObjProgID = "" Then
            Call ParseForModulesAndVars
        End If
    End If
    If Not pVariable Is Nothing Then
        pReturnString = m_Intellisense.GetVarProperties(pShowAfterEqual, pExtVars, pVariable.ObjProgID, pVariable.LibRefrence, pColl)
    Else
        Exit Sub
        Debug.Assert False
    End If
    
    If Not pColl Is Nothing Then
        For Each pTypeInfo In pColl
            Select Case pTypeInfo.TypeKind
                Case iTypeKind.iLibrary:
                    pImageKey = "LIBRARY"
                Case iTypeKind.iModule:
                    pImageKey = "MODULE"
                Case iTypeKind.iClass:
                    pImageKey = "CLASS"
                Case iTypeKind.iEvent:
                    pImageKey = "EVENT"
                Case iTypeKind.iFunction:
                    pImageKey = "FUNCTION"
                Case iTypeKind.iProperty:
                    pImageKey = "PROPERTY"
                Case iTypeKind.iEnum:
                    pImageKey = "ENUM"
                Case iTypeKind.iType:
                    pImageKey = "TYPE"
                Case iTypeKind.iConstant:
                    pImageKey = "CONSTANT"
            End Select
            If pImageKey <> "" Then
                Set pItem = lvwVariables.ListItems.Add(, pTypeInfo.ClassName, pTypeInfo.ClassName, , pImageKey)
                pItem.Tag = pTypeInfo.Tag
            End If
            'pItem.Tag = pTypeInfo.CLSID & "#" & pTypeInfo.Ver
            
            If m_Intellisense.Unload Then Exit Sub
        Next
        lvwVariables.Sorted = True
        
        Call ResizeList(lvwVariables)
    End If
    
    'If Len(pReturnString) > 0 Then
        ctrlTips.TipText = pReturnString
    'End If
    
    
End Sub

Private Sub ResizeList(ByRef pListView As ListView)
    Dim pItem As ListItem
    
    If Not pListView Is Nothing Then
        Call SendMessageLong(pListView.hwnd, _
                           LVM_SETCOLUMNWIDTH, _
                           0, _
                           ByVal LVSCW_AUTOSIZE)
        pListView.ColumnHeaders(1).Width = pListView.ColumnHeaders(1).Width + 200
        pListView.Width = pListView.ColumnHeaders(1).Width + 250
        
        If pListView.ListItems.Count > 7 Then
                pListView.Height = 7 * pListView.ListItems.Item(1).Height
        Else
            If pListView.ListItems.Count > 0 Then
                pListView.Height = pListView.ListItems.Item(1).Height * pListView.ListItems.Count
            End If
        End If
    End If
End Sub



' Scroll the contents of a ListView control horizontally and vertically.
'
' If the ListView control is in List mode, DX is the number of
' columns to scroll. If it is in Report mode, DY is rounded to the
' nearest number of pixels that represent a whole line.

Private Sub ListViewScroll(lvw As ListView, ByVal dx As Long, ByVal dy As Long)
    SendMessage lvw.hwnd, LVM_SCROLL, CInt(dx), CInt(dy)
End Sub


Private Sub ShowListView(Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0)
    Dim pPos As POINTAPI
    If Not m_ActiveListViewControl Is Nothing Then
        If m_ActiveListViewControl.ListItems.Count <= 0 Then
            m_ActiveListViewControl.Visible = False
            Set m_ActiveListViewControl = Nothing
            Set m_LastSearchedItem = Nothing
            m_OperatorType = otUnknown
            Exit Sub
                    
        End If
    End If
    If Not m_Intellisense.IsBusyLoading Then
        If X = 0 Or Y = 0 Then
            X = GetNormalisedColumn(Me, rtEditor)
            Y = GetNormalisedLine(Me, rtEditor)
            m_ActiveListViewControl.Left = (X + 2) * UserControl.TextWidth("W")
            m_ActiveListViewControl.Top = (Y + 0.5) * UserControl.TextHeight("y")
        Else
            m_ActiveListViewControl.Left = X - 330
            m_ActiveListViewControl.Top = Y
        End If
        
        Call CheckListViewBound
        
        If m_ActiveListViewControl.ListItems.Count > 0 Then
            m_ActiveListViewControl.ListItems(1).Selected = True
            Set m_LastSearchedItem = m_ActiveListViewControl.ListItems(1)
            m_LastSearchedItem.Bold = True
        End If
        m_ActiveListViewControl.Visible = True
        m_ActiveListViewControl.Refresh
        'End If
    Else
        Set m_ActiveListViewControl = Nothing
        m_OperatorType = otUnknown
    End If
End Sub


Private Sub CheckListViewBound()
    Dim pLeft As Long
    Dim pTop As Long
    Dim pLeftMax As Long
    Dim pTopMax As Long
    Dim pNewLeft As Long
    Dim pNewTop As Long
    Dim pPos As POINTAPI
    'Exit Sub
    With m_ActiveListViewControl
        pLeft = .Left
        pTop = .Top
        pNewLeft = pLeft
        pNewTop = pTop
        pPos.X = Abs(pNewLeft) / Screen.TwipsPerPixelX
        pPos.Y = Abs(pNewTop) / Screen.TwipsPerPixelY
        Call ClientToScreen(UserControl.hwnd, pPos)
        Call ScreenToClient(0, pPos)
        pNewLeft = pPos.X * Screen.TwipsPerPixelX
        pNewTop = pPos.Y * Screen.TwipsPerPixelY
        .Left = Abs(pNewLeft)
        .Top = Abs(pNewTop)
        m_ActiveListViewControl.ZOrder 0
    End With
End Sub

Private Function ColorIt(Text As String, pLenDelta As Long) As String

Dim strLines() As String
Dim strLine As String
Dim strWord() As String
Dim intWord As Integer
Dim strWord1 As String

Dim strRTF As String
Dim strAllRTF As String
Dim strHeader As String

Dim onComment As Boolean
Dim onQuotation As Boolean


Dim pDimFound As Boolean
Dim i As Integer
Dim j As Integer

strLines = Split(Text, vbLf)

' Color
For i = LBound(strLines) To UBound(strLines)

    'Reset
    onComment = False
    onQuotation = False
    
    strLine = strLines(i)
    
    strLine = Replace(strLine, "\", "\\")
    strLine = Replace(strLine, "}", "\}")
    strLine = Replace(strLine, "{", "\{")
    
    ' Replace space to strline
    For j = 0 To 27
        
        strLine = Replace(strLine, Delimiter(j), Delimiter(j) & " ", , , vbTextCompare)
        
    Next j
    
    ' Split line to word
    strWord = Split(strLine, " ")
    
    For j = LBound(strWord) To UBound(strWord)
        
        Select Case UCase(strWord(j))
        
            ' Comment
                
            Case "'"
                
                If onQuotation = False Then
                    If onComment = False Then
                
                        onComment = True
                        strWord(j) = "\cf4 " & strWord(j)
                        
                        GoTo EndLine
                    
                    End If
                End If
            
            ' Quotation
            Case Chr(34)
            
                If onComment = False Then
                    If onQuotation = False Then
                
                        onQuotation = True
                        strWord(j) = "\cf5" & strWord(j)
                        
                        GoTo EndIt
                
                    Else
                
                        onQuotation = False
                        strWord(j) = strWord(j) & "\cf0"
                        
                        GoTo EndIt
                
                    End If
                End If
                
            ' Comment
            Case "REM"
                
                If onQuotation = False Then
                    If onComment = False Then
                
                        onComment = True
                        strWord(j) = "\cf4 " & strWord(j)
                        
                        GoTo EndLine
                    
                    End If
                End If
            
            Case Else
                
                intWord = InStr(1, strDelimiter, Right(strWord(j), 1))
                
                If intWord > 0 Then
                    
                    
                    strWord1 = Delimiter(intWord - 1)
                    If Len(strWord(j)) <= 0 Then GoTo EndIt
                    If Len(strWord(j)) <> Len(strWord1) Then
                        strWord(j) = Left(strWord(j), Len(strWord(j)) - Len(strWord1))
                    Else
                        strWord(j) = strWord1
                        strWord1 = ""
                        intWord = 0
                        GoTo EndIt
                    End If
                End If
                Dim pTextToReplace As String
                Dim pPos As Long
                
                If onQuotation = False Then
                    
                    pPos = InStr(1, SM_RESERVEDWORDS, " " & Replace(strWord(j), vbCr, "") & " ", vbTextCompare)
                    If pPos > 0 Then
                    
                        If Len(strWord(j)) > 1 Then
                            pTextToReplace = strWord(j)
                            pTextToReplace = Replace(pTextToReplace, vbTab, "")
                            pTextToReplace = Replace(pTextToReplace, vbCr, "")
                            pTextToReplace = Replace(pTextToReplace, vbLf, "")
                            
                            'strWord(j) = UCase(Left(strWord(j), 1)) & LCase(Right(strWord(j), Len(strWord(j)) - 1))
                            strWord(j) = Replace(strWord(j), pTextToReplace, Trim(Mid(SM_RESERVEDWORDS, pPos, Len(strWord(j)) + 1)))
                            If strWord(j) = "Dim" Then
                                pDimFound = True
                            End If
                        Else
                            pPos = 0
                        End If
                        If pDimFound Then
                            If strWord(j) = "As" Then
                                strWord(j) = "'" & strWord(j)
                                pLenDelta = pLenDelta + 1
                            End If
                        End If
                        strWord(j) = "\cf2\b0 " & strWord(j) & "\b0\cf0 "
                    End If
                    
                    If InStr(1, SM_FUNCTIONWORDS, " " & Replace(strWord(j), vbCr, "") & " ", vbTextCompare) > 0 Then
                        
                        'strWord(j) = "\cf3 " & strWord(j) & "\cf0 "
                        If InStr(1, strWord(j), vbCr) Then
                            If Len(strWord(j)) > 1 Then
                                strWord(j) = "\cf3" & Trim(Mid(SM_FUNCTIONWORDS, InStr(1, SM_FUNCTIONWORDS, Replace(strWord(j), vbCr, ""), vbTextCompare), Len(strWord(j)))) & vbCr & "\cf0 "
                            End If
                        Else
                            strWord(j) = "\cf3" & Trim(Mid(SM_FUNCTIONWORDS, InStr(1, SM_FUNCTIONWORDS, strWord(j), vbTextCompare), Len(strWord(j)))) & "\cf0 "
                        End If
                    End If
                    
                End If
                
                If intWord > 0 Then
                
                    ' Comment and Quotation
                    Select Case strWord1
                        ' Comment
                        Case "'"
                            If onQuotation = False Then
                                If onComment = False Then
                
                                    onComment = True
                                    strWord1 = "\cf4 " & strWord1
                                    
                                    GoTo EndColor
                    
                                End If
                            End If
                            
                        'Quotation
                        Case Chr(34)
                        
                            If onComment = False Then
                                If onQuotation = False Then
                
                                    onQuotation = True
                                    strWord1 = "\cf5 " & strWord1
                                    
                                    GoTo EndColor
                        
                
                                Else
                
                                onQuotation = False
                                strWord1 = strWord1 & "\cf0"
                                
                                GoTo EndColor
                
                                End If
                            End If
                    
                    End Select
                    
EndColor:
                
                    strWord(j) = strWord(j) & strWord1
                    
                    If onComment = True Then
                        GoTo EndLine
                    End If
                    
                End If
                
        End Select
        
EndIt:
    
    Next j
EndLine:
        
    strLine = Join(strWord, " ")
    
    For j = 0 To 27
        
        strLine = Replace(strLine, Delimiter(j) & " ", Delimiter(j), , , vbTextCompare)
        
    Next j
    
    
    If onComment = True Then
        strLine = strLine & "\cf0"
    End If
    
    If onQuotation = True Then
        strLine = strLine & "\cf0"
    End If
    
    strLines(i) = strLine
    
Next i

strRTF = Join(strLines, vbLf & "\par ")
strHeader = CreateHeader

strAllRTF = strHeader & strRTF & vbLf & "}"

ColorIt = strAllRTF


End Function

Private Function CreateHeader() As String

Dim H1 As String
Dim H2 As String
Dim ColorH As String
Dim i As Integer

' Color Header
ColorH = "{\colortbl " & ConverColorToRTF(N_Color)
For i = 1 To 2
    ColorH = ColorH & ConverColorToRTF(K_COLOR(i))
Next i

ColorH = ColorH & ConverColorToRTF(C_COLOR)
ColorH = ColorH & ConverColorToRTF(Q_COlOR)
ColorH = ColorH & ";}"

' Header
H1 = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 " & UserControl.Font.Name & ";}}"
H2 = "\viewkind4\uc1\pard\f0\fs" & Round(UserControl.Font.Size * 2) & " "

CreateHeader = H1 & vbLf & ColorH & vbLf & H2

End Function
Private Function ConverColorToRTF(LongColor As Long) As String

    Dim ColorRTFCode As String
    Dim lc As Long
    
    lc = LongColor And &H10000FF
    ColorRTFCode = ";\red" & lc
    lc = (LongColor And &H100FF00) / (2 ^ 8)
    ColorRTFCode = ColorRTFCode & "\green" & lc
    lc = (LongColor And &H1FF0000) / (2 ^ 16)
    ColorRTFCode = ColorRTFCode & "\blue" & lc
    ColorRTFCode = ColorRTFCode & ""
    
    ' Return Var
    ConverColorToRTF = ColorRTFCode
    
End Function





Private Function FindProgID(ByVal pProgID As String) As Refrence
    Dim pRefrence As Refrence
    Dim pClassInfo As cTypeLibInfo
    
    On Error Resume Next
    Set pClassInfo = m_Intellisense.AllClasses.Item(pProgID)
    If Not pClassInfo Is Nothing Then
        Set pRefrence = New Refrence
        With pRefrence
            .BinaryPath = pClassInfo.Path
            .GUID = pClassInfo.CLSID
            .Name = pClassInfo.Name
            .Verion = pClassInfo.Ver
        End With
    End If
    Set FindProgID = pRefrence
    Err.Clear
End Function

Private Sub ParseForModulesAndVars()
    Dim pLines() As String
    Dim pLine As String
    Dim pModuleName As String
    Dim pCurrentModule As Module
    Dim pCounter As Long
    Dim pPos As Long
    Dim pText As String
    Dim pName As String
    Dim pType As String
    Dim pVariable As Variable
    
        pLines = Split(rtEditor.Text, vbCrLf)
    
    Call g_Project.Modules.Clear
    
    pModuleName = GLOBAL_MODULE
    Set pCurrentModule = g_Project.Modules.Add(pModuleName)
    
    For pCounter = LBound(pLines) To UBound(pLines)
        pLine = pLines(pCounter)
        pLine = Replace(pLine, vbTab, " ")
        pLine = Replace(pLine, vbCrLf, "")
        
        'Module Start
        If InStr(1, pLine, "Sub ", vbTextCompare) > 0 Or InStr(1, pLine, "Function ", vbTextCompare) > 0 Then
            pPos = InStr(1, pLine, "Sub ", vbTextCompare)
            If pPos <= 0 Then
                pPos = InStr(1, pLine, "Function ", vbTextCompare)
                If pPos > 0 Then
                    pText = Mid(pLine, pPos + 9)
                End If
            Else
                pText = Mid(pLine, pPos + 4)
            End If
            
            pPos = InStr(1, pText, "(", vbTextCompare)
            If pPos > 0 Then
                pName = Left(pText, pPos - 1)
            Else
                pName = pText
            End If
            pName = Trim(pName)
            Set pCurrentModule = g_Project.Modules.Add(pName)
        'Variable Declaration
        ElseIf (InStr(1, pLine, "Dim ", vbTextCompare) > 0 Or InStr(1, pLine, "Set ", vbTextCompare) > 0) Then
            pPos = InStr(1, pLine, "Dim ", vbTextCompare)
            If pPos <= 0 Then
                pPos = InStr(1, pLine, "Set ", vbTextCompare)
            End If
            If pPos > 0 Then
                pText = Mid(pLine, pPos + 4)
                pPos = InStr(1, pText, " ", vbTextCompare)
                If pPos > 0 Then
                    pName = Trim(Left(pText, pPos))
                    pPos = InStr(1, pText, "'As ", vbTextCompare)
                    If pPos > 0 Then
                        pType = Trim(Mid(pText, pPos + 4))
                    Else
                        pPos = InStr(1, pText, "CreateObject(""", vbTextCompare)
                        pType = Trim(Mid(pText, pPos + 14))
                        pPos = InStr(1, StrReverse(pType), ")""")
                        If pPos > 0 Then
                            pPos = Len(pType) - pPos - 1
                            pType = Left(pType, pPos)
                        End If
                    End If
                    If pName <> "" Then
                        On Error Resume Next
                        Err.Clear
                        If Not pCurrentModule Is Nothing Then pCurrentModule.Variables.Add pName, pType, FindProgID(pType)
                        If Err.Number <> 0 And Not (pCurrentModule Is Nothing) Then
                            Call pCurrentModule.Variables.Remove(pName)
                            pCurrentModule.Variables.Add pName, pType, FindProgID(pType)
                        End If
                    End If
                End If
            End If
            
        'Module Ends
        ElseIf (InStr(1, pLine, "End Sub", vbTextCompare) > 0 Or InStr(1, pLine, "End Function", vbTextCompare) > 0) Then
            pModuleName = GLOBAL_MODULE
            Set pCurrentModule = g_Project.Modules.Item(pModuleName)
        End If
    Next
    
    Call PopulateVBAFunctions
    Dim pItem As ListItem
    Dim pImageKey As String
    For Each pCurrentModule In g_Project.Modules
        pImageKey = "FUNCTION"
        If pCurrentModule.Name <> GLOBAL_MODULE Then
            Set pItem = lvwVBA.ListItems.Add(, pCurrentModule.Name, pCurrentModule.Name, , pImageKey)
        End If
        For Each pVariable In pCurrentModule.Variables
            pImageKey = "PROPERTY"
            Set pItem = lvwVBA.ListItems.Add(, pCurrentModule.Name & "." & pVariable.Name, pVariable.Name, , pImageKey)
        Next
    Next
    Call ResizeList(lvwVBA)
End Sub

Private Function GetTabs(ByVal pInput As String) As String
    Dim pChar As String
    Dim pOutPut As String
    Dim pCharPos As Long
    Dim pMaxPos As Long
    pMaxPos = Len(pInput)
    For pCharPos = 1 To pMaxPos
        pChar = Mid(pInput, pCharPos, 1)
        If pChar = vbTab Then
            pOutPut = pOutPut & pChar
        Else
            Exit For
        End If
    Next
    GetTabs = pOutPut
End Function


Private Sub rtEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim pModName As String
    Dim pLookFrom As Long
    Dim pFound As Boolean
    Dim pTextInQuestion As String
    m_ReturnKeyPressed = False
    If KeyCode = vbKeyBack Then
        'If ctrlTips.Visible Then ctrlTips.ShowTip False
        pTextInQuestion = m_PreviousLine
    Else
        pTextInQuestion = rtEditor.SelText
    End If
    pLookFrom = 1
    If InStr(pLookFrom, pTextInQuestion, "Private ") > 0 Then
        Call ParseForModulesAndVars
    End If
    
    'Delete Variable
    pLookFrom = 1
    If InStr(pLookFrom, pTextInQuestion, "Dim ") > 0 Or InStr(pLookFrom, pTextInQuestion, "Set ") > 0 Then
        ParseForModulesAndVars
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        If ctrlTips.Visible Then ctrlTips.ShowTip False
        If KeyCode = vbKeyReturn Then
            m_TabsInLastLine = GetTabs(m_PreviousLine)
            m_ReturnKeyPressed = True
            If InStr(1, m_PreviousLine, "Private Sub ", vbTextCompare) > 0 Then
                Call ParseForModulesAndVars
            ElseIf InStr(1, m_PreviousLine, "Private Function ", vbTextCompare) > 0 Then
                Call ParseForModulesAndVars
            ElseIf InStr(1, m_PreviousLine, "Dim ", vbTextCompare) > 0 Then
                Call ParseForModulesAndVars
            ElseIf InStr(1, m_PreviousLine, "Set ", vbTextCompare) > 0 Then
                Call ParseForModulesAndVars
            ElseIf InStr(1, m_PreviousLine, "Sub ", vbTextCompare) > 0 Then
                Call ParseForModulesAndVars
            ElseIf InStr(1, m_PreviousLine, "Function ", vbTextCompare) > 0 Then
                Call ParseForModulesAndVars
            End If
            m_PartialText = ""
        End If
        If Not m_ActiveListViewControl Is Nothing Then
            If m_ActiveListViewControl.Visible Then
                If KeyCode = vbKeyEscape Then
                    m_ActiveListViewControl.Visible = False
                    m_OperatorType = otUnknown
                    Set m_ActiveListViewControl = Nothing
                Else
                    Call HideListBox
                End If
            End If
        Else
            Call lvwAllClasses_LostFocus
            Call lvwCreatables_LostFocus
            Call lvwVariables_LostFocus
            Call lvwVBA_LostFocus
        End If
        m_blnIgnoreKey = True
    End If
    
    If KeyCode = vbKeySpace And Shift = vbCtrlMask Then
        KeyCode = 0
'        Call rtEditor_Change
        If m_ActiveListViewControl Is Nothing Then
            m_PartialIniLen = 0 'Len(m_PartialText)
            m_OperatorType = m_Intellisense.ParseOperator(rtEditor, m_PreviousLine, m_PartialText)
            
            
            If m_OperatorType = otCreateObject Then
                Set m_ActiveListViewControl = lvwCreatables
            ElseIf m_OperatorType = otDimAs Then
                Set m_ActiveListViewControl = lvwAllClasses
            ElseIf m_OperatorType = otDot Then
                Set m_ActiveListViewControl = lvwVariables
            Else
                Set m_ActiveListViewControl = lvwVBA
            End If
            
            
            If Not FindAppropriate() Then
                Call ShowListView
                Call rtEditor_Change
            Else
                Set m_ActiveListViewControl = Nothing
            End If
            
        Else
            Call ShowListView
        End If
        
    End If
    
    Dim pNextItem As ListItem
    
    If KeyCode = vbKeyDown Then
        If Not m_ActiveListViewControl Is Nothing Then
            If m_ActiveListViewControl.Visible Then
                KeyCode = 0
                If m_LastSearchedItem Is Nothing Then
                    If m_ActiveListViewControl.ListItems.Count > 0 Then
                        Set m_LastSearchedItem = m_ActiveListViewControl.ListItems(1)
                        m_LastSearchedItem.Bold = True
                        m_LastSearchedItem.Selected = True
                    End If
                End If
                If Not m_LastSearchedItem Is Nothing Then
                    If m_LastSearchedItem.Index < m_ActiveListViewControl.ListItems.Count Then
                        m_LastSearchedItem.Bold = False
                        Set pNextItem = m_ActiveListViewControl.ListItems.Item(m_LastSearchedItem.Index + 1)
                        Set m_LastSearchedItem = pNextItem
                        m_LastSearchedItem.Selected = True
                        m_LastSearchedItem.Bold = True
                        Call ListViewScroll(m_ActiveListViewControl, 0, 1)
                        m_LastSearchedItem.EnsureVisible
                    End If
                End If
                
            End If
        End If
        If ctrlTips.Visible Then ctrlTips.ShowTip False
    ElseIf KeyCode = vbKeyUp Then
        If Not m_ActiveListViewControl Is Nothing Then
            If m_ActiveListViewControl.Visible Then
                KeyCode = 0
                If m_ActiveListViewControl.Visible Then
                    If m_LastSearchedItem Is Nothing Then
                        If m_ActiveListViewControl.ListItems.Count > 0 Then
                            Set m_LastSearchedItem = m_ActiveListViewControl.ListItems(1)
                            m_LastSearchedItem.Bold = True
                            m_LastSearchedItem.Selected = True
                        End If
                    End If
                    If Not m_LastSearchedItem Is Nothing Then
                        If m_LastSearchedItem.Index > 1 Then
                            m_LastSearchedItem.Bold = False
                            Set pNextItem = m_ActiveListViewControl.ListItems.Item(m_LastSearchedItem.Index - 1)
                            Set m_LastSearchedItem = pNextItem
                            m_LastSearchedItem.Selected = True
                            m_LastSearchedItem.Bold = True
                            Call ListViewScroll(m_ActiveListViewControl, 0, 1)
                            m_LastSearchedItem.EnsureVisible
                        End If
                    End If
                End If

            End If
        End If
        If ctrlTips.Visible Then ctrlTips.ShowTip False
    End If
    
        If KeyCode = vbKeyTab Then
            
            If Not m_ActiveListViewControl Is Nothing Then
                If m_ActiveListViewControl.Visible Then
                    
                    Call HideListBox
                    KeyCode = 0
                End If
            End If
        End If

End Sub

Private Function FindAppropriate() As Boolean
    Dim pTextToSearch As String
    Dim pSearchedItem As ListItem
    Dim pNextItem As ListItem
    Dim pPos As Long
    
    If Not m_ActiveListViewControl Is Nothing Then
        If Not m_ActiveListViewControl.Visible Then
        
            pTextToSearch = m_PartialText
            Set pSearchedItem = m_ActiveListViewControl.FindItem(pTextToSearch, lvwText, , lvwPartial)
            If Not pSearchedItem Is Nothing Then
                If pSearchedItem.Index + 1 <= m_ActiveListViewControl.ListItems.Count Then
                    Set pNextItem = m_ActiveListViewControl.ListItems(pSearchedItem.Index + 1)
                    If InStr(1, pNextItem.Text, pTextToSearch, vbTextCompare) > 0 Then
                        FindAppropriate = False
                    Else
                        FindAppropriate = True
                        Set m_LastSearchedItem = pSearchedItem
                        m_LastSearchedItem.Selected = True
                        
                        Call HideListBox(True)
                        
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub rtEditor_LostFocus()
    ctrlTips.ShowTip False
End Sub

Private Sub rtEditor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctrlTips.ShowTip False
    If Not m_ActiveListViewControl Is Nothing Then
        If m_ActiveListViewControl.Visible Then
            m_ActiveListViewControl.Visible = False
            m_OperatorType = otUnknown
            Set m_ActiveListViewControl = Nothing
        End If
    End If
End Sub

Private Sub rtEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_ActiveListViewControl Is Nothing Then
        m_ActiveListViewControl.ZOrder 0
    End If
End Sub

Private Sub rtEditor_SelChange()
    On Error GoTo ErrorTrap
    Dim pLine As Long
    Dim pCol As Long
    Dim pLines() As String
    Dim pText As String
    Dim pSpaceLen As Long
    Dim pItemTextToSearch As String
    
    If m_LoadingScript Then Exit Sub
    With rtEditor
        pLine = GetCurrentLine(rtEditor)
        pCol = GetCurrentColumn(rtEditor)
        
        If pCol < 0 Then Exit Sub
        
        SB1.Panels(3).Text = "Ln " & pLine & ", Col " & pCol
        pLines = Split(.Text, vbCrLf)
        If UBound(pLines) >= (pLine - 1) Then
            pText = pLines(pLine - 1)
        Else
            pText = .Text
        End If
        
        pSpaceLen = InStr(pCol, pText, " ", vbTextCompare)
       
        If pSpaceLen >= pCol Then
            pText = Trim(Left(pText, pSpaceLen))
        End If
        
        m_PreviousLine = pText
        pText = StrReverse(pText)
        pSpaceLen = InStr(1, pText, " ", vbTextCompare)
        If pSpaceLen > 0 Then
            pText = Left(pText, pSpaceLen)
        End If
        pText = StrReverse(pText)
        pText = Replace(pText, vbTab, "")
        m_PartialText = Trim(pText)
        If m_OperatorType = otCreateObject Then
            m_PartialText = Replace(m_PartialText, "CreateObject(""", "")
            
        ElseIf m_OperatorType = otDimAs Then
        ElseIf m_OperatorType = otDot Then
            pSpaceLen = InStr(1, StrReverse(pText), ".")
            If pSpaceLen > 0 Then
                pSpaceLen = Len(pText) - pSpaceLen
                
            End If
            m_PartialText = Mid(pText, pSpaceLen + 2)
            
        ElseIf m_OperatorType = otEqualTo Then
          ' Debug.Assert False
          '
        End If
        If m_OperatorType = otDot Then
        Dim pPos As Long
        pPos = InStr(1, StrReverse(m_PartialText), ".")
        If pPos > 0 Then
            pPos = Len(m_PartialText) - pPos
            m_PartialText = Mid(m_PartialText, pPos + 2)
        End If
        End If
        If Not m_ActiveListViewControl Is Nothing Then
            If m_ActiveListViewControl.Visible Then
                    pItemTextToSearch = m_PartialText
                     If Not m_LastSearchedItem Is Nothing Then
                        m_LastSearchedItem.Bold = False
                    End If
                    If Len(pItemTextToSearch) > 0 Then
                        Set m_LastSearchedItem = m_ActiveListViewControl.FindItem(pItemTextToSearch, lvwText, , lvwPartial)
                    Else
                        If m_ActiveListViewControl.ListItems.Count > 0 Then
                            Set m_LastSearchedItem = m_ActiveListViewControl.ListItems.Item(1)
                        End If
                    End If
                    
                    If m_LastSearchedItem Is Nothing Then
                        If m_ActiveListViewControl.ListItems.Count > 0 Then
                            Set m_LastSearchedItem = m_ActiveListViewControl.ListItems.Item(1)
                        End If
                    End If
                    If Not m_LastSearchedItem Is Nothing Then
                        
                        Call ListViewScroll(m_ActiveListViewControl, 0, 1)
                        m_LastSearchedItem.EnsureVisible
                        m_LastSearchedItem.Bold = True
                        m_LastSearchedItem.Selected = True
                    Else
                       ' Call HideListBox
                    End If
            End If
        End If
        
    End With
        Exit Sub
ErrorTrap:
    MsgBox "rtEditor:Change:" & Err.Description, vbCritical

End Sub

Private Sub SB1_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "REF" Then
        If m_frmRefrences Is Nothing Then
            Set m_frmRefrences = New frmRefrences
        End If
        Set m_frmRefrences.m_Editor = Me
        Call m_frmRefrences.PrepareUI
        m_frmRefrences.Show vbModal, Me
        Call LoadRefrences
    End If
End Sub

Private Sub UserControl_GotFocus()

   If Not IsLoaded Then
        'This is only to get the Cancel property without having to monitor
        'KeyPress or KeyDown for the Esc key.
        'So we make the cancel button invisible by moving it off the screen
        'Command1(0).Left = Command1(0).Width * -2
        IsLoaded = True
        'CodeMain.TextRTF = ColorIt(Trim$(nd.Text))
        rtEditor.SelStart = 0
        isDirty = False
        DoEvents
    End If
End Sub

Private Sub UserControl_Hide()
    m_Intellisense.Unload = True
End Sub

Private Sub UserControl_Initialize()
    Dim i As Long
    N_Color = vbBlack   '' Normal Text Color
    C_COLOR = RGB(0, 128, 0) '' Comment Text Color
    Q_COlOR = RGB(0, 128, 128) ''Quoation Text Color
    K_COLOR(1) = RGB(0, 0, 128) '' SM_RESERVEDWORDS Color
    K_COLOR(2) = RGB(128, 0, 64) '' Function Wrold Color
    rtEditor.RightMargin = UserControl.TextWidth("A") * 3000
    strDelimiter = ",(){}[]-+*%/='~!&|<>?:;.#@" & Chr(34) & vbTab
    
    For i = 0 To Len(strDelimiter) - 1
        'Delimiter
        Delimiter(i) = Mid(strDelimiter, i + 1, 1)
        
        Select Case Delimiter(i)
            
            Case "\"
                Delimiter(i) = "\\"
            Case "}"
                Delimiter(i) = "\}"
            Case "{"
                Delimiter(i) = "\{"
            
        End Select
        
    Next i
    Set m_Intellisense = New cIntellisense
    Set g_Project = New Project
    Call g_Project.Modules.Add(GLOBAL_MODULE)
    'Loading...
    LastLine = -1
    
End Sub


Private Sub PopulateCreatableClasses()
     Dim pClass As cTypeLibInfo
     Dim pImageKey As String
     Dim pItem As ListItem
     lvwCreatables.ListItems.Clear
     For Each pClass In m_Intellisense.CreatableClasses
        Select Case pClass.TypeKind
            Case iTypeKind.iLibrary:
                pImageKey = "LIBRARY"
            Case iTypeKind.iModule:
                pImageKey = "MODULE"
            Case iTypeKind.iClass:
                pImageKey = "CLASS"
            Case iTypeKind.iEvent:
                pImageKey = "EVENT"
            Case iTypeKind.iFunction:
                pImageKey = "FUNCTION"
            Case iTypeKind.iProperty:
                pImageKey = "PROPERTY"
            Case iTypeKind.iEnum:
                pImageKey = "ENUM"
            Case iTypeKind.iType:
                pImageKey = "TYPE"
            Case iTypeKind.iConstant:
                pImageKey = "CONSTANT"
        End Select
        If pClass.TypeKind <> iLibrary Then
            Set pItem = lvwCreatables.ListItems.Add(, pClass.ProgID, pClass.ClassName, , pImageKey)
            pItem.Tag = pClass.Tag
        End If
        If m_Intellisense.Unload Then Exit Sub
     Next
     lvwCreatables.Sorted = True
     Call ResizeList(lvwCreatables)
End Sub

Private Sub PopulateAllClasses()
     Dim pClass As cTypeLibInfo
     Dim pImageKey As String
     Dim pVisible As Boolean
     Dim pItem As ListItem
     lvwAllClasses.ListItems.Clear
     For Each pClass In m_Intellisense.AllClasses
        pImageKey = "BLANK"
        Select Case pClass.TypeKind
            Case iTypeKind.iLibrary:
                pImageKey = "LIBRARY"
            Case iTypeKind.iModule:
                pImageKey = "MODULE"
            Case iTypeKind.iClass:
                pImageKey = "CLASS"
            Case iTypeKind.iEvent:
                pImageKey = "EVENT"
            Case iTypeKind.iFunction:
                pImageKey = "FUNCTION"
            Case iTypeKind.iProperty:
                pImageKey = "PROPERTY"
            Case iTypeKind.iEnum:
                pImageKey = "ENUM"
            Case iTypeKind.iType:
                pImageKey = "TYPE"
            Case iTypeKind.iConstant:
                pImageKey = "CONSTANT"
            Case iTypeKind.iDataType:
                pImageKey = "VAR"
        End Select
        If pClass.TypeKind <> iLibrary Then
            Set pItem = lvwAllClasses.ListItems.Add(, pClass.ProgID, pClass.ClassName, , pImageKey)
            pItem.Tag = pClass.Tag
        End If
        Set pClass = Nothing
        If m_Intellisense.Unload Then Exit Sub
     Next
     
    lvwAllClasses.ListItems.Add , "Boolean", "Boolean", , "VAR"
    lvwAllClasses.ListItems.Add , "Byte", "Byte", , "VAR"
    lvwAllClasses.ListItems.Add , "Currency", "Currency", , "VAR"
    lvwAllClasses.ListItems.Add , "Date", "Date", , "VAR"
    lvwAllClasses.ListItems.Add , "Double", "Double", , "VAR"
    lvwAllClasses.ListItems.Add , "Integer", "Integer", , "VAR"
    lvwAllClasses.ListItems.Add , "Long", "Long", , "VAR"
    lvwAllClasses.ListItems.Add , "Object", "Object", , "VAR"
    lvwAllClasses.ListItems.Add , "Single", "Single", , "VAR"
    lvwAllClasses.ListItems.Add , "String", "String", , "VAR"
    lvwAllClasses.ListItems.Add , "Variant", "Variant", , "VAR"
    
            
    lvwAllClasses.Sorted = True
    Call ResizeList(lvwAllClasses)
End Sub

Private Sub PopulateVBAFunctions()
     Dim pClass As cTypeLibInfo
     Dim pImageKey As String
     Dim pItem As ListItem
     lvwVBA.ListItems.Clear
     SM_FUNCTIONWORDS = " "
     For Each pClass In m_Intellisense.VBAFunctions
        pImageKey = "BLANK"
        Select Case pClass.TypeKind
            Case iTypeKind.iLibrary:
                pImageKey = "LIBRARY"
            Case iTypeKind.iModule:
                pImageKey = "MODULE"
            Case iTypeKind.iClass:
                pImageKey = "CLASS"
            Case iTypeKind.iEvent:
                pImageKey = "EVENT"
            Case iTypeKind.iFunction:
                pImageKey = "FUNCTION"
            Case iTypeKind.iProperty:
                pImageKey = "PROPERTY"
            Case iTypeKind.iEnum:
                pImageKey = "ENUM"
            Case iTypeKind.iType:
                pImageKey = "TYPE"
            Case iTypeKind.iConstant:
                pImageKey = "CONSTANT"
        End Select
        If pClass.TypeKind <> iLibrary Then
            Set pItem = lvwVBA.ListItems.Add(, pClass.ProgID, pClass.ClassName, , pImageKey)
            pItem.Tag = pClass.Tag
            If pClass.TypeKind = iFunction Then
                SM_FUNCTIONWORDS = SM_FUNCTIONWORDS & pClass.ClassName & " "
            End If
        End If
        Set pClass = Nothing
        If m_Intellisense.Unload Then Exit Sub
     Next
     lvwVBA.Sorted = True
      Call ResizeList(lvwVBA)
End Sub

Public Property Get ScriptRTF() As String
    ScriptRTF = rtEditor.TextRTF
End Property

Public Property Let ScriptRTF(ByVal pScriptRTF As String)
    m_LoadingScript = True
    rtEditor.TextRTF = pScriptRTF
    m_LoadingScript = False
End Property

Public Property Get Script() As String
    Script = rtEditor.Text
End Property

Public Property Let Script(ByVal pScript As String)
     Dim pLines() As String
     Dim pLineCounter As Long
     Dim pLine As String
     Dim pCounter
     LastLine = 0
     If Len(pScript) > 0 Then
        pLines = Split(pScript, vbCrLf)
        pLineCounter = UBound(pLines)
        m_LoadingScript = True
        For pCounter = 0 To pLineCounter
           pLine = pLines(pCounter)
           rtEditor.SelRTF = ColorIt(pLine & vbCrLf, 0)
        Next
        m_LoadingScript = False
    Else
        rtEditor.Text = ""
    End If
     
End Property

Private Sub UserControl_InitProperties()
   InitTabTrap True
   
End Sub


Private Sub HideListBox(Optional pLoadAppropriateWord As Boolean = False)
    Dim pItem As ListItem
    Dim pItemTextToSearch As String
    
    ' Update Color
    Dim OStart As Long
    Dim OLen As Long
    Dim pLenDelta As Long
    Dim StartPos As Long
    Dim EndPos As Long
    Dim pLines() As String
    Dim EndLine As Integer
    Dim StartLine As Integer
    Dim X As Long
    Dim pTextBefore As String
    Dim Text As String
    Dim pPos As Long
    Dim pTempText As String
    X = GetNormalisedColumn(Me, rtEditor) - 1
    If Not m_ActiveListViewControl Is Nothing Then
        If m_ActiveListViewControl.Visible Or pLoadAppropriateWord Then
            If Not m_LastSearchedItem Is Nothing Then
                    'If Len(m_PartialText) > m_PartialIniLen Then
                    '    pItemTextToSearch = Right(m_PartialText, Len(m_PartialText) - m_PartialIniLen)
                    '    pItemTextToSearch = Replace(pItemTextToSearch, vbTab, "")
                    '    pItemTextToSearch = Replace(pItemTextToSearch, vbCr, "")
                    '    pItemTextToSearch = Replace(pItemTextToSearch, vbLf, "")
                    '    pItemTextToSearch = Trim(pItemTextToSearch)
                    'End If
                    m_PartialText = Replace(m_PartialText, "=", "")
                    pItemTextToSearch = m_PartialText
                    Set pItem = m_ActiveListViewControl.SelectedItem
                    If Not pItem Is Nothing Then
                         With rtEditor
                                If m_LoadingScript And Not pLoadAppropriateWord Then Exit Sub
                                'If .Text = "" Then Exit Sub
                                 Call LockWindowUpdate(.hwnd)
                                If LastStart > .SelStart Then
                                    EndLine = .GetLineFromChar(LastStart)
                                    StartLine = .GetLineFromChar(.SelStart)
                                Else
                                    StartLine = .GetLineFromChar(LastStart)
                                    EndLine = .GetLineFromChar(.SelStart)
                                End If
                                pLines = Split(.Text, vbCrLf)
                                StartPos = SendMessage(.hwnd, EM_LINEINDEX, StartLine, 0&)
                                If UBound(pLines) > 0 Then
                                    EndPos = StartPos + Len(pLines(StartLine)) 'SendMessage(.hwnd, EM_LINEINDEX, EndLine + 1, 0&)
                                Else
                                    EndPos = StartPos + SendMessage(.hwnd, EM_LINEINDEX, EndLine + 1, 0&)
                                End If
                                
                                If EndPos <= 0 Then EndPos = Len(.Text)
                                m_LoadingScript = True
                                OStart = .SelStart
                                OLen = .SelLength
                                                    
                                .SelStart = StartPos
                                .SelLength = EndPos - StartPos
                                
                                '.SelColor = N_Color
                                '.SelBold = False
                                Dim pSpacePos As Long
                                If Len(pItemTextToSearch) > 0 Then
                                    
                                    pTempText = Left(.SelText, X)
                                    
                                    If InStr(1, StrReverse(pTempText), " ", vbTextCompare) > 0 Then
                                        pSpacePos = Len(pTempText) - InStr(1, StrReverse(pTempText), " ", vbTextCompare) + 1
                                    Else
                                        pSpacePos = 1
                                    End If
                                    If Len(.SelText) > 0 Then
                                        If pSpacePos > 1 Then
                                            Text = Left(.SelText, pSpacePos) & Replace(Mid(.SelText, pSpacePos + 1), pItemTextToSearch, pItem.Text)
                                        Else
                                            pPos = InStr(1, StrReverse(.SelText), ".")
                                            If pPos > 0 Then
                                                pPos = Len(.SelText) - pPos
                                                Text = Left(.SelText, pPos) & Replace(Mid(.SelText, pPos + 1), pItemTextToSearch, pItem.Text)
                                            Else
                                                Text = Replace(Mid(.SelText, pSpacePos), pItemTextToSearch, pItem.Text)
                                            End If
                                        End If
                                    Else
                                        Text = ""
                                    End If
                                Else
                                    Text = .SelText & pItem.Text
                                End If
                                
                                If Not pItem Is Nothing Then
                                    If Len(pItem.Tag) > 0 Then
                                        pSpacePos = InStr(1, pItem.Tag, "CONST:", vbTextCompare)
                                        If pSpacePos > 0 Then
                                            Text = .SelText
                                            pSpacePos = CLng(Mid(pItem.Tag, pSpacePos + 6))
                                            If InStr(1, Text, pItem.Text) > 0 Then
                                                Text = Replace(Text, pItem.Text, pSpacePos & " '" & pItem.Text)
                                            Else
                                                Text = Text & pSpacePos & " '" & pItem.Text
                                            End If
                                        End If
                                    End If
                                End If
                                
                                If Len(Text) > 0 Then
                                   .SelRTF = ColorIt(Text, pLenDelta)
                                End If
                                .SelStart = StartPos + Len(Text) '- 1
                                .SelLength = OLen
                                m_LoadingScript = False
                                 LastStart = .SelStart
                                 Call LockWindowUpdate(0)
                                 
                                End With
                        'rtEditor.SetFocus
                    End If
                    
                If Not m_LastSearchedItem Is Nothing Then m_LastSearchedItem.Bold = False
            End If
            Set m_LastSearchedItem = Nothing
        End If
        If Not m_ActiveListViewControl Is Nothing Then
            m_ActiveListViewControl.Visible = False
        End If
            
        Set m_ActiveListViewControl = Nothing
    End If
    m_OperatorType = otUnknown
End Sub


    

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_ActiveListViewControl Is Nothing Then
        m_ActiveListViewControl.ZOrder 0
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Start subclassing for WM_SETFOCUS if runtime:
   InitTabTrap True

End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With rtEditor
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height - SB1.Height - 50
        
    End With
    With Line1
        .X1 = 0
        .Y1 = rtEditor.Height
        .X2 = UserControl.Width
        .Y2 = .Y1
    End With
    'With fraMessage
    '    .Left = (rtEditor.Width - .Width) / 2
    '    .Top = (rtEditor.Height - .Height) / 2
    'End With
End Sub

Private Sub UserControl_Terminate()
   m_Intellisense.Unload = True
   Set m_Intellisense = Nothing
   If Not m_frmRefrences Is Nothing Then
        Unload m_frmRefrences
   End If
   Set m_frmRefrences = Nothing
   If Not ctrlTips Is Nothing Then
    Unload ctrlTips
   End If
   Set ctrlTips = Nothing
End Sub
