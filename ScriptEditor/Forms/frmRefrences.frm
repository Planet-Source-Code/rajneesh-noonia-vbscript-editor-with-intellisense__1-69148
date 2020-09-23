VERSION 5.00
Begin VB.Form frmRefrences 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Refrences - VB Script Editor"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTypeInfo 
      Caption         =   "TypeInfo"
      Height          =   945
      Left            =   90
      TabIndex        =   5
      Top             =   4440
      Width           =   5685
      Begin VB.TextBox txtLocation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Location"
         Top             =   390
         Width           =   4515
      End
      Begin VB.Label Label2 
         Caption         =   "Location:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.ListBox lstTypeLibs 
      Height          =   3660
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   540
      Width           =   4305
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4590
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   4590
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4590
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Available Refrences:"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   2235
   End
End
Attribute VB_Name = "frmRefrences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long

Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2

Public m_Editor As Editor
Private m_blnIsLoading As Boolean

Private Type ListExItem
     Name As String
     ItemData As Long
End Type

Private Function GetListItemIndex(ByVal pItemToSearch As String) As Long
   GetListItemIndex = SendMessage(lstTypeLibs.hwnd, LB_FINDSTRING, -1, ByVal (pItemToSearch))
End Function


Private Sub CheckAddedRefrences()
    Dim pRefrence As Refrence
    Dim pListIndex As Single
    Dim pNextIndex As Single
    Dim pTypeLibInfo As cTypeLibInfo
    Dim lPtr As Long
    Dim pItem As String
    Dim pItemData  As Long
    For Each pRefrence In g_Project.Refrences
        pListIndex = GetListItemIndex(pRefrence.Name)
        Do
            If pListIndex > -1 Then
                lPtr = lstTypeLibs.ItemData(pListIndex)
                If Not (lPtr = 0) Then
                    Set pTypeLibInfo = ObjectFromPtr(lPtr)
                End If
            Else
                Set pTypeLibInfo = New cTypeLibInfo
            End If
            
            If pTypeLibInfo.Path = pRefrence.BinaryPath Then
                pItem = lstTypeLibs.List(pListIndex)
                pItemData = lstTypeLibs.ItemData(pListIndex)
                lstTypeLibs.RemoveItem (pListIndex)
                lstTypeLibs.AddItem pItem, 0
                lstTypeLibs.ItemData(0) = pItemData
                lstTypeLibs.Selected(0) = True
            End If
            
            pListIndex = pListIndex + 1
            If pListIndex > lstTypeLibs.ListCount Then
                Exit Do
            End If
            If pTypeLibInfo.Name <> pRefrence.Name Then
                Exit Do
            End If
            
        Loop While Not (pTypeLibInfo.Path = pRefrence.BinaryPath)
    Next
End Sub

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As cTypeLibInfo
Dim objT As cTypeLibInfo
   ' Bruce McKinney's code for getting an Object from the
   ' object pointer:
   CopyMemory objT, lPtr, 4
   Set ObjectFromPtr = objT
   CopyMemory objT, 0&, 4
End Property

Public Sub Populate()
    Dim iSectCount As Long, iSect As Long, sSections() As String
    Dim iVerCount As Long, iVer As Long, sVersions() As String
    Dim iExeSectCount As Long, sExeSect() As String
    Dim pItemArray() As ListExItem
    Dim iExeSect As Long
    Dim bFoundExeSect As Boolean
    Dim sExists As String
    Dim cTLI As cTypeLibInfo
    Dim i As olelib.IUnknown
    Dim pArrayIndex As Long
       pClearList
       lstTypeLibs.Clear
       lstTypeLibs.Visible = False
       pArrayIndex = 0
       Dim cR As New cRegistry
       cR.ClassKey = HKEY_CLASSES_ROOT
       cR.ValueType = REG_SZ
       cR.SectionKey = "TypeLib"
       ' Get the registered Type Libs:
       If cR.EnumerateSections(sSections(), iSectCount) Then
          For iSect = 1 To iSectCount
             ' Enumerate the versions for each typelib:
             cR.SectionKey = "TypeLib\" & sSections(iSect)
             If cR.EnumerateSections(sVersions(), iVerCount) Then
                For iVer = 1 To iVerCount
                   Set cTLI = New cTypeLibInfo
                   cTLI.CLSID = sSections(iSect)
                   cTLI.Ver = sVersions(iVer)
                   cR.SectionKey = "TypeLib\" & sSections(iSect) & "\" & sVersions(iVer)
                   cTLI.Name = cR.Value
                   cR.EnumerateSections sExeSect(), iExeSectCount
                   If iExeSectCount > 0 Then
                      bFoundExeSect = False
                      For iExeSect = 1 To iExeSectCount
                         If IsNumeric(sExeSect(iExeSect)) Then
                            cR.SectionKey = cR.SectionKey & "\" & sExeSect(iExeSect) & "\win32"
                            bFoundExeSect = True
                            Exit For
                         End If
                      Next iExeSect
                      If bFoundExeSect Then
                         cTLI.Path = cR.Value
                         If FileExists(cTLI.Path) Then
                            sExists = "Y"
                         Else
                            sExists = "N"
                         End If
                      Else
                         sExists = "N"
                      End If
                   Else
                      sExists = "N"
                   End If
                   cTLI.Exists = (StrComp(sExists, "Y") = 0)
                   If Len(cTLI.Name) > 0 And cTLI.Exists Then
                        If IsValidLib(cTLI) Then
                            ReDim Preserve pItemArray(pArrayIndex)
                            pItemArray(pArrayIndex).Name = cTLI.Name
                            pItemArray(pArrayIndex).ItemData = ObjPtr(cTLI)
                            Set i = cTLI
                            i.AddRef
                            pArrayIndex = pArrayIndex + 1
                        End If
                   End If
                   
                Next iVer
             End If
          Next iSect
       End If
       
       lstTypeLibs.Visible = True
       Call Sort(pItemArray, lstTypeLibs)
End Sub

Private Sub Sort(inpArray() As ListExItem, inpList As ListBox)
   Dim intRet As Long
   Dim intCompare As Long
   Dim intLoopTimes As Long
   Dim strTemp As ListExItem

   For intLoopTimes = 0 To UBound(inpArray)
      For intCompare = LBound(inpArray) To UBound(inpArray) - 1
         intRet = StrComp(inpArray(intCompare).Name, _
         inpArray(intCompare + 1).Name, vbTextCompare)
         If intRet = 1 Then
            ' String1 is greater than String2
            strTemp = inpArray(intCompare)
            inpArray(intCompare) = inpArray(intCompare + 1)
            inpArray(intCompare + 1) = strTemp
         End If
      Next
   Next

   inpList.Clear

   For intCompare = 0 To UBound(inpArray)
      inpList.AddItem inpArray(intCompare).Name
      inpList.ItemData(intCompare) = inpArray(intCompare).ItemData
   Next

End Sub


Private Function IsValidLib(ByVal cTLI As cTypeLibInfo) As Boolean
    Dim pTLIApplication As TLI.TLIApplication
    On Error GoTo ErrorTrap
    Set pTLIApplication = New TLI.TLIApplication
    Call pTLIApplication.TypeLibInfoFromFile(cTLI.Path)
    Set pTLIApplication = Nothing
    IsValidLib = True
Exit Function
ErrorTrap:
    IsValidLib = False
End Function

Private Sub pDeleteEntry(ByVal lPtr As Long)
Dim cTLI As cTypeLibInfo
   
   Set cTLI = ObjectFromPtr(lPtr)
   Dim cR As New cRegistry
   cR.ClassKey = HKEY_CLASSES_ROOT
   cR.SectionKey = "TypeLib\" & cTLI.CLSID & "\" & cTLI.Ver
   
   On Error Resume Next
   If cR.DeleteKey Then
      If Err.Number = 0 Then
         MsgBox "Successfully deleted the item " & cTLI.Name & ", version " & cTLI.Ver, vbInformation
      End If
   End If
   Err.Clear
   On Error GoTo 0
   
End Sub

Private Sub pClearList()
Dim lI As Long
Dim lPtr As Long
Dim i As olelib.IUnknown
Dim cTLI As cTypeLibInfo

   For lI = 0 To lstTypeLibs.ListCount - 1
      lPtr = lstTypeLibs.ItemData(lI)
      If Not (lPtr = 0) Then
         Set cTLI = ObjectFromPtr(lPtr)
         Set i = cTLI
         i.Release
         Set i = Nothing
         Set cTLI = Nothing
      End If
   Next lI

End Sub



Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   m_blnIsLoading = True
   Populate
   
   If lstTypeLibs.ListCount > 0 Then
      lstTypeLibs.ListIndex = 0
      lstTypeLibs_Click
   End If
   m_blnIsLoading = False
   Screen.MousePointer = vbNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   pClearList
End Sub


Private Sub lstTypeLibs_Click()
Dim lI As Long
Dim lPtr As Long
Dim cTLI As cTypeLibInfo
Dim pRefrence As Refrence
   lI = lstTypeLibs.ListIndex
   If lI > -1 Then
      lPtr = lstTypeLibs.ItemData(lI)
      If Not (lPtr = 0) Then
         Set cTLI = ObjectFromPtr(lPtr)
         fraTypeInfo.Caption = cTLI.Name & " (" & cTLI.Ver & ")"
         txtLocation.Text = cTLI.Path
         If Not cTLI.Exists Then
            txtLocation.ForeColor = &HC0&
         Else
            txtLocation.ForeColor = vbWindowText
         End If
         
         If Not m_blnIsLoading Then
            If lstTypeLibs.Selected(lI) = True Then
               On Error Resume Next
               Call g_Project.Refrences.Add(cTLI.CLSID, cTLI.Ver, cTLI.Path, cTLI.Name)
            Else
               On Error Resume Next
               Set pRefrence = g_Project.Refrences.Item(cTLI.CLSID & "#" & cTLI.Ver)
               If Not pRefrence Is Nothing Then
                   g_Project.Refrences.Remove (pRefrence.Key)
                   Set pRefrence = Nothing
               End If
            End If
        End If
      End If
   End If
   
End Sub



Public Sub PrepareUI()
    Call CheckAddedRefrences
End Sub

Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub
