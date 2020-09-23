VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Modules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cModule"
Attribute VB_Ext_KEY = "Member0" ,"cModule"
Option Explicit
Implements VbScriptEditor.PersistPropertyBag
'local variable to hold collection
Private mCol As Collection

Public Function Add(ByVal Name As String, Optional ByRef sKey As String) As Module
    'create a new object
    Dim objNewMember As Module
    On Error GoTo ErrorTrap
    Set objNewMember = New Module
    'set the properties passed into the method
    With objNewMember
        .Name = Name
        Set .Variables = New Variables
    End With
    
    If Len(sKey) = 0 Then
        sKey = Name
    End If
    objNewMember.Key = sKey
    mCol.Add objNewMember, sKey
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
ErrorTrap:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Property Get Item(vntIndexKey As Variant) As Module
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Sub Clear()
    Dim pModule As Module
    For Each pModule In mCol
        mCol.Remove pModule.Key
    Next
    Set mCol = Nothing
    Set mCol = New Collection
End Sub


Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Private Sub PersistPropertyBag_Load(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Dim pModule As Module
    Dim pCounter As Single
    Dim pMaxModCount As Single
    On Error Resume Next
    pMaxModCount = PropBag.ReadProperty("ModuleCount", 0)
    For pCounter = 1 To pMaxModCount
        Set pModule = New Module
        Call PropBag.RestoreState(pModule, "Module" & pCounter)
        mCol.Add pModule, pModule.Key
    Next
End Sub

Private Sub PersistPropertyBag_Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Dim pModule As Module
    Dim pCounter As Single
    Call PropBag.WriteProperty("ModuleCount", mCol.Count)
    pCounter = 1
    For Each pModule In mCol
        Call PropBag.SaveState(pModule, "Module" & pCounter)
        pCounter = pCounter + 1
    Next
End Sub
