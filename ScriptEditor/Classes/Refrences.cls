VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Refrences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Refrence"
Attribute VB_Ext_KEY = "Member0" ,"Refrence"
Option Explicit
Implements VbScriptEditor.PersistPropertyBag
'local variable to hold collection
Private mCol As Collection

Public Function Add(ByVal GUID As String, ByVal Verion As String, ByVal BinaryPath As String, ByVal Name As String, Optional ByRef sKey As String = "") As Refrence
    On Error GoTo ErrorTrap
    'create a new object
    Dim objNewMember As Refrence
    Set objNewMember = New Refrence
    'set the properties passed into the method
    With objNewMember
        .GUID = GUID
        .Verion = Verion
        .BinaryPath = BinaryPath
        .Name = Name
    End With
    
    If Len(sKey) = 0 Then
        sKey = GUID & "#" & Verion
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

Public Property Get Item(vntIndexKey As Variant) As Refrence
Attribute Item.VB_UserMemId = 0
    On Error GoTo ErrorTrap
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
  Exit Property
ErrorTrap:
   ' Err.Raise Err.Number, Err.Source, Err.Description
End Property

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
    Dim pRefrence As Refrence
    Dim pCounter As Single
    Dim pMaxRefCount As Single
    pMaxRefCount = PropBag.ReadProperty("RefrenceCount", 0)
    For pCounter = 1 To pMaxRefCount
        Set pRefrence = New Refrence
        Call PropBag.RestoreState(pRefrence, "Refrence" & pCounter)
        mCol.Add pRefrence, pRefrence.Key
    Next
End Sub

Private Sub PersistPropertyBag_Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Dim pRefrence As Refrence
    Dim pCounter As Single
    Call PropBag.WriteProperty("RefrenceCount", mCol.Count, 0)
    pCounter = 1
    For Each pRefrence In mCol
        Call PropBag.SaveState(pRefrence, "Refrence" & pCounter)
        pCounter = pCounter + 1
    Next
End Sub
