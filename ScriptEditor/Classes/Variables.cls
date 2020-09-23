VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Variables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Variable"
Attribute VB_Ext_KEY = "Member0" ,"Variable"
Option Explicit
Implements VbScriptEditor.PersistPropertyBag

'local variable to hold collection
Private mCol As Collection



Public Function Add(ByVal Name As String, ByVal ObjProgID As String, ByVal LibRefrence As Refrence, Optional ByRef sKey As String = "") As Variable
    'create a new object
    
    Dim objNewMember As Variable
    On Error Resume Next
    Set objNewMember = New Variable
    

    'set the properties passed into the method
    With objNewMember
        .Name = Name
        If Not LibRefrence Is Nothing Then
        .ObjProgID = ObjProgID
        Else
            .ObjProgID = ""
        End If
        
        Set .LibRefrence = LibRefrence
    End With
    
    If Len(sKey) = 0 Then
        sKey = Name
    End If
    objNewMember.Key = sKey
    
    mCol.Add objNewMember, sKey
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Variable
Attribute Item.VB_UserMemId = 0
    On Error GoTo ErrorTrap
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
  Exit Property
ErrorTrap:
  Set Item = Nothing
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
    Dim pVariable As Variable
    Dim pCounter As Single
    Dim pMaxVarCount As Single
    pMaxVarCount = PropBag.ReadProperty("VariableCount", 0)
    For pCounter = 1 To pMaxVarCount
        Set pVariable = New Variable
        Call PropBag.RestoreState(pVariable, "Variable" & pCounter)
        mCol.Add pVariable, pVariable.Key
    Next
End Sub

Private Sub PersistPropertyBag_Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Dim pVariable As Variable
    Dim pCounter As Single
    Call PropBag.WriteProperty("VariableCount", mCol.Count)
    pCounter = 1
    For Each pVariable In mCol
        Call PropBag.SaveState(pVariable, "Variable" & pCounter)
        pCounter = pCounter + 1
    Next
End Sub
