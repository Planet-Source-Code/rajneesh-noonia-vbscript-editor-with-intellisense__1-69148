VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Variable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Implements VbScriptEditor.PersistPropertyBag
Public Name As String
Public ObjProgID As String
Public LibRefrence As Refrence
Public Key As String

Private Sub PersistPropertyBag_Load(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Set LibRefrence = New Refrence
    Name = PropBag.ReadProperty("GUID", "")
    ObjProgID = PropBag.ReadProperty("ObjProgID", "")
    Key = PropBag.ReadProperty("Key", "")
    Call PropBag.RestoreState(LibRefrence, "LibRefrence")
    
End Sub

Private Sub PersistPropertyBag_Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
    
    Call PropBag.WriteProperty("Name", Name, "")
    Call PropBag.WriteProperty("Key", Key, "")
    Call PropBag.WriteProperty("ObjProgID", ObjProgID, "")
    Call PropBag.SaveState(LibRefrence, "LibRefrence")
End Sub
