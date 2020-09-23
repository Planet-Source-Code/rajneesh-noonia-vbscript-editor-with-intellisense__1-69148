VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Refrence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Implements PersistPropertyBag

Public GUID As String
Public Verion As String
Public BinaryPath As String
Public Name As String
Public Key As String

Private Sub PersistPropertyBag_Load(ByVal PropBag As VbScriptEditor.cPropertyBag)
    GUID = PropBag.ReadProperty("GUID", "")
    Verion = PropBag.ReadProperty("Verion", "")
    BinaryPath = PropBag.ReadProperty("BinaryPath", "")
    Name = PropBag.ReadProperty("Name", "")
    Key = PropBag.ReadProperty("Key", "")
End Sub

Private Sub PersistPropertyBag_Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Call PropBag.WriteProperty("GUID", GUID, "")
    Call PropBag.WriteProperty("Verion", Verion, "")
    Call PropBag.WriteProperty("BinaryPath", BinaryPath, "")
    Call PropBag.WriteProperty("Name", Name, "")
    Call PropBag.WriteProperty("Key", Key, "")
End Sub
