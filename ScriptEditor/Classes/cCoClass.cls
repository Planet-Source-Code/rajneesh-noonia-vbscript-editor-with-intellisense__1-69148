VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum iTypeKind
    iLibrary
    iModule
    iClass
    iFunction
    iType
    iEnum
    iConstant
    iProperty
    iEvent
End Enum

Public ProgID As String
Public TypeKind As iTypeKind
Public LibProgID As String
