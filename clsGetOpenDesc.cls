VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetOpenDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SysAllocStringByteLen Lib "OLEAUT32.DLL" (ByVal lpsz As Long, ByVal dwLen As Long) As String
Private Declare Function lstrlenA Lib "KERNEL32.DLL" (ByVal lpsz As Long) As Long

Private Declare Sub EccoInitOpenVar Lib "ECCO.DLL" Alias "_EccoInitOpenVar@4" (Optional ByVal bTraditional As Long = 0)
Private Declare Function EccoIndex Lib "ECCO.DLL" Alias "_EccoIndex@4" (ByVal sDesc As String) As Long
Private Declare Function EccoOpening Lib "ECCO.DLL" Alias "_EccoOpening@4" (ByVal dwEccoIndex As Long) As Long
Private Declare Function EccoVariation Lib "ECCO.DLL" Alias "_EccoVariation@4" (ByVal dwEccoIndex As Long) As Long

Private Sub Class_Initialize()
  EccoInitOpenVar
End Sub

Public Function GetOpen(ByVal sMove As String) As String
  GetOpen = AllocString(EccoOpening(EccoIndex(sMove)))
End Function

Public Function GetVar(ByVal sMove As String) As String
  GetVar = AllocString(EccoVariation(EccoIndex(sMove)))
End Function

Public Function GetEccoNo(ByVal sMove As String) As String
  GetEccoNo = Left(MkL(EccoIndex(sMove)), 3)
End Function

Private Function AllocString(ByVal lpsz As Long) As String
  AllocString = SysAllocStringByteLen(lpsz, lstrlenA(lpsz))
End Function

Private Function MkL(ByVal dw As Long) As String
  MkL = SysAllocStringByteLen(VarPtr(dw), 4)
End Function

