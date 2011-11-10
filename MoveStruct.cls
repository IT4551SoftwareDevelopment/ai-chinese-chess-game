VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoveStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

Public wmv           As Integer
Public ucpcCaptured  As Byte
Public ucbCheck      As Byte
Public dwKey         As Long

Public Sub Set_(mv As Integer, pcCaptured As Long, bCheck As Boolean, dwKey_ As Long)
  wmv = mv
  ucpcCaptured = pcCaptured
  ucbCheck = bCheck
  dwKey = dwKey_
End Sub
