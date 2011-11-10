VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZobristStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

Public dwKey As Long
Public dwLock0 As Long
Public dwLock1 As Long

'”√¡„ÃÓ≥‰Zobrist
Public Sub InitZero()
  dwKey = 0
  dwLock0 = 0
  dwLock1 = 0
End Sub

'”√√‹¬Î¡˜ÃÓ≥‰Zobrist
Public Sub InitRC4(ByRef rc4 As RC4Struct)
  dwKey = rc4.NextLong
  dwLock0 = rc4.NextLong
  dwLock1 = rc4.NextLong
End Sub

Public Sub Xor1(ByRef zobr As ZobristStruct)
  dwKey = dwKey Xor zobr.dwKey
  dwLock0 = dwLock0 Xor zobr.dwLock0
  dwLock1 = dwLock1 Xor zobr.dwLock1
End Sub

Public Sub Xor2(ByRef zobr1 As ZobristStruct, ByRef zobr2 As ZobristStruct)
  dwKey = zobr1.dwKey Xor zobr2.dwKey
  dwLock0 = zobr1.dwLock0 Xor zobr2.dwLock0
  dwLock1 = zobr1.dwLock1 Xor zobr2.dwLock1
End Sub
