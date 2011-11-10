VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RC4Struct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

Private s(0 To 255) As Byte
Private x As Long
Private y As Long

Private Sub XChg(ByRef si As Byte, ByRef sj As Byte)
  Dim uc As Byte

  uc = si
  si = sj
  sj = uc
End Sub

'用空密钥初始化密码流生成器
Public Sub InitZero()
 Dim i As Long
 Dim j As Long
 
 x = 0
 y = 0
 j = 0
 
 For i = 0 To 255
   s(i) = i
 Next i
 
 For i = 0 To 255
   j = (j + s(i)) And &HFF
   XChg s(i), s(j)
 Next i
End Sub

'生成密码流的下一个字节
Private Function NextByte() As Byte
  x = (x + 1) And 255
  y = (y + s(x)) And 255
  XChg s(x), s(y)
  NextByte = s((s(x) + CInt(s(y))) And 255)
End Function

'生成密码流的下四个字节
Public Function NextLong() As Long
  Dim uc(0 To 3) As Byte
  Dim i As Byte

  For i = 0 To 3
    uc(i) = NextByte()
  Next i

  CopyMemory NextLong, uc(0), 4
End Function
