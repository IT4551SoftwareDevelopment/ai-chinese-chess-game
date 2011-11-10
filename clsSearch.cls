VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

'与搜索有关的全局变量
Public mvResult As Long                    '电脑走的棋
Public nBookSize As Integer                '开局库大小
Private mHistoryTable(0 To 65535) As Long  '历史表
Private m_mvKillers(0 To LIMIT_DEPTH - 1, 0 To 1) As Long '杀手走法表
  
Public Property Get nHistoryTable(ByVal id As Long) As Long
  If id < 0 Then
    id = id + 65536
  End If
  
  nHistoryTable = mHistoryTable(id)
End Property

Public Property Let nHistoryTable(ByVal id As Long, ByVal NewValue As Long)
  If id < 0 Then
    id = id + 65536
  End If
  
  mHistoryTable(id) = NewValue
End Property

Public Property Get mvKillers(ByVal id As Long, id2 As Byte) As Long
  mvKillers = m_mvKillers(id, id2)
End Property

Public Property Let mvKillers(ByVal id As Long, id2 As Byte, ByVal NewValue As Long)
  m_mvKillers(id, id2) = NewValue
End Property

