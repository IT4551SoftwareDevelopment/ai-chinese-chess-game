VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'置换表走法和两个杀手走法
Public mvHash As Integer
Public mvKiller1 As Integer
Public mvKiller2 As Integer

'当前阶段，当前采用第几个走法，总共有几个走法
Public nPhase As Long
Public nIndex As Long
Public nGenMoves As Byte

'所有的走法
Private m_mvs(0 To MAX_GEN_MOVES - 1) As Integer
  
Public Property Get mvs(ByVal id As Long) As Integer
  mvs = m_mvs(id)
End Property

Public Property Let mvs(ByVal id As Long, ByVal NewValue As Integer)
  m_mvs(id) = NewValue
End Property

'初始化，设定置换表走法和两个杀手走法
Public Sub Init(mvHash_ As Integer)
  If frmSettings.chkKiller.Value = vbChecked Then
    mvHash = mvHash_
    mvKiller1 = Search.mvKillers(pos.nDistance, 0)
    mvKiller2 = Search.mvKillers(pos.nDistance, 1)
    nPhase = PHASE_HASH
  Else
    nPhase = PHASE_GEN_MOVES
  End If
End Sub

'得到下一个走法
Public Function Next_() As Integer
  Dim mv As Integer

  '"nPhase"表示着法启发的若干阶段，依次为：
  
  If frmSettings.chkKiller.Value = vbChecked Then
    ' 0.置换表着法启发，完成后立即进入下一阶段；
    If nPhase = PHASE_HASH Then
      nPhase = PHASE_KILLER_1
      If (mvHash <> 0) Then
        Next_ = mvHash
        Exit Function
      End If
    End If
    
  ' 1.杀手着法启发(第一个杀手着法)，完成后立即进入下一阶段；
    If nPhase = PHASE_KILLER_1 Then
      nPhase = PHASE_KILLER_2
  
      If (mvKiller1 <> mvHash And mvKiller1 <> 0 And pos.LegalMove(mvKiller1)) Then
        Next_ = mvKiller1
        Exit Function
      End If
    End If
    
  ' 2.杀手着法启发(第二个杀手着法)，完成后立即进入下一阶段；
    If nPhase = PHASE_KILLER_2 Then
      nPhase = PHASE_GEN_MOVES
  
      If (mvKiller2 <> mvHash And mvKiller2 <> 0 And pos.LegalMove(mvKiller2)) Then
        Next_ = mvKiller2
        Exit Function
      End If
    End If
  End If
  
  '3.生成所有着法，完成后立即进入下一阶段；
  If nPhase = PHASE_GEN_MOVES Then
    nPhase = PHASE_REST
    nGenMoves = pos.GenerateMoves(m_mvs)
    If frmSettings.chkSortMoves.Value = vbChecked Then
      If frmSettings.optQuickSort.Value Then
        qsort m_mvs, 0, nGenMoves - 1, "History"
      Else
        SelectSort m_mvs, nGenMoves - 1, "History"
      End If
    End If
    nIndex = 0
  End If
  
  '4.对剩余着法做历史表启发；
  If nPhase = PHASE_REST Then
    Do While (nIndex < nGenMoves)
      mv = mvs(nIndex)
      nIndex = nIndex + 1

      If (mv <> mvHash And mv <> mvKiller1 And mv <> mvKiller2) Or frmSettings.chkKiller.Value = vbUnchecked Then
        Next_ = mv
        Exit Function
      End If
    Loop
  End If
  
  '5.没有着法了，返回零。
  Next_ = 0
End Function
