VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

Public sdPlayer     As Byte '轮到谁走，0=红方，1=黑方
Public vlWhite      As Long '红方的子力价值
Public vlBlack      As Long '黑方的子力价值
Public nDistance    As Long '距离根节点的步数
Public nMoveNum     As Long '历史走法数
Public zobr         As New ZobristStruct
Public cucpcStartup As Variant
Public sMoveSymbolDesc As String
Private mvsList(0 To MAX_MOVES - 1) As New MoveStruct '历史走法信息列表
Private m_ucpcSquares(0 To 255)     As Byte    '棋盘上的棋子

'初始化棋盘
Public Sub Startup()
  Dim sq As Long
  Dim pc As Long
  
  sdPlayer = 0
  vlWhite = 0
  vlBlack = 0
  nDistance = 0
  sMoveSymbolDesc = ""
  
  '棋盘初始设置
  cucpcStartup = Array( _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 20, 19, 18, 17, 16, 17, 18, 19, 20, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 21, 0, 0, 0, 0, 0, 21, 0, 0, 0, 0, 0, _
        0, 0, 0, 22, 0, 22, 0, 22, 0, 22, 0, 22, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 14, 0, 14, 0, 14, 0, 14, 0, 14, 0, 0, 0, 0, _
        0, 0, 0, 0, 13, 0, 0, 0, 0, 0, 13, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 12, 11, 10, 9, 8, 9, 10, 11, 12, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
  
  zobr.InitZero
  
  For sq = 0 To 255
    m_ucpcSquares(sq) = cucpcStartup(sq)
    pc = cucpcStartup(sq)
    If pc <> 0 Then
      AddPiece sq, pc
    End If
  Next sq
  
  SetIrrev
End Sub

Public Sub ChangeSide()
  '交换走子方
  sdPlayer = 1 - sdPlayer
  zobr.Xor1 Zobrist.Player
End Sub

Public Sub AddPiece(ByVal sq As Byte, ByVal pc As Byte)
  '在棋盘上放一枚棋子
  ucpcSquares(sq) = pc
  '红方加分，黑方(注意"cucvlPiecePos"取值要颠倒)减分
  If (pc < 16) Then
    vlWhite = vlWhite + (cucvlPiecePos(pc - 8)(sq))
    zobr.Xor1 Zobrist.Table(pc - 8, sq)
  Else
    vlBlack = vlBlack + (cucvlPiecePos(pc - 16)(SQUARE_FLIP(sq)))
    zobr.Xor1 Zobrist.Table(pc - 9, sq)
  End If
'  Debug.Print "AddPiece", "("; sq \ 16 - 2; ", "; (sq And &HF) - 2; ")", Pc2Name(pc)
End Sub

Public Sub DelPiece(ByVal sq As Byte, ByVal pc As Byte)
  '从棋盘上拿走一枚棋子
  ucpcSquares(sq) = 0
  '红方减分，黑方(注意"cucvlPiecePos"取值要颠倒)加分
  If (pc < 16) Then
    vlWhite = vlWhite - cucvlPiecePos(pc - 8)(sq)
    zobr.Xor1 Zobrist.Table(pc - 8, sq)
  Else
    vlBlack = vlBlack - cucvlPiecePos(pc - 16)(SQUARE_FLIP(sq))
    zobr.Xor1 Zobrist.Table(pc - 9, sq)
  End If
'  Debug.Print "DelPiece", "("; sq \ 16 - 2; ", "; (sq And &HF) - 2; ")", Pc2Name(pc)
End Sub

'搬一步棋的棋子
Public Function MovePiece(mv As Integer) As Integer
  Dim sqSrc As Long
  Dim sqDst As Long
  Dim pc As Integer
  Dim pcCaptured As Integer

  sqSrc = SRC(mv)
  sqDst = DST(mv)
  pcCaptured = ucpcSquares(sqDst)
  
  If pcCaptured <> 0 Then
    DelPiece sqDst, pcCaptured
  End If
  
  pc = ucpcSquares(sqSrc)
  DelPiece sqSrc, pc
  AddPiece sqDst, pc
  MovePiece = pcCaptured
End Function

'撤消搬一步棋的棋子
Public Sub UndoMovePiece(mv As Integer, pcCaptured As Long)
  Dim sqSrc As Integer
  Dim sqDst As Integer
  Dim pc As Integer
  
  sqSrc = SRC(mv)
  sqDst = DST(mv)
  pc = ucpcSquares(sqDst)
  DelPiece sqDst, pc
  AddPiece sqSrc, pc
  If pcCaptured <> 0 Then
    AddPiece sqDst, pcCaptured
  End If
End Sub

'走一步棋
Public Function MakeMove(mv As Integer) As Boolean
  Dim pcCaptured As Long
  Dim dwKey As Long
  
  dwKey = zobr.dwKey
  pcCaptured = MovePiece(mv)
  
  If Checked() Then
    UndoMovePiece mv, pcCaptured
    MakeMove = False
  Else
    ChangeSide
    mvsList(nMoveNum).Set_ mv, pcCaptured, Checked(), dwKey
    nMoveNum = nMoveNum + 1
    nDistance = nDistance + 1
    MakeMove = True
  End If
End Function

'撤消走一步棋
Public Sub UndoMakeMove()
  nDistance = nDistance - 1
  nMoveNum = nMoveNum - 1
  ChangeSide
  UndoMovePiece mvsList(nMoveNum).wmv, mvsList(nMoveNum).ucpcCaptured
End Sub
'判断走法是否合理
Public Function LegalMove(mv As Integer) As Boolean
  Dim sqSrc As Integer
  Dim sqDst As Integer
  Dim sqPin As Integer
  Dim pcSelfSide As Integer
  Dim pcSrc As Integer
  Dim pcDst As Integer
  Dim nDelta As Integer
  
  '判断走法是否合法，需要经过以下的判断过程：

  '1. 判断起始格是否有自己的棋子
  sqSrc = SRC(mv)
  pcSrc = ucpcSquares(sqSrc)
  pcSelfSide = SIDE_TAG(sdPlayer)
  If ((pcSrc And pcSelfSide) = 0) Then
    LegalMove = False
    Exit Function
  End If

  '2. 判断目标格是否有自己的棋子
  sqDst = DST(mv)
  pcDst = ucpcSquares(sqDst)
  If ((pcDst And pcSelfSide) <> 0) Then
    LegalMove = False
    Exit Function
  End If

  '3. 根据棋子的类型检查走法是否合理
  Select Case (pcSrc - pcSelfSide)
    Case PIECE_KING
      LegalMove = IN_FORT(sqDst) And KING_SPAN(sqSrc, sqDst)
    Case PIECE_ADVISOR
      LegalMove = IN_FORT(sqDst) And ADVISOR_SPAN(sqSrc, sqDst)
    Case PIECE_BISHOP
      LegalMove = SAME_HALF(sqSrc, sqDst) And BISHOP_SPAN(sqSrc, sqDst) And ucpcSquares(BISHOP_PIN(sqSrc, sqDst)) = 0
    Case PIECE_KNIGHT
      sqPin = KNIGHT_PIN(sqSrc, sqDst)
      LegalMove = sqPin <> sqSrc And ucpcSquares(sqPin) = 0
    Case PIECE_ROOK, PIECE_CANNON
      If (SAME_RANK(sqSrc, sqDst)) Then
        nDelta = IIf(sqDst < sqSrc, -1, 1)
      ElseIf (SAME_FILE(sqSrc, sqDst)) Then
        nDelta = IIf(sqDst < sqSrc, -16, 16)
      Else
        LegalMove = False
        Exit Function
      End If
      sqPin = sqSrc + nDelta
      Do While (sqPin <> sqDst And ucpcSquares(sqPin) = 0)
        sqPin = sqPin + nDelta
      Loop
      If (sqPin = sqDst) Then
        LegalMove = ((pcDst = 0) Or (pcSrc - pcSelfSide = PIECE_ROOK))
      ElseIf (pcDst <> 0 And pcSrc - pcSelfSide = PIECE_CANNON) Then
        sqPin = sqPin + nDelta
        Do While (sqPin <> sqDst And ucpcSquares(sqPin) = 0)
          sqPin = sqPin + nDelta
        Loop
        LegalMove = (sqPin = sqDst)
      Else
        LegalMove = False
      End If
    Case PIECE_PAWN
      If (AWAY_HALF(sqDst, sdPlayer) And (sqDst = sqSrc - 1 Or sqDst = sqSrc + 1)) Then
        LegalMove = True
      Else
        LegalMove = (sqDst = SQUARE_FORWARD(sqSrc, sdPlayer))
      End If
    Case Else
      LegalMove = False
  End Select
End Function

'判断是否被将军
Public Function Checked() As Boolean
  Dim i As Integer
  Dim j As Integer
  Dim sqSrc As Integer
  Dim sqDst As Integer
  Dim pcSelfSide As Integer
  Dim pcOppSide As Integer
  Dim pcDst As Integer
  Dim nDelta As Integer
  
  
  pcSelfSide = SIDE_TAG(sdPlayer)
  pcOppSide = OPP_SIDE_TAG(sdPlayer)
  '找到棋盘上的帅(将)，再做以下判断：

  For sqSrc = 0 To 255
    If (ucpcSquares(sqSrc) <> pcSelfSide + PIECE_KING) Then
      GoTo continue
    End If

    '1. 判断是否被对方的兵(卒)将军
    If (ucpcSquares(SQUARE_FORWARD(sqSrc, sdPlayer)) = pcOppSide + PIECE_PAWN) Then
      Checked = True
      Exit Function
    End If
    For nDelta = -1 To 1 Step 2
      If (ucpcSquares(sqSrc + nDelta) = pcOppSide + PIECE_PAWN) Then
        Checked = True
        Exit Function
      End If
    Next nDelta

    '2. 判断是否被对方的马将军(以仕(士)的步长当作马腿)
    For i = 0 To 3
      If (ucpcSquares(sqSrc + ccAdvisorDelta(i)) <> 0) Then
        GoTo cont1
      End If
      For j = 0 To 1
        pcDst = ucpcSquares(sqSrc + ccKnightCheckDelta(i)(j))
        If (pcDst = pcOppSide + PIECE_KNIGHT) Then
          Checked = True
          Exit Function
        End If
      Next j
cont1:
    Next i

    '3. 判断是否被对方的车或炮将军(包括将帅对脸)
    For i = 0 To 3
      nDelta = ccKingDelta(i)
      sqDst = sqSrc + nDelta
      Do While (IN_BOARD(sqDst))
        pcDst = ucpcSquares(sqDst)
        If (pcDst <> 0) Then
          If (pcDst = pcOppSide + PIECE_ROOK Or pcDst = pcOppSide + PIECE_KING) Then
            Checked = True
            Exit Function
          End If
          Exit Do
        End If
        sqDst = sqDst + nDelta
      Loop
      sqDst = sqDst + nDelta
      
'      Debug.Assert sqDst <> 103
      Do While (IN_BOARD(sqDst))
        pcDst = ucpcSquares(sqDst)
        If (pcDst <> 0) Then
          If (pcDst = pcOppSide + PIECE_CANNON) Then
            Checked = True
            Exit Function
          End If
          Exit Do
        End If
        sqDst = sqDst + nDelta
      Loop
    Next i
    Checked = False
    Exit Function
continue:
  Next sqSrc
  Checked = False
End Function

'判断是否被杀
Public Function IsMate() As Boolean
  Dim i As Integer
  Dim nGenMoveNum As Integer
  Dim pcCaptured As Long
  Dim mvs(0 To MAX_GEN_MOVES - 1) As Integer

  nGenMoveNum = GenerateMoves(mvs)
  For i = 0 To nGenMoveNum - 1
    pcCaptured = MovePiece(mvs(i))
    If Not Checked() Then
      UndoMovePiece mvs(i), pcCaptured
      IsMate = False
      Exit Function '只要有一种走法能够避开将军，就没被杀
    Else
      UndoMovePiece mvs(i), pcCaptured
    End If
  Next i
  IsMate = True
End Function

Public Property Get ucpcSquares(ByVal id As Byte) As Byte
  ucpcSquares = m_ucpcSquares(id)
End Property
 
Public Property Let ucpcSquares(ByVal id As Byte, ByVal NewValue As Byte)
  m_ucpcSquares(id) = NewValue
End Property

'生成所有走法，如果 bCapture 为 True 则只生成吃子走法
Public Function GenerateMoves(ByRef mvs() As Integer, Optional bCapture As Boolean = False) As Long
  Dim i As Long
  Dim j As Long
  Dim nGenMoves As Long
  Dim nDelta As Long
  Dim sqSrc As Long
  Dim sqDst As Long
  Dim pcSelfSide As Long
  Dim pcOppSide As Long
  Dim pcSrc As Long
  Dim pcDst As Long
  
  
  '生成所有走法，需要经过以下几个步骤：

  nGenMoves = 0
  pcSelfSide = SIDE_TAG(sdPlayer)
  pcOppSide = OPP_SIDE_TAG(sdPlayer)
  For sqSrc = 0 To 255

    '1. 找到一个本方棋子，再做以下判断：
    pcSrc = ucpcSquares(sqSrc)
    If ((pcSrc And pcSelfSide) = 0) Then
      GoTo continue
    End If

    '2. 根据棋子确定走法
    Select Case pcSrc - pcSelfSide
      Case PIECE_KING
        For i = 0 To 3
          sqDst = sqSrc + ccKingDelta(i)
          If (Not IN_FORT(sqDst)) Then
            GoTo cont1
          End If
          
          pcDst = ucpcSquares(sqDst)
          
          If IIf(bCapture, (pcDst And pcOppSide) <> 0, (pcDst And pcSelfSide) = 0) Then
            mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
            nGenMoves = nGenMoves + 1
          End If
cont1:
        Next i
      Case PIECE_ADVISOR
        For i = 0 To 3
          sqDst = sqSrc + ccAdvisorDelta(i)
          If (Not IN_FORT(sqDst)) Then
            GoTo cont2
          End If
          pcDst = ucpcSquares(sqDst)
          If IIf(bCapture, (pcDst And pcOppSide) <> 0, (pcDst And pcSelfSide) = 0) Then
            mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
            nGenMoves = nGenMoves + 1
          End If
cont2:
        Next i
      Case PIECE_BISHOP
        For i = 0 To 3
          sqDst = sqSrc + ccAdvisorDelta(i)
          If Not (IN_BOARD(sqDst) And HOME_HALF(sqDst, sdPlayer) And ucpcSquares(sqDst) = 0) Then
            GoTo cont3
          End If
          sqDst = sqDst + ccAdvisorDelta(i)
          pcDst = ucpcSquares(sqDst)
          If IIf(bCapture, (pcDst And pcOppSide) <> 0, (pcDst And pcSelfSide) = 0) Then
            mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
            nGenMoves = nGenMoves + 1
          End If
cont3:
        Next i
      Case PIECE_KNIGHT
        For i = 0 To 3
          sqDst = sqSrc + ccKingDelta(i)
          If (ucpcSquares(sqDst) <> 0) Then
            GoTo cont4
          End If
          For j = 0 To 1
            sqDst = sqSrc + ccKnightDelta(i)(j)
            If IN_BOARD(sqDst) Then
              pcDst = ucpcSquares(sqDst)
              If IIf(bCapture, (pcDst And pcOppSide) <> 0, (pcDst And pcSelfSide) = 0) Then
                mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
                nGenMoves = nGenMoves + 1
              End If
            End If
          Next j
cont4:
        Next i
      Case PIECE_ROOK
        For i = 0 To 3
          nDelta = ccKingDelta(i)
          sqDst = sqSrc + nDelta
          Do While (IN_BOARD(sqDst))
            pcDst = ucpcSquares(sqDst)
            If (pcDst = 0) Then
              If Not bCapture Then
                mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
                nGenMoves = nGenMoves + 1
              End If
            Else
              If ((pcDst And pcOppSide) <> 0) Then
                mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
                nGenMoves = nGenMoves + 1
              End If
              Exit Do
            End If
            sqDst = sqDst + nDelta
          Loop
        Next i
      Case PIECE_CANNON
        For i = 0 To 3
          nDelta = ccKingDelta(i)
          sqDst = sqSrc + nDelta
          Do While (IN_BOARD(sqDst))
            pcDst = ucpcSquares(sqDst)
            If (pcDst = 0) Then
              If Not bCapture Then
                mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
                nGenMoves = nGenMoves + 1
              End If
            Else
              Exit Do
            End If
            sqDst = sqDst + nDelta
          Loop
          sqDst = sqDst + nDelta
          Do While (IN_BOARD(sqDst))
            pcDst = ucpcSquares(sqDst)
            If (pcDst <> 0) Then
              If ((pcDst And pcOppSide) <> 0) Then
                mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
                nGenMoves = nGenMoves + 1
              End If
              Exit Do
            End If
            sqDst = sqDst + nDelta
          Loop
        Next i
      Case PIECE_PAWN
        sqDst = SQUARE_FORWARD(sqSrc, sdPlayer)
        If (IN_BOARD(sqDst)) Then
          pcDst = ucpcSquares(sqDst)
          If IIf(bCapture, (pcDst And pcOppSide) <> 0, (pcDst And pcSelfSide) = 0) Then
            mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
            nGenMoves = nGenMoves + 1
          End If
        End If
        If (AWAY_HALF(sqSrc, sdPlayer)) Then
          For nDelta = -1 To 1 Step 2
            sqDst = sqSrc + nDelta
            If (IN_BOARD(sqDst)) Then
              pcDst = ucpcSquares(sqDst)
              If IIf(bCapture, (pcDst And pcOppSide) <> 0, (pcDst And pcSelfSide) = 0) Then
                mvs(nGenMoves) = MOVE_(sqSrc, sqDst)
                nGenMoves = nGenMoves + 1
              End If
            End If
          Next nDelta
        End If
    End Select
continue:
  Next sqSrc
  
  GenerateMoves = nGenMoves
End Function

'局面评价函数
Public Function Evaluate() As Long
  Evaluate = IIf(sdPlayer = 0, vlWhite - vlBlack, vlBlack - vlWhite) + ADVANCED_VALUE
End Function

'检测重复局面
Public Function RepStatus(Optional ByVal nRecur As Byte = 1) As Byte
  Dim bSelfSide     As Boolean
  Dim bPerpCheck    As Boolean
  Dim bOppPerpCheck As Boolean
  Dim lpmvs         As Integer
  
  bSelfSide = False
  lpmvs = nMoveNum - 1
  
  '起初 bPerpCheck(本方长将) 和 bOppPerpCheck(对方长将)都设为 True
  bPerpCheck = True
  bOppPerpCheck = True

  Do While (mvsList(lpmvs).wmv <> 0 And mvsList(lpmvs).ucpcCaptured = 0)

    If (bSelfSide) Then '电脑方
      bPerpCheck = bPerpCheck And mvsList(lpmvs).ucbCheck '当一方存在非将走法时就改为 False

      If (mvsList(lpmvs).dwKey = zobr.dwKey) Then 'dwKey 在检查重复局面时用，也作为置换表的键值，dwLock0 和 dwLock1 用作置换表的校验值，另外，dwLock1 还是查找开局库的依据(后面会提到)。
        nRecur = nRecur - 1

        If (nRecur = 0) Then
          RepStatus = 1 + IIf(bPerpCheck, 2, 0) + IIf(bOppPerpCheck, 4, 0)
'     　　返回1(=1+0+0)，表示存在重复局面，但双方都无长将(判和)；
'     　　返回3(=1 + 2)，表示存在重复局面，本方单方面长将(判本方负)；
'     　　返回5(=1 + 4)，表示存在重复局面，对方单方面长将(判对方负)；
'     　　返回7(=1+2+4)，表示存在重复局面，双方长将(判和)。
          Exit Function
        End If
      End If

    Else '人脑方
      bOppPerpCheck = bOppPerpCheck And mvsList(lpmvs).ucbCheck '当一方存在非将走法时就改为 False
    End If

    bSelfSide = Not bSelfSide
    lpmvs = lpmvs - 1 '从后往前检查
  Loop

  RepStatus = 0 '返回0，表示没有重复局面；
End Function

'重复局面分值
Public Function RepValue(ByVal nRepStatus As Byte) As Integer
  Dim vlReturn As Long
  
  vlReturn = IIf((nRepStatus And 2) = 0, 0, nDistance - MATE_VALUE) + IIf((nRepStatus And 4) = 0, 0, MATE_VALUE - nDistance)
  RepValue = IIf(vlReturn = 0, DrawValue(), vlReturn)
End Function

'是否被将军
Public Function InCheck() As Boolean
  InCheck = mvsList(nMoveNum - 1).ucbCheck
End Function
 
'判断是否允许空步裁剪
Public Function NullOkay() As Boolean
  NullOkay = IIf(sdPlayer = 0, vlWhite, vlBlack) > NULL_MARGIN
End Function

'走一步空步
Public Sub NullMove()
  Dim dwKey As Long
  
  dwKey = zobr.dwKey
  ChangeSide
  mvsList(nMoveNum).Set_ 0, 0, False, dwKey
  nMoveNum = nMoveNum + 1
  nDistance = nDistance + 1
End Sub

'撤消走一步空步
Public Sub UndoNullMove()
  nDistance = nDistance - 1
  nMoveNum = nMoveNum - 1
  ChangeSide
End Sub

'上一步是否吃子
Public Function Captured() As Boolean
  Captured = mvsList(nMoveNum - 1).ucpcCaptured <> 0
End Function

'清空(初始化)历史走法信息
Public Sub SetIrrev()
  mvsList(0).Set_ 0, 0, Checked(), zobr.dwKey
  nMoveNum = 1
End Sub

Private Function DrawValue() As Long
  DrawValue = IIf((nDistance And 1) = 0, -DRAW_VALUE, DRAW_VALUE)
End Function

'清空棋盘
Public Sub ClearBoard()
  sdPlayer = 0
  vlWhite = 0
  vlBlack = 0
  nDistance = 0
'  memset ucpcSquares, 0, 256
  zobr.InitZero
End Sub

Public Sub Mirror(ByRef posMirror As PositionStruct)
  Dim sq As Integer
  Dim pc As Long

  posMirror.ClearBoard
  
  For sq = 0 To 255
    pc = ucpcSquares(sq)
    If (pc <> 0) Then
      posMirror.AddPiece MIRROR_SQUARE(sq), pc
    End If
  Next sq
  
  If (sdPlayer = 1) Then
    posMirror.ChangeSide
  End If
  
  posMirror.SetIrrev
End Sub

