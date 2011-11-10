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

Public sdPlayer     As Byte '�ֵ�˭�ߣ�0=�췽��1=�ڷ�
Public vlWhite      As Long '�췽��������ֵ
Public vlBlack      As Long '�ڷ���������ֵ
Public nDistance    As Long '������ڵ�Ĳ���
Public nMoveNum     As Long '��ʷ�߷���
Public zobr         As New ZobristStruct
Public cucpcStartup As Variant
Public sMoveSymbolDesc As String
Private mvsList(0 To MAX_MOVES - 1) As New MoveStruct '��ʷ�߷���Ϣ�б�
Private m_ucpcSquares(0 To 255)     As Byte    '�����ϵ�����

'��ʼ������
Public Sub Startup()
  Dim sq As Long
  Dim pc As Long
  
  sdPlayer = 0
  vlWhite = 0
  vlBlack = 0
  nDistance = 0
  sMoveSymbolDesc = ""
  
  '���̳�ʼ����
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
  '�������ӷ�
  sdPlayer = 1 - sdPlayer
  zobr.Xor1 Zobrist.Player
End Sub

Public Sub AddPiece(ByVal sq As Byte, ByVal pc As Byte)
  '�������Ϸ�һö����
  ucpcSquares(sq) = pc
  '�췽�ӷ֣��ڷ�(ע��"cucvlPiecePos"ȡֵҪ�ߵ�)����
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
  '������������һö����
  ucpcSquares(sq) = 0
  '�췽���֣��ڷ�(ע��"cucvlPiecePos"ȡֵҪ�ߵ�)�ӷ�
  If (pc < 16) Then
    vlWhite = vlWhite - cucvlPiecePos(pc - 8)(sq)
    zobr.Xor1 Zobrist.Table(pc - 8, sq)
  Else
    vlBlack = vlBlack - cucvlPiecePos(pc - 16)(SQUARE_FLIP(sq))
    zobr.Xor1 Zobrist.Table(pc - 9, sq)
  End If
'  Debug.Print "DelPiece", "("; sq \ 16 - 2; ", "; (sq And &HF) - 2; ")", Pc2Name(pc)
End Sub

'��һ���������
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

'������һ���������
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

'��һ����
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

'������һ����
Public Sub UndoMakeMove()
  nDistance = nDistance - 1
  nMoveNum = nMoveNum - 1
  ChangeSide
  UndoMovePiece mvsList(nMoveNum).wmv, mvsList(nMoveNum).ucpcCaptured
End Sub
'�ж��߷��Ƿ����
Public Function LegalMove(mv As Integer) As Boolean
  Dim sqSrc As Integer
  Dim sqDst As Integer
  Dim sqPin As Integer
  Dim pcSelfSide As Integer
  Dim pcSrc As Integer
  Dim pcDst As Integer
  Dim nDelta As Integer
  
  '�ж��߷��Ƿ�Ϸ�����Ҫ�������µ��жϹ��̣�

  '1. �ж���ʼ���Ƿ����Լ�������
  sqSrc = SRC(mv)
  pcSrc = ucpcSquares(sqSrc)
  pcSelfSide = SIDE_TAG(sdPlayer)
  If ((pcSrc And pcSelfSide) = 0) Then
    LegalMove = False
    Exit Function
  End If

  '2. �ж�Ŀ����Ƿ����Լ�������
  sqDst = DST(mv)
  pcDst = ucpcSquares(sqDst)
  If ((pcDst And pcSelfSide) <> 0) Then
    LegalMove = False
    Exit Function
  End If

  '3. �������ӵ����ͼ���߷��Ƿ����
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

'�ж��Ƿ񱻽���
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
  '�ҵ������ϵ�˧(��)�����������жϣ�

  For sqSrc = 0 To 255
    If (ucpcSquares(sqSrc) <> pcSelfSide + PIECE_KING) Then
      GoTo continue
    End If

    '1. �ж��Ƿ񱻶Է��ı�(��)����
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

    '2. �ж��Ƿ񱻶Է���������(����(ʿ)�Ĳ�����������)
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

    '3. �ж��Ƿ񱻶Է��ĳ����ڽ���(������˧����)
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

'�ж��Ƿ�ɱ
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
      Exit Function 'ֻҪ��һ���߷��ܹ��ܿ���������û��ɱ
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

'���������߷������ bCapture Ϊ True ��ֻ���ɳ����߷�
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
  
  
  '���������߷�����Ҫ�������¼������裺

  nGenMoves = 0
  pcSelfSide = SIDE_TAG(sdPlayer)
  pcOppSide = OPP_SIDE_TAG(sdPlayer)
  For sqSrc = 0 To 255

    '1. �ҵ�һ���������ӣ����������жϣ�
    pcSrc = ucpcSquares(sqSrc)
    If ((pcSrc And pcSelfSide) = 0) Then
      GoTo continue
    End If

    '2. ��������ȷ���߷�
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

'�������ۺ���
Public Function Evaluate() As Long
  Evaluate = IIf(sdPlayer = 0, vlWhite - vlBlack, vlBlack - vlWhite) + ADVANCED_VALUE
End Function

'����ظ�����
Public Function RepStatus(Optional ByVal nRecur As Byte = 1) As Byte
  Dim bSelfSide     As Boolean
  Dim bPerpCheck    As Boolean
  Dim bOppPerpCheck As Boolean
  Dim lpmvs         As Integer
  
  bSelfSide = False
  lpmvs = nMoveNum - 1
  
  '��� bPerpCheck(��������) �� bOppPerpCheck(�Է�����)����Ϊ True
  bPerpCheck = True
  bOppPerpCheck = True

  Do While (mvsList(lpmvs).wmv <> 0 And mvsList(lpmvs).ucpcCaptured = 0)

    If (bSelfSide) Then '���Է�
      bPerpCheck = bPerpCheck And mvsList(lpmvs).ucbCheck '��һ�����ڷǽ��߷�ʱ�͸�Ϊ False

      If (mvsList(lpmvs).dwKey = zobr.dwKey) Then 'dwKey �ڼ���ظ�����ʱ�ã�Ҳ��Ϊ�û����ļ�ֵ��dwLock0 �� dwLock1 �����û�����У��ֵ�����⣬dwLock1 ���ǲ��ҿ��ֿ������(������ᵽ)��
        nRecur = nRecur - 1

        If (nRecur = 0) Then
          RepStatus = 1 + IIf(bPerpCheck, 2, 0) + IIf(bOppPerpCheck, 4, 0)
'     ��������1(=1+0+0)����ʾ�����ظ����棬��˫�����޳���(�к�)��
'     ��������3(=1 + 2)����ʾ�����ظ����棬���������泤��(�б�����)��
'     ��������5(=1 + 4)����ʾ�����ظ����棬�Է������泤��(�жԷ���)��
'     ��������7(=1+2+4)����ʾ�����ظ����棬˫������(�к�)��
          Exit Function
        End If
      End If

    Else '���Է�
      bOppPerpCheck = bOppPerpCheck And mvsList(lpmvs).ucbCheck '��һ�����ڷǽ��߷�ʱ�͸�Ϊ False
    End If

    bSelfSide = Not bSelfSide
    lpmvs = lpmvs - 1 '�Ӻ���ǰ���
  Loop

  RepStatus = 0 '����0����ʾû���ظ����棻
End Function

'�ظ������ֵ
Public Function RepValue(ByVal nRepStatus As Byte) As Integer
  Dim vlReturn As Long
  
  vlReturn = IIf((nRepStatus And 2) = 0, 0, nDistance - MATE_VALUE) + IIf((nRepStatus And 4) = 0, 0, MATE_VALUE - nDistance)
  RepValue = IIf(vlReturn = 0, DrawValue(), vlReturn)
End Function

'�Ƿ񱻽���
Public Function InCheck() As Boolean
  InCheck = mvsList(nMoveNum - 1).ucbCheck
End Function
 
'�ж��Ƿ������ղ��ü�
Public Function NullOkay() As Boolean
  NullOkay = IIf(sdPlayer = 0, vlWhite, vlBlack) > NULL_MARGIN
End Function

'��һ���ղ�
Public Sub NullMove()
  Dim dwKey As Long
  
  dwKey = zobr.dwKey
  ChangeSide
  mvsList(nMoveNum).Set_ 0, 0, False, dwKey
  nMoveNum = nMoveNum + 1
  nDistance = nDistance + 1
End Sub

'������һ���ղ�
Public Sub UndoNullMove()
  nDistance = nDistance - 1
  nMoveNum = nMoveNum - 1
  ChangeSide
End Sub

'��һ���Ƿ����
Public Function Captured() As Boolean
  Captured = mvsList(nMoveNum - 1).ucpcCaptured <> 0
End Function

'���(��ʼ��)��ʷ�߷���Ϣ
Public Sub SetIrrev()
  mvsList(0).Set_ 0, 0, Checked(), zobr.dwKey
  nMoveNum = 1
End Sub

Private Function DrawValue() As Long
  DrawValue = IIf((nDistance And 1) = 0, -DRAW_VALUE, DRAW_VALUE)
End Function

'�������
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
