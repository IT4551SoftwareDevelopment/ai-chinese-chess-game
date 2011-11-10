VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "VB AI �й����� (Ver 6.1)"
   ClientHeight    =   8640
   ClientLeft      =   2475
   ClientTop       =   660
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9345
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4425
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ListBox lstMoveDesc 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8520
      Left            =   7800
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Timer tmrSearch 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3630
      Top             =   1725
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuHumanRed 
         Caption         =   "������(&R)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuHumanBlack 
         Caption         =   "��������(&B)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuChart 
         Caption         =   "���Ʊ仯ͼ(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSearchInfo 
         Caption         =   "����״̬��Ϣ(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "����(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFlipped 
         Caption         =   "��ת����(&F)"
      End
   End
   Begin VB.Menu mnuComputer 
      Caption         =   "����(&E)"
      Begin VB.Menu mnuComputerMoveRed 
         Caption         =   "����ִ��(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuComputerMoveBlack 
         Caption         =   "����ִ��(&B)"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan
'ȥ���Ի�������������첽���Ż���û��ʵ�֣�
'���������˶���ʱ�������Ϊ1��3��5���� ��bug

Private Const IDB_BOARD As Byte = 200
Private Const IDB_SELECTED As Byte = 201

Private Const IDB_RK  As Byte = 208
Private Const IDB_RA  As Byte = 209
Private Const IDB_RB  As Byte = 210
Private Const IDB_RN  As Byte = 211
Private Const IDB_RR  As Byte = 212
Private Const IDB_RC  As Byte = 213
Private Const IDB_RP  As Byte = 214

Private Const IDB_BK  As Byte = 216
Private Const IDB_BA  As Byte = 217
Private Const IDB_BB  As Byte = 218
Private Const IDB_BN  As Byte = 219
Private Const IDB_BR  As Byte = 220
Private Const IDB_BC  As Byte = 221
Private Const IDB_BP  As Byte = 222

Private Const IDR_CLICK    As Integer = 300
Private Const IDR_ILLEGAL  As Integer = 301
Private Const IDR_MOVE     As Integer = 302
Private Const IDR_MOVE2    As Integer = 303
Private Const IDR_CAPTURE  As Integer = 304
Private Const IDR_CAPTURE2 As Integer = 305
Private Const IDR_CHECK    As Integer = 306
Private Const IDR_CHECK2   As Integer = 307
Private Const IDR_WIN      As Integer = 308
Private Const IDR_DRAW     As Integer = 309
Private Const IDR_LOSS     As Integer = 310

Private Const NO_NULL As Boolean = True

'���ںͻ�ͼ����
Private Const MASK_COLOR As Long = vbGreen
Private Const SQUARE_SIZE As Integer = 56
Private Const BOARD_EDGE As Integer = 8
Private Const BOARD_WIDTH As Integer = BOARD_EDGE + SQUARE_SIZE * 9 + BOARD_EDGE
Private Const BOARD_HEIGHT As Integer = BOARD_EDGE + SQUARE_SIZE * 10 + BOARD_EDGE

'���̷�Χ
Private Const RANK_TOP As Byte = 3
Private Const RANK_BOTTOM As Byte = 12
Private Const FILE_LEFT As Byte = 3
Private Const FILE_RIGHT As Byte = 11

'"DrawSquare"����
Private Const DRAW_SELECTED As Boolean = True

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "KERNEL32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Type udtXqwl
  hInst As Long
  hWnd As Long
  hdc As Long
  hdcTmp As Long              '�豸�����ֻ��"ClickSquare"��������Ч
  bmpBoard As Long
  bmpSelected As Long
  bmpPieces(0 To 23) As Long  '��ԴͼƬ���
  sqSelected As Byte
  mvLast As Long              'ѡ�еĸ��ӣ���һ����
  bFlipped As Boolean         '�Ƿ�ת����
  bGameOver As Boolean        '�Ƿ���Ϸ����(���ü�������ȥ)
End Type

Private Xqwl As udtXqwl
Private pic(0 To 15) As StdPicture

Private mSearchStartTime As Long

Private mVisitNodesALayer As Long
Private mQSNodesALayer    As Long
Private mMoveCount        As Integer
'Private mSaveGenMoves     As Byte

Private mbIsRedTurn       As Boolean
Private mGetOpenDesc      As New clsGetOpenDesc

Private Type MSGBOXPARAMS
  cbSize As Long
  hwndOwner As Long
  hInstance As Long
  lpszText As String
  lpszCaption As String
  dwStyle As Long
  lpszIcon As Long
  dwContextHelpId As Long
  lpfnMsgBoxCallback As Long
  dwLanguageId As Long
End Type

Private Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Const MB_ICONASTERISK = &H40&
Private Const MB_USERICON = &H80&

Private m_MoveList() As Integer
Private mbCallFromOutside As Boolean

'��ʼ�����
Private Sub Startup()
  pos.Startup
    
  Xqwl.sqSelected = 0
  Xqwl.mvLast = 0
  Xqwl.bGameOver = False
  mbIsRedTurn = True
  
  frmSearchInfo.lsvValue.ListItems.Clear
  frmChart.mHGrids = 3
  frmChart.mStepWidth = frmChart.picChart.Width / frmChart.mHGrids
  mMoveCount = 0
  lstMoveDesc.Clear
  lstMoveDesc.AddItem "===��ʼ==="
End Sub
 
Private Sub Form_Load()
  Dim i As Byte
  Dim s As String
  
  '��ʼ��ȫ�ֱ���
  InitConstantArray
  Randomize Timer
  InitZobrist
  
  Xqwl.hInst = App.hInstance
  Xqwl.hWnd = Me.hWnd
  LoadBook
  Xqwl.bFlipped = False
  Startup
  
  'װ��ͼƬ
  Xqwl.bmpBoard = LoadResBmp(IDB_BOARD)
  Xqwl.bmpSelected = LoadResBmp(IDB_SELECTED)
  
  For i = PIECE_KING To PIECE_PAWN
    Xqwl.bmpPieces(SIDE_TAG(0) + i) = LoadResBmp(IDB_RK + i)
    Xqwl.bmpPieces(SIDE_TAG(1) + i) = LoadResBmp(IDB_BK + i)
  Next i
  
  Me.Width = Me.ScaleX(BOARD_WIDTH, vbPixels, vbTwips) + (Me.Width - Me.ScaleWidth) + lstMoveDesc.Width
  Me.Height = Me.ScaleY(BOARD_HEIGHT, vbPixels, vbTwips) + (Me.Height - Me.ScaleHeight)
  
  frmSearchInfo.Move Screen.Width - frmSearchInfo.Width, 0
  Me.Move frmSearchInfo.Left - Me.Width, 0
  frmSettings.Move Screen.Width - frmSettings.Width, Screen.Height - frmSettings.Height
  frmChart.Move frmSettings.Left - frmChart.Width, Me.Top + Me.Height
  
  frmChart.Show
  frmSearchInfo.Show
  frmSettings.Show
  
  With frmSettings
    Open App.Path & "\Settings.ini" For Input As #1
    Line Input #1, s
    .optLimitTime.Value = s
    Line Input #1, s
    .optLimitDepth.Value = s
    Line Input #1, s
    .txtTimeLimit.Text = s
    Line Input #1, s
    .txtDepthLimit.Text = s
    Line Input #1, s
    .chkOpeningBook.Value = Val(s)
    Line Input #1, s
    .txtRANDOM_MASK.Text = s
    Close #1
  End With
  
  ReDim m_MoveList(0 To 0)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Me.MousePointer = vbHourglass Then
    Exit Sub
  End If
  
  x = Me.ScaleX(x, vbTwips, vbPixels)
  y = Me.ScaleY(y, vbTwips, vbPixels)
  x = FILE_LEFT + (x - BOARD_EDGE) \ SQUARE_SIZE
  y = RANK_TOP + (y - BOARD_EDGE) \ SQUARE_SIZE
  
  If (x >= FILE_LEFT And x <= FILE_RIGHT And y >= RANK_TOP And y <= RANK_BOTTOM) Then
    ClickSquare COORD_XY(x, y)
  End If
End Sub

'���Ƹ���
Private Sub DrawSquare(sq As Byte, Optional bSelected As Boolean = False)
  Dim sqFlipped As Integer
  Dim xx As Integer
  Dim yy As Integer
  Dim pc As Integer

  sqFlipped = IIf(Xqwl.bFlipped, SQUARE_FLIP(sq), sq)
  xx = BOARD_EDGE + (FILE_X(sqFlipped) - FILE_LEFT) * SQUARE_SIZE
  yy = BOARD_EDGE + (RANK_Y(sqFlipped) - RANK_TOP) * SQUARE_SIZE
  SelectObject Xqwl.hdcTmp, Xqwl.bmpBoard
  BitBlt Xqwl.hdc, xx, yy, SQUARE_SIZE, SQUARE_SIZE, Xqwl.hdcTmp, xx, yy, vbSrcCopy
  pc = pos.ucpcSquares(sq)
  If (pc <> 0) Then
    DrawTransBmp Xqwl.hdc, Xqwl.hdcTmp, xx, yy, Xqwl.bmpPieces(pc)
  End If
  If (bSelected) Then
    DrawTransBmp Xqwl.hdc, Xqwl.hdcTmp, xx, yy, Xqwl.bmpSelected
  End If
End Sub

'��������¼�����
Private Sub ClickSquare(sq As Byte)
  Dim pc    As Long
  Dim mv    As Integer
  Dim vlRep As Long
  Dim sVar  As String
  
  Xqwl.hdc = GetDC(Xqwl.hWnd)
  Xqwl.hdcTmp = CreateCompatibleDC(Me.hdc)
  sq = IIf(Xqwl.bFlipped, SQUARE_FLIP(sq), sq)
  pc = pos.ucpcSquares(sq)


  If ((pc And SIDE_TAG(pos.sdPlayer)) <> 0) Then
    '�������Լ����ӣ���ôֱ��ѡ�и���
    If (Xqwl.sqSelected <> 0) Then
      DrawSquare (Xqwl.sqSelected)
    End If
    
    Xqwl.sqSelected = sq
    DrawSquare sq, DRAW_SELECTED
    
    If (Xqwl.mvLast <> 0) Then
      DrawSquare (SRC(Xqwl.mvLast))
      DrawSquare (DST(Xqwl.mvLast))
    End If
    
    PlayResWav (IDR_CLICK) '���ŵ��������
    
    If lstMoveDesc.ListIndex < lstMoveDesc.ListCount - 1 And Xqwl.bGameOver Then
      Xqwl.bGameOver = False
    End If
    
  ElseIf (Xqwl.sqSelected <> 0 And Not Xqwl.bGameOver) Then
    '�������Ĳ����Լ����ӣ�������ѡ����(һ�����Լ�����)����ô�������
    mv = MOVE_(Xqwl.sqSelected, sq)
    
    If pos.LegalMove(mv) Then
      If pos.MakeMove(mv) Then
        Xqwl.mvLast = mv
        
        '2011��11���������ܣ���¼ÿ�����ƶ������ڻ���
        With lstMoveDesc
          If .ListIndex < .ListCount - 1 And .ListIndex > -1 Then
            ReDim Preserve m_MoveList(0 To .ListIndex + 1)
            m_MoveList(.ListIndex) = mv
            mMoveCount = Round(CSng(.ListIndex) / 2 + 0.1)
            Dim i As Integer
            For i = .ListCount - 1 To .ListIndex + 1 Step -1
              .RemoveItem i
            Next i
          Else
            m_MoveList(UBound(m_MoveList)) = mv
            ReDim Preserve m_MoveList(0 To UBound(m_MoveList) + 1)
          End If
        End With
        
        
        DrawSquare Xqwl.sqSelected, DRAW_SELECTED
        DrawSquare sq, DRAW_SELECTED
        Xqwl.sqSelected = 0
        
        '���߷��������� ListBox ��
        If mnuComputerMoveRed.Checked Or mnuComputerMoveBlack.Checked Or mbIsRedTurn Then
          mMoveCount = mMoveCount + 1
        End If
        
        If Not mbIsRedTurn Then
          lstMoveDesc.AddItem Space$(Fix(1 * Log(mMoveCount) / Log(10)) + 2) & GetMoveDesc(mv, False)
        Else
          lstMoveDesc.AddItem CStr(mMoveCount) & "." & GetMoveDesc(mv, False)
        End If
        
        mbCallFromOutside = True
        lstMoveDesc.ListIndex = lstMoveDesc.ListCount - 1
        mbCallFromOutside = False
        
        '��ʾ��������
        pos.sMoveSymbolDesc = pos.sMoveSymbolDesc & GetMoveDesc(mv, False, True)
        sVar = mGetOpenDesc.GetVar(pos.sMoveSymbolDesc)
        frmChart.lblOpening.Caption = mGetOpenDesc.GetOpen(pos.sMoveSymbolDesc) _
                              & "(" & mGetOpenDesc.GetEccoNo(pos.sMoveSymbolDesc) & ")" _
                                    & IIf(sVar = "", "", "����" & sVar)
        
        '�������ӷ�
        mbIsRedTurn = Not mbIsRedTurn
        
        '����ظ�����
        vlRep = pos.RepStatus(3)
        
        If pos.IsMate() Then
          '����ֳ�ʤ������ô����ʤ�������������ҵ�����ʾ��
          PlayResWav IDR_WIN
          MessageBoxMute "ף����ȡ��ʤ����"
          Xqwl.bGameOver = True
        ElseIf vlRep > 0 Then
          vlRep = pos.RepValue(vlRep)
          'ע�⣺"vlRep"�ǶԵ�����˵�ķ�ֵ
          PlayResWav IIf(vlRep > WIN_VALUE, IDR_LOSS, IIf(vlRep < -WIN_VALUE, IDR_WIN, IDR_DRAW))
          MessageBoxMute IIf(vlRep > WIN_VALUE, "�����������벻Ҫ���٣�", IIf(vlRep < -WIN_VALUE, "���Գ���������ף����ȡ��ʤ����", "˫���������ͣ������ˣ�"))
          Xqwl.bGameOver = True
        ElseIf pos.nMoveNum > 100 Then
          PlayResWav (IDR_DRAW)
          MessageBoxMute "������Ȼ�������ͣ������ˣ�"
          Xqwl.bGameOver = True
        Else
          '���û�зֳ�ʤ������ô���Ž��������ӻ�һ�����ӵ�����
          PlayResWav IIf(pos.Checked(), IDR_CHECK, IIf(pc <> 0, IDR_CAPTURE, IDR_MOVE))
          If pos.Captured() Then
            pos.SetIrrev
          End If
          
          Do While ((mbIsRedTurn And mnuComputerMoveRed.Checked) Or (Not mbIsRedTurn And mnuComputerMoveBlack.Checked)) And Not Xqwl.bGameOver
            ResponseMove  '�ֵ���������
            If mbIsRedTurn Then
              mMoveCount = mMoveCount + 1
            End If
            mbIsRedTurn = Not mbIsRedTurn
          Loop
        End If
      Else
        PlayResWav IDR_ILLEGAL '���ű�����������
      End If
    End If
    '��������Ͳ������߷�(������������)����ô���������
  End If
  
  DeleteDC Xqwl.hdcTmp
  ReleaseDC Xqwl.hWnd, Xqwl.hdc
End Sub

Private Sub Form_Paint()
  If Me.MousePointer <> vbHourglass Then
    DrawBoard Me.hdc
  End If
End Sub

'װ����ԴͼƬ
Private Function LoadResBmp(nResId As Integer) As Long
  Static p As Byte
  
  Set pic(p) = LoadResPicture(nResId, 0)
  LoadResBmp = pic(p).Handle
  p = p + 1
End Function

'����͸��ͼƬ
Private Sub DrawTransBmp(lngHdc As Long, hdcTmp As Long, xx As Integer, yy As Integer, bmp As Long)
  SelectObject hdcTmp, bmp
  TransparentBlt lngHdc, xx, yy, SQUARE_SIZE, SQUARE_SIZE, hdcTmp, 0, 0, 56, 56, MASK_COLOR
End Sub

'��������
Private Sub DrawBoard(lngHdc As Long)
  Dim x As Integer
  Dim y As Integer
  Dim xx As Integer
  Dim yy As Integer
  Dim sq As Integer
  Dim pc As Integer
  Dim hdcTmp As Long

  '������
  hdcTmp = CreateCompatibleDC(lngHdc)
  SelectObject hdcTmp, Xqwl.bmpBoard
  TransparentBlt lngHdc, 0, 0, BOARD_WIDTH, BOARD_HEIGHT, hdcTmp, 0, 0, 520, 576, MASK_COLOR
  
  '������
  For x = FILE_LEFT To FILE_RIGHT
    For y = RANK_TOP To RANK_BOTTOM
    
      If (Xqwl.bFlipped) Then
        xx = BOARD_EDGE + (FILE_FLIP(x) - FILE_LEFT) * SQUARE_SIZE
        yy = BOARD_EDGE + (RANK_FLIP(y) - RANK_TOP) * SQUARE_SIZE
      Else
        xx = BOARD_EDGE + (x - FILE_LEFT) * SQUARE_SIZE
        yy = BOARD_EDGE + (y - RANK_TOP) * SQUARE_SIZE
      End If
      
      sq = COORD_XY(x, y)
      pc = pos.ucpcSquares(sq)
      
      If (pc <> 0) Then
        DrawTransBmp lngHdc, hdcTmp, xx, yy, Xqwl.bmpPieces(pc)
      End If
      
      '��ѡ���
      If sq = Xqwl.sqSelected Or sq = SRC(Xqwl.mvLast) Or sq = DST(Xqwl.mvLast) Then
        DrawTransBmp lngHdc, hdcTmp, xx, yy, Xqwl.bmpSelected
      End If
      
    Next y
  Next x
  DeleteDC (hdcTmp)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Open App.Path & "\Settings.ini" For Output As #1
  With frmSettings
    Print #1, .optLimitTime.Value
    Print #1, .optLimitDepth.Value
    Print #1, .txtTimeLimit.Text
    Print #1, .txtDepthLimit.Text
    Print #1, CStr(.chkOpeningBook.Value)
    Print #1, .txtRANDOM_MASK.Text
  End With
  Close #1
  End
End Sub

Private Sub lstMoveDesc_Click()
  Dim i As Integer
  
  If mbCallFromOutside Then
    Exit Sub
  End If
  
  pos.Startup
  Xqwl.sqSelected = 0

  For i = 0 To lstMoveDesc.ListIndex - 1
    pos.MakeMove m_MoveList(i)
  Next i
  
  If lstMoveDesc.ListIndex = 0 Then
    Xqwl.mvLast = 0
  Else
    Xqwl.mvLast = m_MoveList(i - 1)
  End If
  
  Form_Paint
End Sub

Public Sub mnuChart_Click()
  mnuChart.Checked = Not mnuChart.Checked
  frmChart.Visible = mnuChart.Checked
End Sub

Private Sub mnuComputerMoveBlack_Click()
  mnuComputerMoveBlack.Checked = Not mnuComputerMoveBlack.Checked
  
  If mnuComputerMoveBlack.Checked Then
    Xqwl.hdc = GetDC(Xqwl.hWnd)
    Xqwl.hdcTmp = CreateCompatibleDC(Xqwl.hdc)
    
    Do While ((mbIsRedTurn And mnuComputerMoveRed.Checked) Or (Not mbIsRedTurn And mnuComputerMoveBlack.Checked)) And Not Xqwl.bGameOver
      If mbIsRedTurn Then
        mMoveCount = mMoveCount + 1
      End If
      ResponseMove  '�ֵ���������
      mbIsRedTurn = Not mbIsRedTurn
    Loop
  End If
End Sub

Private Sub mnuComputerMoveRed_Click()
  mnuComputerMoveRed.Checked = Not mnuComputerMoveRed.Checked

  If mnuComputerMoveRed.Checked Then
    Xqwl.hdc = GetDC(Xqwl.hWnd)
    Xqwl.hdcTmp = CreateCompatibleDC(Xqwl.hdc)
    
    Do While ((mbIsRedTurn And mnuComputerMoveRed.Checked) Or (Not mbIsRedTurn And mnuComputerMoveBlack.Checked)) And Not Xqwl.bGameOver
      ResponseMove  '�ֵ���������
      If mbIsRedTurn Then
        mMoveCount = mMoveCount + 1
      End If
      mbIsRedTurn = Not mbIsRedTurn
    Loop
  End If
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuFlipped_Click()
  mnuFlipped.Checked = Not mnuFlipped.Checked
  Xqwl.bFlipped = mnuFlipped.Checked
  DrawBoard hdc
End Sub

Private Sub mnuHumanRed_Click()
  Xqwl.bFlipped = False
  mnuFlipped.Checked = False
  mnuComputerMoveBlack.Checked = True
  mnuComputerMoveRed.Checked = False
  ResetGame
End Sub

Private Sub mnuHumanBlack_Click()
  Xqwl.bFlipped = True
  mnuFlipped.Checked = True
  mnuComputerMoveRed.Checked = True
  mnuComputerMoveBlack.Checked = False
  ResetGame
End Sub

Private Sub ResetGame()
  Dim hdc As Long
  
  Startup
  hdc = GetDC(Xqwl.hWnd)
  DrawBoard hdc
  
  If (Xqwl.bFlipped) Then
    Xqwl.hdc = hdc
    Xqwl.hdcTmp = CreateCompatibleDC(Xqwl.hdc)
    ResponseMove
    mbIsRedTurn = Not mbIsRedTurn
    DeleteDC Xqwl.hdcTmp
  End If
      
  ReleaseDC Xqwl.hWnd, hdc
  
  frmChart.lblOpening.Caption = "��������"
End Sub

'��̬(Quiescence)��������
Private Function SearchQuiesc(vlAlpha As Long, vlBeta As Long) As Long
  Dim i As Long
  Dim nGenMoves As Long
  Dim vl As Long
  Dim vlBest As Long
  Dim mvs(0 To MAX_GEN_MOVES - 1) As Integer

  '"GenerateMoves"����
  Const GEN_CAPTURE As Boolean = True

  mQSNodesALayer = mQSNodesALayer + 1
  'һ����̬������Ϊ���¼����׶�

  '1. ����ظ�����
  vl = pos.RepStatus()
  If (vl <> 0) Then
    SearchQuiesc = pos.RepValue(vl)
    Exit Function
  End If

  '2. ���Ｋ����Ⱦͷ��ؾ�������
  If (pos.nDistance = LIMIT_DEPTH) Then
    SearchQuiesc = pos.Evaluate()
    Exit Function
  End If

  '3. ��ʼ�����ֵ
  vlBest = -MATE_VALUE '��������֪�����Ƿ�һ���߷���û�߹�(ɱ��)

  If (pos.InCheck()) Then
    '4. �����������������ȫ���߷�
    nGenMoves = pos.GenerateMoves(mvs)
    qsort mvs, 0, nGenMoves - 1, "History"
  Else
    '5. �������������������������
    vl = pos.Evaluate()
    If (vl > vlBest) Then
      vlBest = vl
      If (vl >= vlBeta) Then
        SearchQuiesc = vl
        Exit Function
      End If

      If (vl > vlAlpha) Then
        vlAlpha = vl
      End If
    End If

    '6. �����������û�нضϣ������ɳ����߷�
    nGenMoves = pos.GenerateMoves(mvs, GEN_CAPTURE)
    If frmSettings.chkSortByMvvLva.Value = vbChecked Then
      If frmSettings.optQuickSort.Value Then
        qsort mvs, 0, nGenMoves - 1, "MvvLva"
      Else
        SelectSort mvs, nGenMoves, "MvvLva"
      End If
    End If
  End If

  '7. ��һ����Щ�߷��������еݹ�
  For i = 0 To nGenMoves - 1
    If (pos.MakeMove(mvs(i))) Then
      vl = -SearchQuiesc(-vlBeta, -vlAlpha)
      pos.UndoMakeMove

      '8. ����Alpha-Beta��С�жϺͽض�
      If (vl > vlBest) Then    '�ҵ����ֵ(������ȷ����Alpha��PV����Beta�߷�)
        vlBest = vl            'vlBest ����ĿǰҪ���ص����ֵ�����ܳ���Alpha-Beta�߽�
        If (vl >= vlBeta) Then '�ҵ�һ��Beta�߷�
          SearchQuiesc = vl
          Exit Function 'Beta�ض�
        End If
        
        If (vl > vlAlpha) Then '�ҵ�һ��PV�߷�
          vlAlpha = vl     '��СAlpha-Beta�߽�
        End If
      End If
    End If
  Next i

  '9. �����߷����������ˣ��������ֵ
  SearchQuiesc = IIf(vlBest = -MATE_VALUE, pos.nDistance - MATE_VALUE, vlBest)
End Function

'�����߽�(Fail-Soft)��Alpha-Beta��������
Private Function SearchFull(vlAlpha As Long, vlBeta As Long, ByVal nDepth As Long, Optional bNoNull As Boolean = False) As Long
  'bInCheckExt �����������������ʱ ListView ���ڸ��ڵ㴦�������ӽڵ㣩����߷���ɵ�bug
'  Dim i         As Byte
  Dim vl        As Long
  Dim vlBest    As Long
  Dim mvBest    As Long
  Dim nHashFlag As Long
  Dim mv        As Integer
  Dim mvHash    As Integer
  Dim Sort      As New SortStruct
  Dim nNewDepth As Integer
  Dim bCheckExt As Boolean

  'ͳ�Ʒ��ʵĽڵ���
  mVisitNodesALayer = mVisitNodesALayer + 1
  
  'һ��Alpha-Beta��ȫ������Ϊ���¼����׶�
  
  '1. ����ˮƽ�ߣ�����þ�̬����(ע�⣺���ڿղ��ü�����ȿ���С����)
  If (nDepth <= 0) Then
    If frmSettings.chkSearchQuiesc.Value = vbChecked Then
      SearchFull = SearchQuiesc(vlAlpha, vlBeta)
    Else
      SearchFull = pos.Evaluate()
    End If
    Exit Function
  End If

  '1-1. ����ظ�����(ע�⣺��Ҫ�ڸ��ڵ��飬�����û���߷���)
  vl = pos.RepStatus()
  If (vl <> 0) Then
    SearchFull = pos.RepValue(vl)
    Exit Function
  End If

  '1-2. ���Ｋ����Ⱦͷ��ؾ�������
  If (pos.nDistance = LIMIT_DEPTH) Then
    SearchFull = pos.Evaluate()
    Exit Function
  End If

  '1-3. �����û���ü������õ��û����߷�
  If frmSettings.chkTranTable.Value = vbChecked Then
    vl = ProbeHash(vlAlpha, vlBeta, nDepth, mvHash)
    If (vl > -MATE_VALUE) Then
      SearchFull = vl
      Exit Function
    End If
  End If

  '1-4. ���Կղ��ü�(���ڵ��Betaֵ��"MATE_VALUE"�����Բ����ܷ����ղ��ü�)
  If ((Not bNoNull) And (Not pos.InCheck()) And pos.NullOkay() And (frmSettings.chkNullMove.Value = vbChecked)) Then
    pos.NullMove
    vl = -SearchFull(-vlBeta, 1 - vlBeta, nDepth - NULL_DEPTH - 1, NO_NULL)
    pos.UndoNullMove
    If (vl >= vlBeta) Then
      SearchFull = vl
      Exit Function
    End If
  End If
  
  '2. ��ʼ�����ֵ������߷�
  nHashFlag = HASH_ALPHA
  vlBest = -MATE_VALUE '��������֪�����Ƿ�һ���߷���û�߹�(ɱ��)
  mvBest = 0           '��������֪�����Ƿ���������Beta�߷���PV�߷����Ա㱣�浽��ʷ��

  '3. ��ʼ���߷�����ṹ
  Sort.Init mvHash
  
  '4. ��һ����Щ�߷��������еݹ�
  mv = Sort.Next_
  
  DoEvents '�ü�ʱʱ���ߣ�ͬʱˢ�� ListView ����ʾ

  Do While (mv <> 0)
    If pos.MakeMove(mv) Then
      '��������
      If pos.InCheck And frmSettings.chkInCheckExt.Value = vbChecked Then
        nNewDepth = nDepth
        bCheckExt = True
      Else
        nNewDepth = nDepth - 1
        bCheckExt = False
      End If
      
      'PVS
      If vlBest = -MATE_VALUE Then
        vl = -SearchFull(-vlBeta, -vlAlpha, nNewDepth)
      Else
        vl = -SearchFull(-vlAlpha - 1, -vlAlpha, nNewDepth)
        If (vl > vlAlpha And vl < vlBeta) Then
          vl = -SearchFull(-vlBeta, -vlAlpha, nNewDepth)
        End If
      End If
        
      pos.UndoMakeMove

      '5. ����Alpha-Beta��С�жϺͽض�
      If (vl > vlBest) Then    '�ҵ����ֵ(������ȷ����Alpha��PV����Beta�߷�)
        vlBest = vl            '"vlBest"����ĿǰҪ���ص����ֵ�����ܳ���Alpha-Beta�߽�
        If (vl >= vlBeta) Then '�ҵ�һ��Beta�߷�
          nHashFlag = HASH_BETA
          mvBest = mv          'Beta�߷�Ҫ���浽��ʷ��
          Exit Do              'Beta�ض�
        End If
        If (vl > vlAlpha) Then '�ҵ�һ��PV�߷�
          nHashFlag = HASH_PV
          mvBest = mv          'PV�߷�Ҫ���浽��ʷ��
          vlAlpha = vl         '��СAlpha-Beta�߽�
        End If
      End If
    End If
    
    mv = Sort.Next_
  Loop
  
  '5. �����߷����������ˣ�������߷�(������Alpha�߷�)���浽��ʷ���������ֵ
  If (vlBest = -MATE_VALUE) Then
    '�����ɱ�壬�͸���ɱ�岽����������
    SearchFull = pos.nDistance - MATE_VALUE
    Exit Function
  End If

  If frmSettings.chkTranTable.Value = vbChecked Then
    '��¼���û���
    RecordHash nHashFlag, vlBest, nDepth, mvBest
  End If
  
  If (mvBest <> 0) Then
    '�������Alpha�߷����ͽ�����߷����浽��ʷ��
    SetBestMove mvBest, nDepth
  End If
  SearchFull = vlBest
End Function

'���ڵ��Alpha-Beta��������
Private Function SearchRoot(nDepth As Long) As Long
  Dim vl        As Long
  Dim vlBest    As Long
  Dim mv        As Integer
  Dim nNewDepth As Byte
  Dim Sort      As New SortStruct
  Dim i         As Byte
  Dim vlRealBest As Long

  vlBest = -MATE_VALUE
  Sort.Init Search.mvResult
  
  mv = Sort.Next_()
  If Not pos.LegalMove(mv) Then
    mv = Sort.Next_()
  End If
  i = i + 1
  '���洦��
  frmSearchInfo.lsvMoveList.ListItems.Clear
  frmSearchInfo.lsvMoveList.ListItems.Add , , CStr(i)
  frmSearchInfo.lsvMoveList.ListItems(i).SubItems(1) = GetMoveDesc(mv)
  frmSearchInfo.lsvDepthTimeCost.ListItems(nDepth).SubItems(1) = CStr(i) & "/" & CStr(Sort.nGenMoves)
  
  DoEvents '�ü�ʱʱ���ߣ�ͬʱˢ�� ListView ����ʾ

  Do While (mv <> 0)
    If (pos.MakeMove(mv)) Then
      nNewDepth = IIf(pos.InCheck(), nDepth, nDepth - 1)
      If (vlBest = -MATE_VALUE) Then
        vl = -SearchFull(-MATE_VALUE, MATE_VALUE, nNewDepth, NO_NULL)
      Else
        vl = -SearchFull(-vlBest - 1, -vlBest, nNewDepth)
        If (vl > vlBest) Then
          vl = -SearchFull(-MATE_VALUE, -vlBest, nNewDepth, NO_NULL)
        End If
      End If
      
      '���洦��
'      If nDepth = frmSearchInfo.lsvDepthTimeCost.ListItems.Count Then
        frmSearchInfo.lsvMoveList.ListItems(i).SubItems(2) = CStr(vl)
        
        If (vl > vlBest) Then
          frmSearchInfo.lsvMoveList.ListItems(i).Selected = True
        End If
'      End If
      
      pos.UndoMakeMove
      
      If (vl > vlBest) Then
        vlBest = vl
        vlRealBest = vl
        Search.mvResult = mv
        If (vlBest > -WIN_VALUE And vlBest < WIN_VALUE) Then
          vlBest = vlBest + ((Rnd * Val(frmSettings.txtRANDOM_MASK.Text)) - (Rnd * Val(frmSettings.txtRANDOM_MASK.Text)))  '����������Եľ������ڰ�~~~
        End If
      End If
    Else '���ⲽ��󱻽���
      frmSearchInfo.lsvMoveList.ListItems(i).SubItems(2) = CStr(-MATE_VALUE)
    End If
    mv = Sort.Next_()
    i = i + 1
    
    '���洦��
    If mv <> 0 Then
      frmSearchInfo.lsvMoveList.ListItems.Add , , CStr(i)
      frmSearchInfo.lsvMoveList.ListItems(i).SubItems(1) = GetMoveDesc(mv)
      frmSearchInfo.lsvDepthTimeCost.ListItems(nDepth).SubItems(1) = CStr(i) & "/" & CStr(Sort.nGenMoves)
    End If
  Loop
  RecordHash HASH_PV, vlBest, nDepth, Search.mvResult
  SetBestMove Search.mvResult, nDepth
  SearchRoot = vlRealBest
End Function

'������߷��Ĵ���
Private Sub SetBestMove(mv As Long, nDepth As Long)
  Search.nHistoryTable(mv) = Search.nHistoryTable(mv) + (nDepth * nDepth)
  
  If (Search.mvKillers(pos.nDistance, 0) <> mv) Then
    Search.mvKillers(pos.nDistance, 1) = Search.mvKillers(pos.nDistance, 0)
    Search.mvKillers(pos.nDistance, 0) = mv
  End If
End Sub

'���Ի�Ӧһ����
Private Sub ResponseMove()
  Dim vlRep As Long
  Dim sVar  As String

  '������һ����
  Me.MousePointer = vbHourglass
  SearchMain
  Me.MousePointer = vbArrow
  pos.MakeMove Search.mvResult
  m_MoveList(UBound(m_MoveList)) = Search.mvResult
  ReDim Preserve m_MoveList(0 To UBound(m_MoveList) + 1)
  
  If mbIsRedTurn Then
    lstMoveDesc.AddItem CStr(mMoveCount + 1) & "." & GetMoveDesc(Search.mvResult, False)
  Else
    lstMoveDesc.AddItem Space$(Fix(CSng(Log(mMoveCount) / Log(10))) + 2) & GetMoveDesc(Search.mvResult, False)
  End If
  
  mbCallFromOutside = True
  lstMoveDesc.ListIndex = lstMoveDesc.ListCount - 1
  mbCallFromOutside = False
  
  '��ʾ��������
  pos.sMoveSymbolDesc = pos.sMoveSymbolDesc & GetMoveDesc(Search.mvResult, False, True)
  sVar = mGetOpenDesc.GetVar(pos.sMoveSymbolDesc)
  frmChart.lblOpening.Caption = mGetOpenDesc.GetOpen(pos.sMoveSymbolDesc) _
                        & "(" & mGetOpenDesc.GetEccoNo(pos.sMoveSymbolDesc) & ")" _
                              & IIf(sVar = "", "", "����" & sVar)
  
  '�����һ�����ѡ����
  DrawSquare (SRC(Xqwl.mvLast))
  DrawSquare (DST(Xqwl.mvLast))
  
  '�ѵ����ߵ����ǳ���
  Xqwl.mvLast = Search.mvResult
  DrawSquare SRC(Xqwl.mvLast), DRAW_SELECTED
  DrawSquare DST(Xqwl.mvLast), DRAW_SELECTED
  
  '����ظ�����
  vlRep = pos.RepStatus(3)
  
  If (pos.IsMate()) Then
    '����ֳ�ʤ������ô����ʤ�������������ҵ���������������ʾ��
    PlayResWav (IDR_LOSS)
    MessageBoxMute "���ٽ�������"
    Xqwl.bGameOver = True
  ElseIf vlRep > 0 Then
    vlRep = pos.RepValue(vlRep)
    'ע�⣺"vlRep"�Ƕ������˵�ķ�ֵ
    PlayResWav IIf(vlRep < -WIN_VALUE, IDR_LOSS, IIf(vlRep > WIN_VALUE, IDR_WIN, IDR_DRAW))
    MessageBoxMute IIf(vlRep < -WIN_VALUE, "�����������벻Ҫ���٣�", IIf(vlRep > WIN_VALUE, "���Գ���������ף����ȡ��ʤ����", "˫���������ͣ������ˣ�"))
    Xqwl.bGameOver = True
  ElseIf pos.nMoveNum > 100 Then
    PlayResWav (IDR_DRAW)
    MessageBoxMute "������Ȼ�������ͣ������ˣ�"
    Xqwl.bGameOver = True
  Else
    '���û�зֳ�ʤ������ô���Ž��������ӻ�һ�����ӵ�����
    PlayResWav IIf(pos.Checked(), IDR_CHECK2, IIf(pos.Captured, IDR_CAPTURE2, IDR_MOVE2))
    If pos.Captured() Then
      pos.SetIrrev
    End If
  End If
End Sub

'����������������
Private Sub SearchMain()
  Dim i               As Long
  Dim vl              As Long
  Dim TimeElapsed     As Long
  Dim VisitNodesTotal As Long
  Dim mvs(0 To MAX_GEN_MOVES - 1) As Integer
  Dim nGenMoves       As Long
  Dim c               As Byte

  '��ʼ��
  For i = 0 To 65535
    Search.nHistoryTable(i) = 0 '�����ʷ��
  Next i
  
  For i = 0 To LIMIT_DEPTH - 1
    Search.mvKillers(i, 0) = 0 '���ɱ���߷���
    Search.mvKillers(i, 1) = 0 '���ɱ���߷���
  Next i
  
  For i = 0 To HASH_SIZE - 1 '����û���
    Search_HashTable(i).dwLock0 = 0
    Search_HashTable(i).dwLock1 = 0
    Search_HashTable(i).svl = 0
    Search_HashTable(i).ucDepth = 0
    Search_HashTable(i).ucFlag = 0
    Search_HashTable(i).wmv = 0
    Search_HashTable(i).wReserved = 0
  Next i
  
  mSearchStartTime = GetTickCount()  '��ʼ����ʱ��
  pos.nDistance = 0 '��ʼ����
  
  frmSearchInfo.lsvDepthTimeCost.ListItems.Clear
  tmrSearch.Enabled = True
  VisitNodesTotal = 0
  
  If frmSettings.chkOpeningBook.Value = vbChecked Then
    '�������ֿ�
    Search.mvResult = SearchBook()
    If (Search.mvResult <> 0) Then
      pos.MakeMove Search.mvResult
      frmSearchInfo.lsvValue.ListItems(frmSearchInfo.lsvValue.ListItems.Count).SubItems(3) = CStr(GetTickCount() - mSearchStartTime)
      If (pos.RepStatus(3) = 0) Then
        pos.UndoMakeMove
        Exit Sub
      End If
      pos.UndoMakeMove
    End If
    
    '����Ƿ�ֻ��Ψһ�߷�
    vl = 0
    nGenMoves = pos.GenerateMoves(mvs)
    For i = 0 To nGenMoves - 1
      If (pos.MakeMove(mvs(i))) Then
        pos.UndoMakeMove
        Search.mvResult = mvs(i)
        vl = vl + 1
      End If
    Next i
    If (vl = 1) Then
      Exit Sub
    End If
  End If
  
  '�����������
  For i = 1 To LIMIT_DEPTH
    frmSearchInfo.lsvDepthTimeCost.ListItems.Add CStr(i), , CStr(i)
    
    mVisitNodesALayer = 0
    mQSNodesALayer = 0
    
    '�������ڵ�
    vl = SearchRoot(i)
    
    VisitNodesTotal = VisitNodesTotal + mVisitNodesALayer + mQSNodesALayer
    
    TimeElapsed = GetTickCount() - mSearchStartTime
    
    With frmSearchInfo.lsvDepthTimeCost.ListItems(i)
      .SubItems(2) = CStr(TimeElapsed)
      .SubItems(3) = GetMoveDesc(Search.mvResult)
      .SubItems(4) = CStr(vl)
      .SubItems(5) = CStr(mVisitNodesALayer)
      .SubItems(6) = CStr(mQSNodesALayer)
    End With
    
    '������ɱ�壬����ֹ����
    If (vl > WIN_VALUE Or vl < -WIN_VALUE) Then
      Exit For
    End If
    
    If frmSettings.optLimitTime.Value Then
      '����ʱ�ޣ�����ֹ����
      If (TimeElapsed > Val(frmSettings.txtTimeLimit.Text)) Then
        Exit For
      End If
    Else
      '�ﵽ��ȣ�����ֹ����
      If i = Val(frmSettings.txtDepthLimit.Text) Then
        Exit For
      End If
    End If
  Next i
  
  frmSearchInfo.lsvValue.ListItems.Add , , CStr(frmSearchInfo.lsvValue.ListItems.Count + 1)
  c = frmSearchInfo.lsvValue.ListItems.Count
  With frmSearchInfo.lsvValue.ListItems(c)
    .SubItems(1) = CStr(vl)
    .SubItems(2) = CStr(IIf(i > LIMIT_DEPTH, i - 1, i))
    .SubItems(3) = CStr(TimeElapsed)
    .SubItems(4) = CStr(VisitNodesTotal)
    .Selected = True
    .EnsureVisible
  End With

  frmChart.picChart.Cls
  frmChart.DrawGrid
  frmChart.DrawValueChart
  
  tmrSearch.Enabled = False
End Sub

Private Sub mnuSave_Click()
  Dim i As Integer
  Dim s As String
  
  On Error GoTo hErr
  With cdlg
    .Filter = "����ֲ���� (*.PGN)|*.pgn"
    .FileName = Replace(CStr(Now), ":", "��")
    .ShowSave
    Open .FileName For Output As #1
    Print #1, "[Game ""Chinese Chess""]"
    For i = 1 To lstMoveDesc.ListCount
      s = lstMoveDesc.List(i - 1)
      Print #1, s
    Next i
    Close #1
  End With
  Exit Sub
hErr:
End Sub

Public Sub mnuSearchInfo_Click()
  mnuSearchInfo.Checked = Not mnuSearchInfo.Checked
  frmSearchInfo.Visible = mnuSearchInfo.Checked
End Sub

Public Sub mnuSettings_Click()
  mnuSettings.Checked = Not mnuSettings.Checked
  frmSettings.Visible = mnuSettings.Checked
End Sub

Private Sub tmrSearch_Timer()
  Dim TimeElapsed As Integer
  
  TimeElapsed = (GetTickCount() - mSearchStartTime) \ 1000
  frmSearchInfo.lblSearchSeconds.Caption = "������ʱ�䣺" & CStr(TimeElapsed) & " ��"
End Sub

'����������������ʾ��
Private Sub MessageBoxMute(lpszText As String)
  Dim mbp As MSGBOXPARAMS

  mbp.cbSize = Len(mbp)
  mbp.hwndOwner = Xqwl.hWnd
  mbp.lpszText = lpszText
  mbp.lpszCaption = "VB AI �й�����"
  mbp.dwStyle = MB_USERICON
  mbp.lpszIcon = 104
  MessageBoxIndirect mbp
End Sub



