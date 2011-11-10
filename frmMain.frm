VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "VB AI 中国象棋 (Ver 6.1)"
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
         Name            =   "宋体"
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
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuHumanRed 
         Caption         =   "我先走(&R)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuHumanBlack 
         Caption         =   "电脑先走(&B)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuChart 
         Caption         =   "局势变化图(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSearchInfo 
         Caption         =   "搜索状态信息(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "设置(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFlipped 
         Caption         =   "翻转棋盘(&F)"
      End
   End
   Begin VB.Menu mnuComputer 
      Caption         =   "电脑(&E)"
      Begin VB.Menu mnuComputerMoveRed 
         Caption         =   "电脑执红(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuComputerMoveBlack 
         Caption         =   "电脑执黑(&B)"
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
'去掉对话框的声音（但异步播放还是没能实现）
'修正了两人对下时棋谱序号为1、3、5…… 的bug

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

'窗口和绘图属性
Private Const MASK_COLOR As Long = vbGreen
Private Const SQUARE_SIZE As Integer = 56
Private Const BOARD_EDGE As Integer = 8
Private Const BOARD_WIDTH As Integer = BOARD_EDGE + SQUARE_SIZE * 9 + BOARD_EDGE
Private Const BOARD_HEIGHT As Integer = BOARD_EDGE + SQUARE_SIZE * 10 + BOARD_EDGE

'棋盘范围
Private Const RANK_TOP As Byte = 3
Private Const RANK_BOTTOM As Byte = 12
Private Const FILE_LEFT As Byte = 3
Private Const FILE_RIGHT As Byte = 11

'"DrawSquare"参数
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
  hdcTmp As Long              '设备句柄，只在"ClickSquare"过程中有效
  bmpBoard As Long
  bmpSelected As Long
  bmpPieces(0 To 23) As Long  '资源图片句柄
  sqSelected As Byte
  mvLast As Long              '选中的格子，上一步棋
  bFlipped As Boolean         '是否翻转棋盘
  bGameOver As Boolean        '是否游戏结束(不让继续玩下去)
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

'初始化棋局
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
  lstMoveDesc.AddItem "===开始==="
End Sub
 
Private Sub Form_Load()
  Dim i As Byte
  Dim s As String
  
  '初始化全局变量
  InitConstantArray
  Randomize Timer
  InitZobrist
  
  Xqwl.hInst = App.hInstance
  Xqwl.hWnd = Me.hWnd
  LoadBook
  Xqwl.bFlipped = False
  Startup
  
  '装入图片
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

'绘制格子
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

'点击格子事件处理
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
    '如果点击自己的子，那么直接选中该子
    If (Xqwl.sqSelected <> 0) Then
      DrawSquare (Xqwl.sqSelected)
    End If
    
    Xqwl.sqSelected = sq
    DrawSquare sq, DRAW_SELECTED
    
    If (Xqwl.mvLast <> 0) Then
      DrawSquare (SRC(Xqwl.mvLast))
      DrawSquare (DST(Xqwl.mvLast))
    End If
    
    PlayResWav (IDR_CLICK) '播放点击的声音
    
    If lstMoveDesc.ListIndex < lstMoveDesc.ListCount - 1 And Xqwl.bGameOver Then
      Xqwl.bGameOver = False
    End If
    
  ElseIf (Xqwl.sqSelected <> 0 And Not Xqwl.bGameOver) Then
    '如果点击的不是自己的子，但有子选中了(一定是自己的子)，那么走这个子
    mv = MOVE_(Xqwl.sqSelected, sq)
    
    If pos.LegalMove(mv) Then
      If pos.MakeMove(mv) Then
        Xqwl.mvLast = mv
        
        '2011年11月新增功能，记录每步的移动，用于悔棋
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
        
        '将走法描述加入 ListBox 中
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
        
        '显示开局描述
        pos.sMoveSymbolDesc = pos.sMoveSymbolDesc & GetMoveDesc(mv, False, True)
        sVar = mGetOpenDesc.GetVar(pos.sMoveSymbolDesc)
        frmChart.lblOpening.Caption = mGetOpenDesc.GetOpen(pos.sMoveSymbolDesc) _
                              & "(" & mGetOpenDesc.GetEccoNo(pos.sMoveSymbolDesc) & ")" _
                                    & IIf(sVar = "", "", "——" & sVar)
        
        '交换走子方
        mbIsRedTurn = Not mbIsRedTurn
        
        '检查重复局面
        vlRep = pos.RepStatus(3)
        
        If pos.IsMate() Then
          '如果分出胜负，那么播放胜负的声音，并且弹出提示框
          PlayResWav IDR_WIN
          MessageBoxMute "祝贺你取得胜利！"
          Xqwl.bGameOver = True
        ElseIf vlRep > 0 Then
          vlRep = pos.RepValue(vlRep)
          '注意："vlRep"是对电脑来说的分值
          PlayResWav IIf(vlRep > WIN_VALUE, IDR_LOSS, IIf(vlRep < -WIN_VALUE, IDR_WIN, IDR_DRAW))
          MessageBoxMute IIf(vlRep > WIN_VALUE, "长打作负，请不要气馁！", IIf(vlRep < -WIN_VALUE, "电脑长打作负，祝贺你取得胜利！", "双方不变作和，辛苦了！"))
          Xqwl.bGameOver = True
        ElseIf pos.nMoveNum > 100 Then
          PlayResWav (IDR_DRAW)
          MessageBoxMute "超过自然限着作和，辛苦了！"
          Xqwl.bGameOver = True
        Else
          '如果没有分出胜负，那么播放将军、吃子或一般走子的声音
          PlayResWav IIf(pos.Checked(), IDR_CHECK, IIf(pc <> 0, IDR_CAPTURE, IDR_MOVE))
          If pos.Captured() Then
            pos.SetIrrev
          End If
          
          Do While ((mbIsRedTurn And mnuComputerMoveRed.Checked) Or (Not mbIsRedTurn And mnuComputerMoveBlack.Checked)) And Not Xqwl.bGameOver
            ResponseMove  '轮到电脑走棋
            If mbIsRedTurn Then
              mMoveCount = mMoveCount + 1
            End If
            mbIsRedTurn = Not mbIsRedTurn
          Loop
        End If
      Else
        PlayResWav IDR_ILLEGAL '播放被将军的声音
      End If
    End If
    '如果根本就不符合走法(例如马不走日字)，那么程序不予理会
  End If
  
  DeleteDC Xqwl.hdcTmp
  ReleaseDC Xqwl.hWnd, Xqwl.hdc
End Sub

Private Sub Form_Paint()
  If Me.MousePointer <> vbHourglass Then
    DrawBoard Me.hdc
  End If
End Sub

'装入资源图片
Private Function LoadResBmp(nResId As Integer) As Long
  Static p As Byte
  
  Set pic(p) = LoadResPicture(nResId, 0)
  LoadResBmp = pic(p).Handle
  p = p + 1
End Function

'绘制透明图片
Private Sub DrawTransBmp(lngHdc As Long, hdcTmp As Long, xx As Integer, yy As Integer, bmp As Long)
  SelectObject hdcTmp, bmp
  TransparentBlt lngHdc, xx, yy, SQUARE_SIZE, SQUARE_SIZE, hdcTmp, 0, 0, 56, 56, MASK_COLOR
End Sub

'绘制棋盘
Private Sub DrawBoard(lngHdc As Long)
  Dim x As Integer
  Dim y As Integer
  Dim xx As Integer
  Dim yy As Integer
  Dim sq As Integer
  Dim pc As Integer
  Dim hdcTmp As Long

  '画棋盘
  hdcTmp = CreateCompatibleDC(lngHdc)
  SelectObject hdcTmp, Xqwl.bmpBoard
  TransparentBlt lngHdc, 0, 0, BOARD_WIDTH, BOARD_HEIGHT, hdcTmp, 0, 0, 520, 576, MASK_COLOR
  
  '画棋子
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
      
      '画选择框
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
      ResponseMove  '轮到电脑走棋
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
      ResponseMove  '轮到电脑走棋
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
  
  frmChart.lblOpening.Caption = "开局描述"
End Sub

'静态(Quiescence)搜索过程
Private Function SearchQuiesc(vlAlpha As Long, vlBeta As Long) As Long
  Dim i As Long
  Dim nGenMoves As Long
  Dim vl As Long
  Dim vlBest As Long
  Dim mvs(0 To MAX_GEN_MOVES - 1) As Integer

  '"GenerateMoves"参数
  Const GEN_CAPTURE As Boolean = True

  mQSNodesALayer = mQSNodesALayer + 1
  '一个静态搜索分为以下几个阶段

  '1. 检查重复局面
  vl = pos.RepStatus()
  If (vl <> 0) Then
    SearchQuiesc = pos.RepValue(vl)
    Exit Function
  End If

  '2. 到达极限深度就返回局面评价
  If (pos.nDistance = LIMIT_DEPTH) Then
    SearchQuiesc = pos.Evaluate()
    Exit Function
  End If

  '3. 初始化最佳值
  vlBest = -MATE_VALUE '这样可以知道，是否一个走法都没走过(杀棋)

  If (pos.InCheck()) Then
    '4. 如果被将军，则生成全部走法
    nGenMoves = pos.GenerateMoves(mvs)
    qsort mvs, 0, nGenMoves - 1, "History"
  Else
    '5. 如果不被将军，先做局面评价
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

    '6. 如果局面评价没有截断，再生成吃子走法
    nGenMoves = pos.GenerateMoves(mvs, GEN_CAPTURE)
    If frmSettings.chkSortByMvvLva.Value = vbChecked Then
      If frmSettings.optQuickSort.Value Then
        qsort mvs, 0, nGenMoves - 1, "MvvLva"
      Else
        SelectSort mvs, nGenMoves, "MvvLva"
      End If
    End If
  End If

  '7. 逐一走这些走法，并进行递归
  For i = 0 To nGenMoves - 1
    If (pos.MakeMove(mvs(i))) Then
      vl = -SearchQuiesc(-vlBeta, -vlAlpha)
      pos.UndoMakeMove

      '8. 进行Alpha-Beta大小判断和截断
      If (vl > vlBest) Then    '找到最佳值(但不能确定是Alpha、PV还是Beta走法)
        vlBest = vl            'vlBest 就是目前要返回的最佳值，可能超出Alpha-Beta边界
        If (vl >= vlBeta) Then '找到一个Beta走法
          SearchQuiesc = vl
          Exit Function 'Beta截断
        End If
        
        If (vl > vlAlpha) Then '找到一个PV走法
          vlAlpha = vl     '缩小Alpha-Beta边界
        End If
      End If
    End If
  Next i

  '9. 所有走法都搜索完了，返回最佳值
  SearchQuiesc = IIf(vlBest = -MATE_VALUE, pos.nDistance - MATE_VALUE, vlBest)
End Function

'超出边界(Fail-Soft)的Alpha-Beta搜索过程
Private Function SearchFull(vlAlpha As Long, vlBeta As Long, ByVal nDepth As Long, Optional bNoNull As Boolean = False) As Long
  'bInCheckExt 是用来解决将军延伸时 ListView 不在根节点处（而在子节点）添加走法造成的bug
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

  '统计访问的节点数
  mVisitNodesALayer = mVisitNodesALayer + 1
  
  '一个Alpha-Beta完全搜索分为以下几个阶段
  
  '1. 到达水平线，则调用静态搜索(注意：由于空步裁剪，深度可能小于零)
  If (nDepth <= 0) Then
    If frmSettings.chkSearchQuiesc.Value = vbChecked Then
      SearchFull = SearchQuiesc(vlAlpha, vlBeta)
    Else
      SearchFull = pos.Evaluate()
    End If
    Exit Function
  End If

  '1-1. 检查重复局面(注意：不要在根节点检查，否则就没有走法了)
  vl = pos.RepStatus()
  If (vl <> 0) Then
    SearchFull = pos.RepValue(vl)
    Exit Function
  End If

  '1-2. 到达极限深度就返回局面评价
  If (pos.nDistance = LIMIT_DEPTH) Then
    SearchFull = pos.Evaluate()
    Exit Function
  End If

  '1-3. 尝试置换表裁剪，并得到置换表走法
  If frmSettings.chkTranTable.Value = vbChecked Then
    vl = ProbeHash(vlAlpha, vlBeta, nDepth, mvHash)
    If (vl > -MATE_VALUE) Then
      SearchFull = vl
      Exit Function
    End If
  End If

  '1-4. 尝试空步裁剪(根节点的Beta值是"MATE_VALUE"，所以不可能发生空步裁剪)
  If ((Not bNoNull) And (Not pos.InCheck()) And pos.NullOkay() And (frmSettings.chkNullMove.Value = vbChecked)) Then
    pos.NullMove
    vl = -SearchFull(-vlBeta, 1 - vlBeta, nDepth - NULL_DEPTH - 1, NO_NULL)
    pos.UndoNullMove
    If (vl >= vlBeta) Then
      SearchFull = vl
      Exit Function
    End If
  End If
  
  '2. 初始化最佳值和最佳走法
  nHashFlag = HASH_ALPHA
  vlBest = -MATE_VALUE '这样可以知道，是否一个走法都没走过(杀棋)
  mvBest = 0           '这样可以知道，是否搜索到了Beta走法或PV走法，以便保存到历史表

  '3. 初始化走法排序结构
  Sort.Init mvHash
  
  '4. 逐一走这些走法，并进行递归
  mv = Sort.Next_
  
  DoEvents '让计时时钟走，同时刷新 ListView 的显示

  Do While (mv <> 0)
    If pos.MakeMove(mv) Then
      '将军延伸
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

      '5. 进行Alpha-Beta大小判断和截断
      If (vl > vlBest) Then    '找到最佳值(但不能确定是Alpha、PV还是Beta走法)
        vlBest = vl            '"vlBest"就是目前要返回的最佳值，可能超出Alpha-Beta边界
        If (vl >= vlBeta) Then '找到一个Beta走法
          nHashFlag = HASH_BETA
          mvBest = mv          'Beta走法要保存到历史表
          Exit Do              'Beta截断
        End If
        If (vl > vlAlpha) Then '找到一个PV走法
          nHashFlag = HASH_PV
          mvBest = mv          'PV走法要保存到历史表
          vlAlpha = vl         '缩小Alpha-Beta边界
        End If
      End If
    End If
    
    mv = Sort.Next_
  Loop
  
  '5. 所有走法都搜索完了，把最佳走法(不能是Alpha走法)保存到历史表，返回最佳值
  If (vlBest = -MATE_VALUE) Then
    '如果是杀棋，就根据杀棋步数给出评价
    SearchFull = pos.nDistance - MATE_VALUE
    Exit Function
  End If

  If frmSettings.chkTranTable.Value = vbChecked Then
    '记录到置换表
    RecordHash nHashFlag, vlBest, nDepth, mvBest
  End If
  
  If (mvBest <> 0) Then
    '如果不是Alpha走法，就将最佳走法保存到历史表
    SetBestMove mvBest, nDepth
  End If
  SearchFull = vlBest
End Function

'根节点的Alpha-Beta搜索过程
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
  '界面处理
  frmSearchInfo.lsvMoveList.ListItems.Clear
  frmSearchInfo.lsvMoveList.ListItems.Add , , CStr(i)
  frmSearchInfo.lsvMoveList.ListItems(i).SubItems(1) = GetMoveDesc(mv)
  frmSearchInfo.lsvDepthTimeCost.ListItems(nDepth).SubItems(1) = CStr(i) & "/" & CStr(Sort.nGenMoves)
  
  DoEvents '让计时时钟走，同时刷新 ListView 的显示

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
      
      '界面处理
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
          vlBest = vlBest + ((Rnd * Val(frmSettings.txtRANDOM_MASK.Text)) - (Rnd * Val(frmSettings.txtRANDOM_MASK.Text)))  '搜索的随机性的精华所在啊~~~
        End If
      End If
    Else '走这步棋后被将军
      frmSearchInfo.lsvMoveList.ListItems(i).SubItems(2) = CStr(-MATE_VALUE)
    End If
    mv = Sort.Next_()
    i = i + 1
    
    '界面处理
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

'对最佳走法的处理
Private Sub SetBestMove(mv As Long, nDepth As Long)
  Search.nHistoryTable(mv) = Search.nHistoryTable(mv) + (nDepth * nDepth)
  
  If (Search.mvKillers(pos.nDistance, 0) <> mv) Then
    Search.mvKillers(pos.nDistance, 1) = Search.mvKillers(pos.nDistance, 0)
    Search.mvKillers(pos.nDistance, 0) = mv
  End If
End Sub

'电脑回应一步棋
Private Sub ResponseMove()
  Dim vlRep As Long
  Dim sVar  As String

  '电脑走一步棋
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
  
  '显示开局描述
  pos.sMoveSymbolDesc = pos.sMoveSymbolDesc & GetMoveDesc(Search.mvResult, False, True)
  sVar = mGetOpenDesc.GetVar(pos.sMoveSymbolDesc)
  frmChart.lblOpening.Caption = mGetOpenDesc.GetOpen(pos.sMoveSymbolDesc) _
                        & "(" & mGetOpenDesc.GetEccoNo(pos.sMoveSymbolDesc) & ")" _
                              & IIf(sVar = "", "", "——" & sVar)
  
  '清除上一步棋的选择标记
  DrawSquare (SRC(Xqwl.mvLast))
  DrawSquare (DST(Xqwl.mvLast))
  
  '把电脑走的棋标记出来
  Xqwl.mvLast = Search.mvResult
  DrawSquare SRC(Xqwl.mvLast), DRAW_SELECTED
  DrawSquare DST(Xqwl.mvLast), DRAW_SELECTED
  
  '检查重复局面
  vlRep = pos.RepStatus(3)
  
  If (pos.IsMate()) Then
    '如果分出胜负，那么播放胜负的声音，并且弹出不带声音的提示框
    PlayResWav (IDR_LOSS)
    MessageBoxMute "请再接再厉！"
    Xqwl.bGameOver = True
  ElseIf vlRep > 0 Then
    vlRep = pos.RepValue(vlRep)
    '注意："vlRep"是对玩家来说的分值
    PlayResWav IIf(vlRep < -WIN_VALUE, IDR_LOSS, IIf(vlRep > WIN_VALUE, IDR_WIN, IDR_DRAW))
    MessageBoxMute IIf(vlRep < -WIN_VALUE, "长打作负，请不要气馁！", IIf(vlRep > WIN_VALUE, "电脑长打作负，祝贺你取得胜利！", "双方不变作和，辛苦了！"))
    Xqwl.bGameOver = True
  ElseIf pos.nMoveNum > 100 Then
    PlayResWav (IDR_DRAW)
    MessageBoxMute "超过自然限着作和，辛苦了！"
    Xqwl.bGameOver = True
  Else
    '如果没有分出胜负，那么播放将军、吃子或一般走子的声音
    PlayResWav IIf(pos.Checked(), IDR_CHECK2, IIf(pos.Captured, IDR_CAPTURE2, IDR_MOVE2))
    If pos.Captured() Then
      pos.SetIrrev
    End If
  End If
End Sub

'迭代加深搜索过程
Private Sub SearchMain()
  Dim i               As Long
  Dim vl              As Long
  Dim TimeElapsed     As Long
  Dim VisitNodesTotal As Long
  Dim mvs(0 To MAX_GEN_MOVES - 1) As Integer
  Dim nGenMoves       As Long
  Dim c               As Byte

  '初始化
  For i = 0 To 65535
    Search.nHistoryTable(i) = 0 '清空历史表
  Next i
  
  For i = 0 To LIMIT_DEPTH - 1
    Search.mvKillers(i, 0) = 0 '清空杀手走法表
    Search.mvKillers(i, 1) = 0 '清空杀手走法表
  Next i
  
  For i = 0 To HASH_SIZE - 1 '清空置换表
    Search_HashTable(i).dwLock0 = 0
    Search_HashTable(i).dwLock1 = 0
    Search_HashTable(i).svl = 0
    Search_HashTable(i).ucDepth = 0
    Search_HashTable(i).ucFlag = 0
    Search_HashTable(i).wmv = 0
    Search_HashTable(i).wReserved = 0
  Next i
  
  mSearchStartTime = GetTickCount()  '初始化定时器
  pos.nDistance = 0 '初始步数
  
  frmSearchInfo.lsvDepthTimeCost.ListItems.Clear
  tmrSearch.Enabled = True
  VisitNodesTotal = 0
  
  If frmSettings.chkOpeningBook.Value = vbChecked Then
    '搜索开局库
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
    
    '检查是否只有唯一走法
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
  
  '迭代加深过程
  For i = 1 To LIMIT_DEPTH
    frmSearchInfo.lsvDepthTimeCost.ListItems.Add CStr(i), , CStr(i)
    
    mVisitNodesALayer = 0
    mQSNodesALayer = 0
    
    '搜索根节点
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
    
    '搜索到杀棋，就终止搜索
    If (vl > WIN_VALUE Or vl < -WIN_VALUE) Then
      Exit For
    End If
    
    If frmSettings.optLimitTime.Value Then
      '超过时限，就终止搜索
      If (TimeElapsed > Val(frmSettings.txtTimeLimit.Text)) Then
        Exit For
      End If
    Else
      '达到深度，就终止搜索
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
    .Filter = "可移植棋谱 (*.PGN)|*.pgn"
    .FileName = Replace(CStr(Now), ":", "：")
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
  frmSearchInfo.lblSearchSeconds.Caption = "已搜索时间：" & CStr(TimeElapsed) & " 秒"
End Sub

'弹出不带声音的提示框
Private Sub MessageBoxMute(lpszText As String)
  Dim mbp As MSGBOXPARAMS

  mbp.cbSize = Len(mbp)
  mbp.hwndOwner = Xqwl.hWnd
  mbp.lpszText = lpszText
  mbp.lpszCaption = "VB AI 中国象棋"
  mbp.dwStyle = MB_USERICON
  mbp.lpszIcon = 104
  MessageBoxIndirect mbp
End Sub



