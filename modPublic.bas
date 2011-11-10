Attribute VB_Name = "modPublic"
Option Explicit

'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)

Public pos As New PositionStruct

'棋子编号
Public Const PIECE_KING = 0
Public Const PIECE_ADVISOR = 1
Public Const PIECE_BISHOP = 2
Public Const PIECE_KNIGHT = 3
Public Const PIECE_ROOK = 4
Public Const PIECE_CANNON = 5
Public Const PIECE_PAWN = 6

'其他常数
Public Const MAX_GEN_MOVES As Integer = 128   '最大的生成走法数
Public Const MAX_MOVES As Integer = 256       '最大的历史走法数
Public Const MATE_VALUE As Integer = 10000    '最高分值，即将死的分值
Public Const ADVANCED_VALUE  As Byte = 3      '先行权分值
Public Const LIMIT_DEPTH  As Byte = 32        '最大的搜索深度
Public Const WIN_VALUE  As Integer = MATE_VALUE - 100 '搜索出胜负的分值界限，超出此值就说明已经搜索出杀棋了
Public Const NULL_MARGIN As Integer = 400             '空步裁剪的子力边界
Public Const NULL_DEPTH  As Byte = 2                  '空步裁剪的裁剪深度
Public Const DRAW_VALUE As Byte = 20     '和棋时返回的分数(取负值)

Public Const HASH_SIZE As Long = 2 ^ 20  ' 置换表大小
Public Const HASH_ALPHA As Long = 1      ' ALPHA节点的置换表项
Public Const HASH_BETA As Long = 2       ' BETA节点的置换表项
Public Const HASH_PV As Long = 3         ' PV节点的置换表项
Public Const BOOK_SIZE As Long = 16384   ' 开局库大小
Public Const c_SizeofBookItem As Byte = 8 '开局库中一项的字节数

'播放声音
Public Const SND_ASYNC = &H1
Public Const SND_NOSTOP = &H10
Public Const SND_NODEFAULT = &H2
Public Const SND_NOWAIT = &H2000
Public Const SND_SYNC = &H0
Public Const SND_RESOURCE = &H40004
Public Const SND_MEMORY = &H4     '指向一个内存文件
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Declare Function GetTickCount Lib "KERNEL32" () As Long

'仕(士)的步长
Public ccAdvisorDelta As Variant

'马的步长，以帅(将)的步长作为马腿
Public ccKnightDelta As Variant

'马被将军的步长，以仕(士)的步长作为马腿
Public ccKnightCheckDelta As Variant

'帅(将)的步长
Public ccKingDelta As Variant

'判断棋子是否在棋盘中的数组
Public ccInBoard As Variant

'判断棋子是否在九宫的数组
Public ccInFort As Variant
 
'判断步长是否符合特定走法的数组，1=帅(将)，2=仕(士)，3=相(象)
Public ccLegalSpan As Variant

'根据步长判断马是否蹩腿的数组
Public ccKnightPin As Variant

'子力位置价值表
Public cucvlPiecePos As Variant

'MVV/LVA每种子力的价值
Public cucMvvLva As Variant

Public Search As New clsSearch

Public Type udtZobrist
  Player As New ZobristStruct
  Table(0 To 13, 0 To 255) As New ZobristStruct
End Type

Public Zobrist As udtZobrist

'置换表项结构
Public Type HashItem
  ucDepth   As Byte
  ucFlag    As Byte
  svl       As Integer
  wmv       As Integer
  wReserved As Long
  dwLock0   As Long
  dwLock1   As Long
End Type

'开局库项结构
Public Type BookItem
  dwLock As Long
  wmv    As Integer
  wvl    As Integer
End Type

Public Search_HashTable(0 To HASH_SIZE - 1) As HashItem '置换表
Public Search_BookTable(0 To BOOK_SIZE - 1) As BookItem '开局库

'走法排序阶段
Public Const PHASE_HASH      As Long = 0
Public Const PHASE_KILLER_1  As Long = 1
Public Const PHASE_KILLER_2  As Long = 2
Public Const PHASE_GEN_MOVES As Long = 3
Public Const PHASE_REST      As Long = 4

Public Sub InitConstantArray()
  ccInBoard = Array( _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      
  ccInFort = Array( _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      
  ccLegalSpan = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 3, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 2, 1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 2, 1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 3, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    
  ccKnightPin = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, -16, 0, -16, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, -1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, -1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 16, 0, 16, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
  
  ccKingDelta = Array(-16, -1, 1, 16)
  
  ccAdvisorDelta = Array(-17, -15, 15, 17)
  
  ccKnightDelta = Array(Array(-33, -31), Array(-18, 14), Array(-14, 18), Array(31, 33))
  
  ccKnightCheckDelta = Array(Array(-33, -18), Array(-31, -14), Array(14, 31), Array(18, 33))
  

  cucvlPiecePos = Array( _
    Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 11, 15, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
    Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 20, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 23, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
    Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 18, 0, 0, 0, 23, 0, 0, 0, 18, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
    Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 90, 90, 90, 96, 90, 96, 90, 90, 90, 0, 0, 0, 0, 0, 0, 0, 90, 96, 103, 97, 94, 97, 103, 96, 90, 0, 0, 0, 0, 0, 0, 0, 92, 98, 99, 103, 99, 103, 99, 98, 92, 0, 0, 0, 0, 0, 0, 0, 93, 108, 100, 107, 100, 107, 100, 108, 93, 0, 0, 0, 0, 0, 0, 0, 90, 100, 99, 103, 104, 103, 99, 100, 90, 0, 0, 0, 0, _
      0, 0, 0, 90, 98, 101, 102, 103, 102, 101, 98, 90, 0, 0, 0, 0, 0, 0, 0, 92, 94, 98, 95, 98, 95, 98, 94, 92, 0, 0, 0, 0, 0, 0, 0, 93, 92, 94, 95, 92, 95, 94, 92, 93, 0, 0, 0, 0, 0, 0, 0, 85, 90, 92, 93, 78, 93, 92, 90, 85, 0, 0, 0, 0, _
      0, 0, 0, 88, 85, 90, 88, 90, 88, 90, 85, 88, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
    Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 206, 208, 207, 213, 214, 213, 207, 208, 206, 0, 0, 0, 0, 0, 0, 0, 206, 212, 209, 216, 233, 216, 209, 212, 206, 0, 0, 0, 0, 0, 0, 0, 206, 208, 207, 214, 216, 214, 207, 208, 206, 0, 0, 0, 0, _
      0, 0, 0, 206, 213, 213, 216, 216, 216, 213, 213, 206, 0, 0, 0, 0, 0, 0, 0, 208, 211, 211, 214, 215, 214, 211, 211, 208, 0, 0, 0, 0, 0, 0, 0, 208, 212, 212, 214, 215, 214, 212, 212, 208, 0, 0, 0, 0, 0, 0, 0, 204, 209, 204, 212, 214, 212, 204, 209, 204, 0, 0, 0, 0, 0, 0, 0, 198, 208, 204, 212, 212, 212, 204, 208, 198, 0, 0, 0, 0, 0, 0, 0, 200, 208, 206, 212, 200, 212, 206, 208, 200, 0, 0, 0, 0, 0, 0, 0, 194, 206, 204, 212, 200, 212, 204, 206, 194, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
    Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 100, 96, 91, 90, 91, 96, 100, 100, 0, 0, 0, 0, 0, 0, 0, 98, 98, 96, 92, 89, 92, 96, 98, 98, 0, 0, 0, 0, 0, 0, 0, 97, 97, 96, 91, 92, 91, 96, 97, 97, 0, 0, 0, 0, 0, 0, 0, 96, 99, 99, 98, 100, 98, 99, 99, 96, 0, 0, 0, 0, 0, 0, 0, 96, 96, 96, 96, 100, 96, 96, 96, 96, 0, 0, 0, 0, 0, 0, 0, 95, 96, 99, 96, 100, 96, 99, 96, 95, 0, 0, 0, 0, 0, 0, 0, 96, 96, 96, 96, 96, 96, 96, 96, 96, 0, 0, 0, 0, _
      0, 0, 0, 97, 96, 100, 99, 101, 99, 100, 96, 97, 0, 0, 0, 0, 0, 0, 0, 96, 97, 98, 98, 98, 98, 98, 97, 96, 0, 0, 0, 0, 0, 0, 0, 96, 96, 97, 99, 99, 99, 97, 96, 96, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
    Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 9, 9, 9, 11, 13, 11, 9, 9, 9, 0, 0, 0, 0, 0, 0, 0, 19, 24, 34, 42, 44, 42, 34, 24, 19, 0, 0, 0, 0, 0, 0, 0, 19, 24, 32, 37, 37, 37, 32, 24, 19, 0, 0, 0, 0, _
      0, 0, 0, 19, 23, 27, 29, 30, 29, 27, 23, 19, 0, 0, 0, 0, 0, 0, 0, 14, 18, 20, 27, 29, 27, 20, 18, 14, 0, 0, 0, 0, 0, 0, 0, 7, 0, 13, 0, 16, 0, 13, 0, 7, 0, 0, 0, 0, 0, 0, 0, 7, 0, 7, 0, 15, 0, 7, 0, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))

  cucMvvLva = Array(0, 0, 0, 0, 0, 0, 0, 0, _
                    5, 1, 1, 3, 4, 3, 2, 0, _
                    5, 1, 1, 3, 4, 3, 2, 0)
End Sub

'获得走法的起点
Public Function SRC(ByVal mv As Long) As Long
  SRC = mv And &HFF '低8位
End Function

'获得走法的终点
Public Function DST(ByVal mv As Integer) As Byte
 ' CopyMemory DST, ByVal (VarPtr(mv) + 1), 1 '高8位
  DST = (mv And &HFF00&) \ &H100
End Function

'获得红黑标记(红子是8，黑子是16)
Public Function SIDE_TAG(sd As Byte) As Byte
  SIDE_TAG = IIf(sd = 0, 8, 16)
End Function

'获得对方红黑标记
Public Function OPP_SIDE_TAG(sd As Byte) As Integer
  OPP_SIDE_TAG = IIf(sd = 0, 16, 8)
End Function

'格子水平镜像
Public Function SQUARE_FORWARD(ByVal sq As Integer, sd As Byte)
  SQUARE_FORWARD = sq - 16 + (sd * 2 ^ 5)
End Function

'判断棋子是否在棋盘中
Public Function IN_BOARD(ByVal sq As Integer) As Boolean
  IN_BOARD = CBool(ccInBoard(sq) <> 0)
End Function

'判断棋子是否在九宫中
Public Function IN_FORT(ByVal sq As Long) As Boolean
  IN_FORT = CBool(ccInFort(sq) <> 0)
End Function

'获得格子的纵坐标
Public Function RANK_Y(sq As Integer) As Integer
  RANK_Y = sq \ 2 ^ 4
End Function

'获得格子的横坐标
Public Function FILE_X(sq As Integer) As Integer
  FILE_X = sq And &HF
End Function

'根据纵坐标和横坐标获得格子
Public Function COORD_XY(ByVal x As Integer, ByVal y As Integer) As Integer
  COORD_XY = x + (y * 2 ^ 4)
End Function

'翻转格子
Public Function SQUARE_FLIP(sq As Byte) As Integer
  SQUARE_FLIP = 254 - sq
End Function

'横坐标水平镜像
Public Function FILE_FLIP(x As Integer) As Integer
  FILE_FLIP = 14 - x
End Function

'纵坐标垂直镜像
Public Function RANK_FLIP(y As Integer) As Integer
  RANK_FLIP = 15 - y
End Function

'格子水平镜像
Public Function MIRROR_SQUARE(sq As Integer)
  MIRROR_SQUARE = COORD_XY(FILE_FLIP(FILE_X(sq)), RANK_Y(sq))
End Function


'走法是否符合帅(将)的步长
Public Function KING_SPAN(sqSrc As Integer, sqDst As Integer) As Boolean
  KING_SPAN = ccLegalSpan(sqDst - sqSrc + 256) = 1
End Function

'走法是否符合仕(士)的步长
Public Function ADVISOR_SPAN(sqSrc As Integer, sqDst As Integer) As Boolean
  ADVISOR_SPAN = (ccLegalSpan(sqDst - sqSrc + 256) = 2)
End Function

'走法是否符合相(象)的步长
Public Function BISHOP_SPAN(sqSrc As Integer, sqDst As Integer) As Boolean
  BISHOP_SPAN = (ccLegalSpan(sqDst - sqSrc + 256) = 3)
End Function

'相(象)眼的位置
Public Function BISHOP_PIN(sqSrc As Integer, sqDst As Integer) As Integer
  BISHOP_PIN = (sqSrc + sqDst) \ 2
End Function

'马腿的位置
Public Function KNIGHT_PIN(sqSrc As Integer, sqDst As Integer) As Integer
  KNIGHT_PIN = sqSrc + ccKnightPin(sqDst - sqSrc + 256)
End Function

'是否未过河
Public Function HOME_HALF(ByVal sq As Integer, ByVal sd As Integer) As Boolean
  HOME_HALF = (sq And &H80) <> (sd * 2 ^ 7)
End Function

'是否已过河
Public Function AWAY_HALF(ByVal sq As Integer, ByVal sd As Integer) As Boolean
  AWAY_HALF = (sq And &H80) = (sd * 2 ^ 7)
End Function

'是否在河的同一边
Public Function SAME_HALF(sqSrc As Integer, sqDst As Integer) As Boolean
  SAME_HALF = ((sqSrc Xor sqDst) And &H80) = 0
End Function

'是否在同一行
Public Function SAME_RANK(sqSrc As Integer, sqDst As Integer) As Boolean
  SAME_RANK = ((sqSrc Xor sqDst) And &HF0) = 0
End Function

'是否在同一列
Public Function SAME_FILE(sqSrc As Integer, sqDst As Integer) As Boolean
  SAME_FILE = ((sqSrc Xor sqDst) And &HF) = 0
End Function

'根据起点和终点获得走法
Public Function MOVE_(ByVal sqSrc As Byte, ByVal sqDst As Byte) As Integer
  CopyMemory MOVE_, sqSrc, 1
  CopyMemory ByVal (VarPtr(MOVE_) + 1), sqDst, 1
End Function

'走法水平镜像
Public Function MIRROR_MOVE(mv As Integer) As Integer
  MIRROR_MOVE = MOVE_(MIRROR_SQUARE(SRC(mv)), MIRROR_SQUARE(DST(mv)))
End Function
 
'播放资源声音
Public Sub PlayResWav(nResId As Integer)
  Dim WavData() As Byte

  WavData = LoadResData(nResId, "WAVE")
  sndPlaySound WavData(0), SND_SYNC Or SND_NOWAIT Or SND_MEMORY
End Sub

'按历史表排序的比较函数
Public Function CompareHistory(lpmv1 As Integer, lpmv2 As Integer) As Boolean
  If frmSettings.optGreateOrEqual.Value Then
    CompareHistory = Search.nHistoryTable(lpmv2) - Search.nHistoryTable(lpmv1) >= 0
  Else
    CompareHistory = Search.nHistoryTable(lpmv2) - Search.nHistoryTable(lpmv1) > 0
  End If
End Function

'按MVV/LVA值排序的比较函数
Public Function CompareMvvLva(lpmv1 As Integer, lpmv2 As Integer) As Boolean
  If frmSettings.optGreateOrEqual.Value Then
    CompareMvvLva = MvvLva(lpmv2) - MvvLva(lpmv1) >= 0
  Else
    CompareMvvLva = MvvLva(lpmv2) - MvvLva(lpmv1) > 0
  End If
End Function

Public Sub qsort(ByRef mvs() As Integer, ByVal lngLeft As Long, ByVal lngRight As Long, ByVal sAccording As String)
  Dim i    As Long
  Dim j    As Long
  Dim x    As Long
  Dim Temp As Long
  
  '设要排序的数组是A[lngLeft]……A[lngRight]，首先任意选取一个数据（通常选用第一个数据）作为关键数据，
  '然后将所有比它小的数都放到它前面，所有比它大的数都放到它后面，这个过程称为一趟快速排序。
  '一趟快速排序的算法是：
  
  '1）设置两个变量i、j，排序开始的时候：i = lngLeft，j = lngRight；
  i = lngLeft
  j = lngRight
  
  '2）以第一个数组元素作为关键数据，赋值给X，即 X=A[lngLeft]；
  x = mvs(lngLeft)
  
  Do While i < j
    '3）从j开始向前搜索，即由后开始向前搜索（j=j-1），找到第一个小于X的值，让该值与X交换（找到就行.找到后i大小不变）；
    If sAccording = "History" Then
      Do While (CompareHistory(mvs(j), mvs(i)) And i < j)
        j = j - 1
      Loop
    Else
      Debug.Assert sAccording = "MvvLva"
      Do While (CompareMvvLva(mvs(j), mvs(i)) And i < j)
        j = j - 1
      Loop
    End If
    
    If i < j Then
      Temp = mvs(i)
      mvs(i) = mvs(j)
      mvs(j) = Temp
      i = i + 1
    End If
    
    '4）从i开始向后搜索，即由前开始向后搜索（i=i+1），找到第一个大于X的值，让该值与X交换（找到就行.找到后j大小不变）；
    If sAccording = "History" Then
      Do While (CompareHistory(mvs(j), mvs(i)) And i < j)
        i = i + 1
      Loop
    Else
      Debug.Assert sAccording = "MvvLva"
      Do While (CompareMvvLva(mvs(j), mvs(i)) And i < j)
        j = j - 1
      Loop
    End If
    
    If i < j Then
      Temp = mvs(j)
      mvs(j) = mvs(i)
      mvs(i) = Temp
      j = j - 1
    End If
    '5）重复第3、4步，直到 i=j； (3,4步是在程序中没找到时候j=j-1，i=i+1。
    '找到并交换的时候i， j指针位置不变。另外当i=j这过程一定正好是i+1或j-1完成的最后令循环结束)
  Loop
  
  '快速排序就是递归调用此过程——以关键数据（通常选用第一个数据）为中点分割这个数据序列，分别对前面一部分和后面一部分进行类似的快速排序，
  '从而完成全部数据序列的快速排序，最后把此数据序列变成一个有序的序列
  If lngLeft < i - 1 Then
    qsort mvs, lngLeft, i - 1, sAccording
  End If
  
  If lngRight > i + 1 Then
    qsort mvs, i + 1, lngRight, sAccording
  End If
End Sub

Public Sub SelectSort(ByRef mvs() As Integer, ByVal n As Long, ByVal sAccording As String)
  Dim i    As Integer
  Dim j    As Integer
  Dim Temp As Long
  
  If sAccording = "History" Then
    For i = 0 To n - 1
      For j = i + 1 To n
        If CompareHistory(mvs(i), mvs(j)) Then  'j-i>0
          Temp = mvs(i)
          mvs(i) = mvs(j)
          mvs(j) = Temp
        End If
      Next j
    Next i
  Else
    Debug.Assert sAccording = "MvvLva"
    For i = 0 To n - 1
      For j = i + 1 To n
        If CompareMvvLva(mvs(i), mvs(j)) Then   'j-i>0
          Temp = mvs(i)
          mvs(i) = mvs(j)
          mvs(j) = Temp
        End If
      Next j
    Next i
  End If
End Sub

Public Function Pc2Name(ByVal pc As Byte, ByVal bSymbolMode As Boolean) As String
  If bSymbolMode Then
    Select Case pc
      Case 8, 16
        Pc2Name = "K"
      Case 9, 17
        Pc2Name = "A"
      Case 10, 18
        Pc2Name = "B"
      Case 11, 19
        Pc2Name = "N"
      Case 12, 20
        Pc2Name = "R"
      Case 13, 21
        Pc2Name = "C"
      Case 14, 22
        Pc2Name = "P"
    End Select
  Else
    Select Case pc
      Case 8
        Pc2Name = "帅"
      Case 9
        Pc2Name = "仕"
      Case 10
        Pc2Name = "相"
      Case 11
        Pc2Name = "马"
      Case 12
        Pc2Name = "车"
      Case 13
        Pc2Name = "炮"
      Case 14
        Pc2Name = "兵"
    
    
      Case 16
        Pc2Name = "将"
      Case 17
        Pc2Name = "士"
      Case 18
        Pc2Name = "象"
      Case 19
        Pc2Name = "马"
      Case 20
        Pc2Name = "车"
      Case 21
        Pc2Name = "炮"
      Case 22
        Pc2Name = "卒"
        
      Case Else
        Pc2Name = "Error"
    End Select
  End If
End Function

'求MVV/LVA值
Public Function MvvLva(mv As Integer) As Long
  MvvLva = (cucMvvLva(pos.ucpcSquares(DST(mv))) * 8) - cucMvvLva(pos.ucpcSquares(SRC(mv)))
End Function

Public Sub InitZobrist()
  Dim i As Integer
  Dim j As Integer
  Dim rc4 As New RC4Struct
  
  rc4.InitZero
  Zobrist.Player.InitRC4 rc4
  
  For i = 0 To 13
    For j = 0 To 255
      Zobrist.Table(i, j).InitRC4 rc4
    Next j
  Next i
End Sub

'提取置换表项
Public Function ProbeHash(vlAlpha As Long, vlBeta As Long, nDepth As Long, ByRef mv As Integer) As Long
  Dim bMate As Boolean '杀棋标志：如果是杀棋，那么不需要满足深度条件
  Dim hsh   As HashItem
  
  hsh = Search_HashTable(pos.zobr.dwKey And (HASH_SIZE - 1))
  If (hsh.dwLock0 <> pos.zobr.dwLock0 Or hsh.dwLock1 <> pos.zobr.dwLock1) Then
    mv = 0
    ProbeHash = -MATE_VALUE
    Exit Function
  End If

  mv = hsh.wmv
  bMate = False
  
  If (hsh.svl > WIN_VALUE) Then
    hsh.svl = hsh.svl - (pos.nDistance)
    bMate = True
  ElseIf (hsh.svl < -WIN_VALUE) Then
    hsh.svl = hsh.svl + (pos.nDistance)
    bMate = True
  End If
  
  If (hsh.ucDepth >= nDepth Or bMate) Then
    If (hsh.ucFlag = HASH_BETA) Then
      ProbeHash = IIf(hsh.svl >= vlBeta, hsh.svl, -MATE_VALUE)
      Exit Function
    ElseIf (hsh.ucFlag = HASH_ALPHA) Then
      ProbeHash = IIf(hsh.svl <= vlAlpha, hsh.svl, -MATE_VALUE)
      Exit Function
    End If

    ProbeHash = hsh.svl
    Exit Function
  End If
  ProbeHash = -MATE_VALUE
End Function

'保存置换表项
Public Sub RecordHash(nFlag As Long, vl As Long, nDepth As Long, mv As Long)
  Dim hsh As HashItem
  
  hsh = Search_HashTable(pos.zobr.dwKey And (HASH_SIZE - 1))
  
  If (hsh.ucDepth > nDepth) Then
    Exit Sub
  End If
  
  hsh.ucFlag = nFlag
  hsh.ucDepth = nDepth
  
  If (vl > WIN_VALUE) Then
    hsh.svl = vl + pos.nDistance
  ElseIf (vl < -WIN_VALUE) Then
    hsh.svl = vl - pos.nDistance
  Else
    hsh.svl = vl
  End If
  
  hsh.wmv = mv
  hsh.dwLock0 = pos.zobr.dwLock0
  hsh.dwLock1 = pos.zobr.dwLock1
  Search_HashTable(pos.zobr.dwKey And (HASH_SIZE - 1)) = hsh
End Sub

'搜索开局库
Public Function SearchBook() As Long
  Dim i          As Long
  Dim vl         As Long
  Dim nBookMoves As Long
  Dim mv         As Integer
  Dim mvs(0 To MAX_GEN_MOVES - 1) As Long
  Dim vls(0 To MAX_GEN_MOVES - 1) As Long
  Dim bMirror    As Boolean
  Dim bkToSearch As BookItem
  Dim lpbk       As Integer
  Dim posMirror  As New PositionStruct
  
  
  '搜索开局库的过程有以下几个步骤

  '1. 如果没有开局库，则立即返回
  If (Search.nBookSize = 0) Then
    SearchBook = 0
    Exit Function
  End If

  '2. 搜索当前局面
  bMirror = False
  bkToSearch.dwLock = pos.zobr.dwLock1
  lpbk = bsearch(bkToSearch.dwLock)
  
  '3. 如果没有找到，那么搜索当前局面的镜像局面
  If lpbk = -1 Then
    bMirror = True
    pos.Mirror posMirror
    bkToSearch.dwLock = posMirror.zobr.dwLock1
    lpbk = bsearch(bkToSearch.dwLock)
  End If
  
  '4. 如果镜像局面也没找到，则立即返回
  If lpbk = -1 Then
    SearchBook = 0
    Exit Function
  End If

  '5. 如果找到，则向前查第一个开局库项
  Do While (lpbk >= 0 And Search_BookTable(lpbk).dwLock = bkToSearch.dwLock)
    lpbk = lpbk - 1
  Loop
  
  lpbk = lpbk + 1
  
  '6. 把走法和分值写入到"mvs"和"vls"数组中
  vl = 0
  nBookMoves = 0
  frmSearchInfo.lsvMoveList.ListItems.Clear
  Do While (lpbk < Search.nBookSize And Search_BookTable(lpbk).dwLock = bkToSearch.dwLock)
    mv = IIf(bMirror, MIRROR_MOVE(Search_BookTable(lpbk).wmv), Search_BookTable(lpbk).wmv)
    If (pos.LegalMove(mv)) Then
      mvs(nBookMoves) = mv
      vls(nBookMoves) = Search_BookTable(lpbk).wvl
      vl = vl + vls(nBookMoves)
      nBookMoves = nBookMoves + 1
      frmSearchInfo.lsvMoveList.ListItems.Add , , CStr(nBookMoves)
      frmSearchInfo.lsvMoveList.ListItems(nBookMoves).SubItems(1) = GetMoveDesc(mv)
      frmSearchInfo.lsvMoveList.ListItems(nBookMoves).SubItems(2) = CStr(vls(nBookMoves - 1))
      If (nBookMoves = MAX_GEN_MOVES) Then
        Exit Do  '防止"BOOK.DAT"中含有异常数据
      End If
    End If
    lpbk = lpbk + 1
  Loop
  
  If (vl = 0) Then
    SearchBook = 0
    Exit Function '防止"BOOK.DAT"中含有异常数据
  End If
  
  '7. 根据权重随机选择一个走法
  vl = CInt(Rnd * 32767) Mod vl
  For i = 0 To nBookMoves - 1
    vl = vl - vls(i)
    If (vl < 0) Then
      Exit For
    End If
  Next i
  
  frmSearchInfo.lsvValue.ListItems.Add , , CStr(frmSearchInfo.lsvValue.ListItems.Count + 1)
  frmSearchInfo.lsvValue.ListItems(frmSearchInfo.lsvValue.ListItems.Count).SubItems(1) = "-"
  frmSearchInfo.lsvValue.ListItems(frmSearchInfo.lsvValue.ListItems.Count).SubItems(2) = "-"
  frmSearchInfo.lsvValue.ListItems(frmSearchInfo.lsvValue.ListItems.Count).SubItems(4) = "开局库"
  SearchBook = mvs(i)
End Function

Public Function GetMoveDesc(ByVal mv As Integer, Optional bSimulateMove As Boolean = True, Optional bSymbolMode As Boolean = False) As String
  Dim SrcBlk As Byte
  Dim DstBlk As Byte
  Dim SrcX As Byte  '绝对坐标（黑棋时两坐标相等）
  Dim SrcX2 As Byte '相对坐标（红棋从右往左数）
  Dim SrcY As Byte
  Dim DstX As Byte
  Dim DstY As Byte
  Dim pc As Byte  '(包含有棋子颜色信息)
  Dim pc2 As Byte '(脱去了棋子颜色信息)
  Dim sAct As String
  Dim sStepDstX As String
  Dim bIsRed As Boolean
  Dim y As Byte
  Dim ScanPc As Byte
  Dim sWhich As String
  Dim bCondition1 As Boolean
  Dim bCondition2 As Boolean
  
  DstBlk = (mv And &HFF00&) \ &H100
  'CopyMemory DstBlk, ByVal (VarPtr(mv) + 1), 1 '高8位
  DstX = (DstBlk And &HF) - 2
  DstY = DstBlk \ &H10 - 2
  SrcBlk = mv And &HFF
  SrcX = (SrcBlk And &HF) - 2
  SrcY = SrcBlk \ &H10 - 2
  
  If bSimulateMove Then
    pc = pos.ucpcSquares(SrcBlk)
  Else
    pc = pos.ucpcSquares(DstBlk)
  End If
  
  bIsRed = (pc And &H10) = 0
  pc2 = IIf(bIsRed, pc - 8, pc - 16)
  SrcX2 = IIf(bIsRed, 10 - SrcX, SrcX)
  
  sWhich = ""
  If pc2 <> PIECE_BISHOP And pc2 <> PIECE_ADVISOR Then '如果出现“前象退一”或“后仕进五”这样的棋谱会被人笑掉大牙
    For y = 1 To 10
      ScanPc = pos.ucpcSquares(COORD_XY(SrcX + 2, y + 2))
      '重合条件1：在该列找到相同的棋子
      bCondition1 = CBool(ScanPc = pc)
      
      '如果是马并且不是思考时的模拟移动，只需满足重合条件1即可，即重合条件2永远满足；否则（如果不是马或者是模拟移动）还要判断在该列找到的棋子是不是自己
      bCondition2 = IIf(pc2 = PIECE_KNIGHT And Not bSimulateMove, True, CBool(y <> IIf(bSimulateMove, SrcY, DstY)))
      
      If bCondition1 And bCondition2 Then
        If bIsRed Then
          sWhich = IIf(y > SrcY, "前", "后")
        Else
          sWhich = IIf(y > SrcY, "后", "前")
        End If
        Exit For
      End If
    Next y
  End If
  
  Select Case pc2
    Case 0, 4, 5, 6 '直行类
      If DstY > SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "-", "退")
        Else
          sAct = IIf(bSymbolMode, "+", "进")
        End If
        sStepDstX = IIf(bSymbolMode, CStr(DstY - SrcY), CCStr(DstY - SrcY, bIsRed))
      ElseIf DstY < SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "+", "进")
        Else
          sAct = IIf(bSymbolMode, "-", "退")
        End If
        sStepDstX = IIf(bSymbolMode, CStr(SrcY - DstY), CCStr(SrcY - DstY, bIsRed))
      Else
        sAct = IIf(bSymbolMode, ".", "平")
        DstX = IIf(bIsRed, 10 - DstX, DstX)
        sStepDstX = IIf(bSymbolMode, CStr(DstX), CCStr(DstX, bIsRed))
      End If
    Case 1, 2, 3 '斜行类
      If DstY > SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "-", "退")
        Else
          sAct = IIf(bSymbolMode, "+", "进")
        End If
      ElseIf DstY < SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "+", "进")
        Else
          sAct = IIf(bSymbolMode, "-", "退")
        End If
      End If
      DstX = IIf(bIsRed, 10 - DstX, DstX)
      sStepDstX = IIf(bSymbolMode, CStr(DstX), CCStr(DstX, bIsRed))
  End Select
  
  If sWhich = "" Then
    GetMoveDesc = Pc2Name(pc, bSymbolMode) & IIf(bSymbolMode, CStr(SrcX2), CCStr(SrcX2, bIsRed)) & sAct & sStepDstX
  Else
    GetMoveDesc = sWhich & Pc2Name(pc, bSymbolMode) & sAct & sStepDstX
  End If
End Function

Private Function CCStr(ByVal n As Byte, ByVal bIsRed As Boolean) As String
  Const sChiNum As String = "一二三四五六七八九"
  Const sAlbNum As String = "１２３４５６７８９"
  
  CCStr = Mid$(IIf(bIsRed, sChiNum, sAlbNum), n, 1)
End Function

Private Function Unsign(ByVal n As Long) As Double
  If n >= 0 Then
    Unsign = n
  Else
    Unsign = n + 2 ^ 32
  End If
End Function

Public Function bsearch(ByVal NumToSearch As Double) As Integer
  Dim lo As Long
  Dim hi As Long
  Dim mi As Long
  
  lo = 0
  hi = Search.nBookSize - 1
  
  NumToSearch = Unsign(NumToSearch)
  
  Do While (lo + 1 < hi)
    mi = (lo + hi) \ 2
    If Unsign(Search_BookTable(mi).dwLock) = NumToSearch Then
      bsearch = mi
      Exit Function
    ElseIf Unsign(Search_BookTable(mi).dwLock) < NumToSearch Then
      lo = mi
    Else
      Debug.Assert Unsign(Search_BookTable(mi).dwLock) > NumToSearch
      hi = mi
    End If
  Loop
  bsearch = -1
End Function

'装入开局库
Public Sub LoadBook()
  Dim a() As Byte
  Dim i As Long
  
  a = LoadResData("BOOK_DATA", 10)
  Search.nBookSize = (UBound(a) + 1) / c_SizeofBookItem
  If (Search.nBookSize > BOOK_SIZE) Then
    Search.nBookSize = BOOK_SIZE
  End If
  
  CopyMemory Search_BookTable(0), a(0), (UBound(a) + 1)
End Sub
