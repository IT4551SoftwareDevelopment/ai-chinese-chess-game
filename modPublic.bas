Attribute VB_Name = "modPublic"
Option Explicit

'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)

Public pos As New PositionStruct

'���ӱ��
Public Const PIECE_KING = 0
Public Const PIECE_ADVISOR = 1
Public Const PIECE_BISHOP = 2
Public Const PIECE_KNIGHT = 3
Public Const PIECE_ROOK = 4
Public Const PIECE_CANNON = 5
Public Const PIECE_PAWN = 6

'��������
Public Const MAX_GEN_MOVES As Integer = 128   '���������߷���
Public Const MAX_MOVES As Integer = 256       '������ʷ�߷���
Public Const MATE_VALUE As Integer = 10000    '��߷�ֵ���������ķ�ֵ
Public Const ADVANCED_VALUE  As Byte = 3      '����Ȩ��ֵ
Public Const LIMIT_DEPTH  As Byte = 32        '�����������
Public Const WIN_VALUE  As Integer = MATE_VALUE - 100 '������ʤ���ķ�ֵ���ޣ�������ֵ��˵���Ѿ�������ɱ����
Public Const NULL_MARGIN As Integer = 400             '�ղ��ü��������߽�
Public Const NULL_DEPTH  As Byte = 2                  '�ղ��ü��Ĳü����
Public Const DRAW_VALUE As Byte = 20     '����ʱ���صķ���(ȡ��ֵ)

Public Const HASH_SIZE As Long = 2 ^ 20  ' �û����С
Public Const HASH_ALPHA As Long = 1      ' ALPHA�ڵ���û�����
Public Const HASH_BETA As Long = 2       ' BETA�ڵ���û�����
Public Const HASH_PV As Long = 3         ' PV�ڵ���û�����
Public Const BOOK_SIZE As Long = 16384   ' ���ֿ��С
Public Const c_SizeofBookItem As Byte = 8 '���ֿ���һ����ֽ���

'��������
Public Const SND_ASYNC = &H1
Public Const SND_NOSTOP = &H10
Public Const SND_NODEFAULT = &H2
Public Const SND_NOWAIT = &H2000
Public Const SND_SYNC = &H0
Public Const SND_RESOURCE = &H40004
Public Const SND_MEMORY = &H4     'ָ��һ���ڴ��ļ�
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Declare Function GetTickCount Lib "KERNEL32" () As Long

'��(ʿ)�Ĳ���
Public ccAdvisorDelta As Variant

'��Ĳ�������˧(��)�Ĳ�����Ϊ����
Public ccKnightDelta As Variant

'�������Ĳ���������(ʿ)�Ĳ�����Ϊ����
Public ccKnightCheckDelta As Variant

'˧(��)�Ĳ���
Public ccKingDelta As Variant

'�ж������Ƿ��������е�����
Public ccInBoard As Variant

'�ж������Ƿ��ھŹ�������
Public ccInFort As Variant
 
'�жϲ����Ƿ�����ض��߷������飬1=˧(��)��2=��(ʿ)��3=��(��)
Public ccLegalSpan As Variant

'���ݲ����ж����Ƿ����ȵ�����
Public ccKnightPin As Variant

'����λ�ü�ֵ��
Public cucvlPiecePos As Variant

'MVV/LVAÿ�������ļ�ֵ
Public cucMvvLva As Variant

Public Search As New clsSearch

Public Type udtZobrist
  Player As New ZobristStruct
  Table(0 To 13, 0 To 255) As New ZobristStruct
End Type

Public Zobrist As udtZobrist

'�û�����ṹ
Public Type HashItem
  ucDepth   As Byte
  ucFlag    As Byte
  svl       As Integer
  wmv       As Integer
  wReserved As Long
  dwLock0   As Long
  dwLock1   As Long
End Type

'���ֿ���ṹ
Public Type BookItem
  dwLock As Long
  wmv    As Integer
  wvl    As Integer
End Type

Public Search_HashTable(0 To HASH_SIZE - 1) As HashItem '�û���
Public Search_BookTable(0 To BOOK_SIZE - 1) As BookItem '���ֿ�

'�߷�����׶�
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

'����߷������
Public Function SRC(ByVal mv As Long) As Long
  SRC = mv And &HFF '��8λ
End Function

'����߷����յ�
Public Function DST(ByVal mv As Integer) As Byte
 ' CopyMemory DST, ByVal (VarPtr(mv) + 1), 1 '��8λ
  DST = (mv And &HFF00&) \ &H100
End Function

'��ú�ڱ��(������8��������16)
Public Function SIDE_TAG(sd As Byte) As Byte
  SIDE_TAG = IIf(sd = 0, 8, 16)
End Function

'��öԷ���ڱ��
Public Function OPP_SIDE_TAG(sd As Byte) As Integer
  OPP_SIDE_TAG = IIf(sd = 0, 16, 8)
End Function

'����ˮƽ����
Public Function SQUARE_FORWARD(ByVal sq As Integer, sd As Byte)
  SQUARE_FORWARD = sq - 16 + (sd * 2 ^ 5)
End Function

'�ж������Ƿ���������
Public Function IN_BOARD(ByVal sq As Integer) As Boolean
  IN_BOARD = CBool(ccInBoard(sq) <> 0)
End Function

'�ж������Ƿ��ھŹ���
Public Function IN_FORT(ByVal sq As Long) As Boolean
  IN_FORT = CBool(ccInFort(sq) <> 0)
End Function

'��ø��ӵ�������
Public Function RANK_Y(sq As Integer) As Integer
  RANK_Y = sq \ 2 ^ 4
End Function

'��ø��ӵĺ�����
Public Function FILE_X(sq As Integer) As Integer
  FILE_X = sq And &HF
End Function

'����������ͺ������ø���
Public Function COORD_XY(ByVal x As Integer, ByVal y As Integer) As Integer
  COORD_XY = x + (y * 2 ^ 4)
End Function

'��ת����
Public Function SQUARE_FLIP(sq As Byte) As Integer
  SQUARE_FLIP = 254 - sq
End Function

'������ˮƽ����
Public Function FILE_FLIP(x As Integer) As Integer
  FILE_FLIP = 14 - x
End Function

'�����괹ֱ����
Public Function RANK_FLIP(y As Integer) As Integer
  RANK_FLIP = 15 - y
End Function

'����ˮƽ����
Public Function MIRROR_SQUARE(sq As Integer)
  MIRROR_SQUARE = COORD_XY(FILE_FLIP(FILE_X(sq)), RANK_Y(sq))
End Function


'�߷��Ƿ����˧(��)�Ĳ���
Public Function KING_SPAN(sqSrc As Integer, sqDst As Integer) As Boolean
  KING_SPAN = ccLegalSpan(sqDst - sqSrc + 256) = 1
End Function

'�߷��Ƿ������(ʿ)�Ĳ���
Public Function ADVISOR_SPAN(sqSrc As Integer, sqDst As Integer) As Boolean
  ADVISOR_SPAN = (ccLegalSpan(sqDst - sqSrc + 256) = 2)
End Function

'�߷��Ƿ������(��)�Ĳ���
Public Function BISHOP_SPAN(sqSrc As Integer, sqDst As Integer) As Boolean
  BISHOP_SPAN = (ccLegalSpan(sqDst - sqSrc + 256) = 3)
End Function

'��(��)�۵�λ��
Public Function BISHOP_PIN(sqSrc As Integer, sqDst As Integer) As Integer
  BISHOP_PIN = (sqSrc + sqDst) \ 2
End Function

'���ȵ�λ��
Public Function KNIGHT_PIN(sqSrc As Integer, sqDst As Integer) As Integer
  KNIGHT_PIN = sqSrc + ccKnightPin(sqDst - sqSrc + 256)
End Function

'�Ƿ�δ����
Public Function HOME_HALF(ByVal sq As Integer, ByVal sd As Integer) As Boolean
  HOME_HALF = (sq And &H80) <> (sd * 2 ^ 7)
End Function

'�Ƿ��ѹ���
Public Function AWAY_HALF(ByVal sq As Integer, ByVal sd As Integer) As Boolean
  AWAY_HALF = (sq And &H80) = (sd * 2 ^ 7)
End Function

'�Ƿ��ںӵ�ͬһ��
Public Function SAME_HALF(sqSrc As Integer, sqDst As Integer) As Boolean
  SAME_HALF = ((sqSrc Xor sqDst) And &H80) = 0
End Function

'�Ƿ���ͬһ��
Public Function SAME_RANK(sqSrc As Integer, sqDst As Integer) As Boolean
  SAME_RANK = ((sqSrc Xor sqDst) And &HF0) = 0
End Function

'�Ƿ���ͬһ��
Public Function SAME_FILE(sqSrc As Integer, sqDst As Integer) As Boolean
  SAME_FILE = ((sqSrc Xor sqDst) And &HF) = 0
End Function

'���������յ����߷�
Public Function MOVE_(ByVal sqSrc As Byte, ByVal sqDst As Byte) As Integer
  CopyMemory MOVE_, sqSrc, 1
  CopyMemory ByVal (VarPtr(MOVE_) + 1), sqDst, 1
End Function

'�߷�ˮƽ����
Public Function MIRROR_MOVE(mv As Integer) As Integer
  MIRROR_MOVE = MOVE_(MIRROR_SQUARE(SRC(mv)), MIRROR_SQUARE(DST(mv)))
End Function
 
'������Դ����
Public Sub PlayResWav(nResId As Integer)
  Dim WavData() As Byte

  WavData = LoadResData(nResId, "WAVE")
  sndPlaySound WavData(0), SND_SYNC Or SND_NOWAIT Or SND_MEMORY
End Sub

'����ʷ������ıȽϺ���
Public Function CompareHistory(lpmv1 As Integer, lpmv2 As Integer) As Boolean
  If frmSettings.optGreateOrEqual.Value Then
    CompareHistory = Search.nHistoryTable(lpmv2) - Search.nHistoryTable(lpmv1) >= 0
  Else
    CompareHistory = Search.nHistoryTable(lpmv2) - Search.nHistoryTable(lpmv1) > 0
  End If
End Function

'��MVV/LVAֵ����ıȽϺ���
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
  
  '��Ҫ�����������A[lngLeft]����A[lngRight]����������ѡȡһ�����ݣ�ͨ��ѡ�õ�һ�����ݣ���Ϊ�ؼ����ݣ�
  'Ȼ�����б���С�������ŵ���ǰ�棬���б�����������ŵ������棬������̳�Ϊһ�˿�������
  'һ�˿���������㷨�ǣ�
  
  '1��������������i��j������ʼ��ʱ��i = lngLeft��j = lngRight��
  i = lngLeft
  j = lngRight
  
  '2���Ե�һ������Ԫ����Ϊ�ؼ����ݣ���ֵ��X���� X=A[lngLeft]��
  x = mvs(lngLeft)
  
  Do While i < j
    '3����j��ʼ��ǰ���������ɺ�ʼ��ǰ������j=j-1�����ҵ���һ��С��X��ֵ���ø�ֵ��X�������ҵ�����.�ҵ���i��С���䣩��
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
    
    '4����i��ʼ�������������ǰ��ʼ���������i=i+1�����ҵ���һ������X��ֵ���ø�ֵ��X�������ҵ�����.�ҵ���j��С���䣩��
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
    '5���ظ���3��4����ֱ�� i=j�� (3,4�����ڳ�����û�ҵ�ʱ��j=j-1��i=i+1��
    '�ҵ���������ʱ��i�� jָ��λ�ò��䡣���⵱i=j�����һ��������i+1��j-1��ɵ������ѭ������)
  Loop
  
  '����������ǵݹ���ô˹��̡����Թؼ����ݣ�ͨ��ѡ�õ�һ�����ݣ�Ϊ�е�ָ�����������У��ֱ��ǰ��һ���ֺͺ���һ���ֽ������ƵĿ�������
  '�Ӷ����ȫ���������еĿ����������Ѵ��������б��һ�����������
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
        Pc2Name = "˧"
      Case 9
        Pc2Name = "��"
      Case 10
        Pc2Name = "��"
      Case 11
        Pc2Name = "��"
      Case 12
        Pc2Name = "��"
      Case 13
        Pc2Name = "��"
      Case 14
        Pc2Name = "��"
    
    
      Case 16
        Pc2Name = "��"
      Case 17
        Pc2Name = "ʿ"
      Case 18
        Pc2Name = "��"
      Case 19
        Pc2Name = "��"
      Case 20
        Pc2Name = "��"
      Case 21
        Pc2Name = "��"
      Case 22
        Pc2Name = "��"
        
      Case Else
        Pc2Name = "Error"
    End Select
  End If
End Function

'��MVV/LVAֵ
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

'��ȡ�û�����
Public Function ProbeHash(vlAlpha As Long, vlBeta As Long, nDepth As Long, ByRef mv As Integer) As Long
  Dim bMate As Boolean 'ɱ���־�������ɱ�壬��ô����Ҫ�����������
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

'�����û�����
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

'�������ֿ�
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
  
  
  '�������ֿ�Ĺ��������¼�������

  '1. ���û�п��ֿ⣬����������
  If (Search.nBookSize = 0) Then
    SearchBook = 0
    Exit Function
  End If

  '2. ������ǰ����
  bMirror = False
  bkToSearch.dwLock = pos.zobr.dwLock1
  lpbk = bsearch(bkToSearch.dwLock)
  
  '3. ���û���ҵ�����ô������ǰ����ľ������
  If lpbk = -1 Then
    bMirror = True
    pos.Mirror posMirror
    bkToSearch.dwLock = posMirror.zobr.dwLock1
    lpbk = bsearch(bkToSearch.dwLock)
  End If
  
  '4. ����������Ҳû�ҵ�������������
  If lpbk = -1 Then
    SearchBook = 0
    Exit Function
  End If

  '5. ����ҵ�������ǰ���һ�����ֿ���
  Do While (lpbk >= 0 And Search_BookTable(lpbk).dwLock = bkToSearch.dwLock)
    lpbk = lpbk - 1
  Loop
  
  lpbk = lpbk + 1
  
  '6. ���߷��ͷ�ֵд�뵽"mvs"��"vls"������
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
        Exit Do  '��ֹ"BOOK.DAT"�к����쳣����
      End If
    End If
    lpbk = lpbk + 1
  Loop
  
  If (vl = 0) Then
    SearchBook = 0
    Exit Function '��ֹ"BOOK.DAT"�к����쳣����
  End If
  
  '7. ����Ȩ�����ѡ��һ���߷�
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
  frmSearchInfo.lsvValue.ListItems(frmSearchInfo.lsvValue.ListItems.Count).SubItems(4) = "���ֿ�"
  SearchBook = mvs(i)
End Function

Public Function GetMoveDesc(ByVal mv As Integer, Optional bSimulateMove As Boolean = True, Optional bSymbolMode As Boolean = False) As String
  Dim SrcBlk As Byte
  Dim DstBlk As Byte
  Dim SrcX As Byte  '�������꣨����ʱ��������ȣ�
  Dim SrcX2 As Byte '������꣨���������������
  Dim SrcY As Byte
  Dim DstX As Byte
  Dim DstY As Byte
  Dim pc As Byte  '(������������ɫ��Ϣ)
  Dim pc2 As Byte '(��ȥ��������ɫ��Ϣ)
  Dim sAct As String
  Dim sStepDstX As String
  Dim bIsRed As Boolean
  Dim y As Byte
  Dim ScanPc As Byte
  Dim sWhich As String
  Dim bCondition1 As Boolean
  Dim bCondition2 As Boolean
  
  DstBlk = (mv And &HFF00&) \ &H100
  'CopyMemory DstBlk, ByVal (VarPtr(mv) + 1), 1 '��8λ
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
  If pc2 <> PIECE_BISHOP And pc2 <> PIECE_ADVISOR Then '������֡�ǰ����һ���򡰺��˽��塱���������׻ᱻ��Ц������
    For y = 1 To 10
      ScanPc = pos.ucpcSquares(COORD_XY(SrcX + 2, y + 2))
      '�غ�����1���ڸ����ҵ���ͬ������
      bCondition1 = CBool(ScanPc = pc)
      
      '��������Ҳ���˼��ʱ��ģ���ƶ���ֻ�������غ�����1���ɣ����غ�����2��Զ���㣻������������������ģ���ƶ�����Ҫ�ж��ڸ����ҵ��������ǲ����Լ�
      bCondition2 = IIf(pc2 = PIECE_KNIGHT And Not bSimulateMove, True, CBool(y <> IIf(bSimulateMove, SrcY, DstY)))
      
      If bCondition1 And bCondition2 Then
        If bIsRed Then
          sWhich = IIf(y > SrcY, "ǰ", "��")
        Else
          sWhich = IIf(y > SrcY, "��", "ǰ")
        End If
        Exit For
      End If
    Next y
  End If
  
  Select Case pc2
    Case 0, 4, 5, 6 'ֱ����
      If DstY > SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "-", "��")
        Else
          sAct = IIf(bSymbolMode, "+", "��")
        End If
        sStepDstX = IIf(bSymbolMode, CStr(DstY - SrcY), CCStr(DstY - SrcY, bIsRed))
      ElseIf DstY < SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "+", "��")
        Else
          sAct = IIf(bSymbolMode, "-", "��")
        End If
        sStepDstX = IIf(bSymbolMode, CStr(SrcY - DstY), CCStr(SrcY - DstY, bIsRed))
      Else
        sAct = IIf(bSymbolMode, ".", "ƽ")
        DstX = IIf(bIsRed, 10 - DstX, DstX)
        sStepDstX = IIf(bSymbolMode, CStr(DstX), CCStr(DstX, bIsRed))
      End If
    Case 1, 2, 3 'б����
      If DstY > SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "-", "��")
        Else
          sAct = IIf(bSymbolMode, "+", "��")
        End If
      ElseIf DstY < SrcY Then
        If bIsRed Then
          sAct = IIf(bSymbolMode, "+", "��")
        Else
          sAct = IIf(bSymbolMode, "-", "��")
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
  Const sChiNum As String = "һ�����������߰˾�"
  Const sAlbNum As String = "������������������"
  
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

'װ�뿪�ֿ�
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
