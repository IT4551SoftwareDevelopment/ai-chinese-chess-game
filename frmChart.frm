VERSION 5.00
Begin VB.Form frmChart 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "局势变化图"
   ClientHeight    =   2235
   ClientLeft      =   2460
   ClientTop       =   9675
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picChart 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   2265
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   12705
      TabIndex        =   0
      Top             =   0
      Width           =   12765
      Begin VB.Label lblOpening 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开局描述"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

Public mHGrids As Integer
Public mStepWidth As Integer

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuChart_Click
  Cancel = 1
End Sub

Private Sub picChart_Paint()
  DrawGrid
  DrawValueChart
End Sub

Public Sub DrawGrid()
  Dim i As Integer
  
  picChart.ForeColor = RGB(0, 255, 0)
  picChart.Line (0, picChart.Height / 2)-(picChart.Width, picChart.Height / 2)
  picChart.ForeColor = RGB(0, 64, 0)
  
  If frmSearchInfo.lsvValue.ListItems.Count > mHGrids Then
    mHGrids = frmSearchInfo.lsvValue.ListItems.Count
    mStepWidth = picChart.Width / mHGrids
  End If
  
  For i = 1 To mHGrids - 1
    picChart.Line (i * mStepWidth, 0)-(i * mStepWidth, picChart.Height)
  Next i
End Sub

Public Sub DrawValueChart()
  Dim i As Byte
  Dim y As Integer
  Dim max As Integer
  
  '找最大值
  max = 0
  For i = 1 To frmSearchInfo.lsvValue.ListItems.Count
    If Abs(Val(frmSearchInfo.lsvValue.ListItems(i).SubItems(1))) > max Then
      max = Abs(Val(frmSearchInfo.lsvValue.ListItems(i).SubItems(1)))
    End If
  Next i
  
  '画圆
  picChart.ForeColor = vbYellow
  
  picChart.Circle (0, picChart.Height / 2), 30, vbYellow
  For i = 1 To frmSearchInfo.lsvValue.ListItems.Count
    If max = 0 Then
      y = picChart.ScaleHeight / 2
    Else
      y = picChart.ScaleHeight / 2 - Val(frmSearchInfo.lsvValue.ListItems(i).SubItems(1)) / max * (picChart.ScaleHeight / 2)
    End If
    picChart.Circle (i * mStepWidth, y), 30, vbYellow
  Next i
  
  
  '画折线
  picChart.PSet (0, picChart.Height / 2)
  For i = 1 To frmSearchInfo.lsvValue.ListItems.Count
    If max = 0 Then
      y = picChart.ScaleHeight / 2
    Else
      y = picChart.ScaleHeight / 2 - Val(frmSearchInfo.lsvValue.ListItems(i).SubItems(1)) / max * (picChart.ScaleHeight / 2)
    End If
    picChart.Line -(i * mStepWidth, y)
  Next i
End Sub

