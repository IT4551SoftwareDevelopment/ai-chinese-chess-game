VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "搜索信息"
   ClientHeight    =   8985
   ClientLeft      =   11925
   ClientTop       =   285
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lsvDepthTimeCost 
      Height          =   1680
      Left            =   30
      TabIndex        =   1
      Top             =   7260
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   2963
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lsvMoveList 
      Height          =   8925
      Left            =   4140
      TabIndex        =   0
      Top             =   15
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   15743
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lsvValue 
      Height          =   6915
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   12197
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblSearchSeconds 
      AutoSize        =   -1  'True
      Caption         =   "已搜索时间：0 秒"
      Height          =   195
      Left            =   1395
      TabIndex        =   3
      Top             =   6990
      Width           =   1395
   End
End
Attribute VB_Name = "frmSearchInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

Private Sub Form_Load()
  With frmSearchInfo.lsvValue
    .ColumnHeaders.Add , , "No", 356
    .ColumnHeaders.Add , , "分数", 650
    .ColumnHeaders.Add , , "深度", 540
    .ColumnHeaders.Add , , "耗时", 650
    .ColumnHeaders.Add , , "访问节点数", 1080
  End With
  
  With frmSearchInfo.lsvDepthTimeCost
    .ColumnHeaders.Add , , "深度", 540
    .ColumnHeaders.Add , , "进度", 650
    .ColumnHeaders.Add , , "耗时", 650
    .ColumnHeaders.Add , , "走法", 900
    .ColumnHeaders.Add , , "分数", 650
    .ColumnHeaders.Add , , "动态搜索节点数", 800
    .ColumnHeaders.Add , , "静态搜索节点数", 1450
  End With
  
  With frmSearchInfo.lsvMoveList
    .ColumnHeaders.Add , , "No", 356
    .ColumnHeaders.Add , , "走法", 900
    .ColumnHeaders.Add , , "分数", 713
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuSearchInfo_Click
  Cancel = 1
End Sub

