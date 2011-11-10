VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置"
   ClientHeight    =   2685
   ClientLeft      =   13710
   ClientTop       =   9645
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRANDOM_MASK 
      Height          =   285
      Left            =   4935
      MaxLength       =   3
      TabIndex        =   23
      Text            =   "7"
      Top             =   2325
      Width           =   375
   End
   Begin VB.CheckBox chkOpeningBook 
      Caption         =   "开局库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2685
      TabIndex        =   21
      Top             =   2355
      Value           =   1  'Checked
      Width           =   840
   End
   Begin VB.CheckBox chkKiller 
      Caption         =   "杀手走法"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1515
      TabIndex        =   20
      Top             =   2355
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox chkTranTable 
      Caption         =   "置换表裁剪"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   19
      Top             =   2355
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.CheckBox chkSortByMvvLva 
      Caption         =   "在静态搜索中将生成的吃子走法按 Mvv/Lva 进行排序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   150
      TabIndex        =   15
      Top             =   570
      Value           =   1  'Checked
      Width           =   2850
   End
   Begin VB.Frame Frame4 
      Caption         =   "克服水平线效应"
      Height          =   1140
      Left            =   3285
      TabIndex        =   14
      Top             =   1110
      Width           =   2040
      Begin VB.CheckBox chkNullMove 
         Caption         =   "空步裁剪"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   18
         Top             =   810
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkInCheckExt 
         Caption         =   "将军延伸"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   17
         Top             =   510
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkSearchQuiesc 
         Caption         =   "静态搜索"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   255
         Value           =   1  'Checked
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "搜索控制"
      Height          =   960
      Left            =   3285
      TabIndex        =   8
      Top             =   0
      Width           =   2100
      Begin VB.TextBox txtDepthLimit 
         Height          =   285
         Left            =   1110
         TabIndex        =   12
         Text            =   "4"
         Top             =   585
         Width           =   615
      End
      Begin VB.TextBox txtTimeLimit 
         Height          =   285
         Left            =   1110
         TabIndex        =   11
         Text            =   "3000"
         Top             =   225
         Width           =   600
      End
      Begin VB.OptionButton optLimitDepth 
         Caption         =   "限定深度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   585
         Width           =   1035
      End
      Begin VB.OptionButton optLimitTime 
         Caption         =   "限定时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   255
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ms"
         Height          =   195
         Left            =   1800
         TabIndex        =   13
         Top             =   285
         Width           =   195
      End
   End
   Begin VB.Frame fraSort 
      Caption         =   "排序"
      Height          =   2250
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3180
      Begin VB.CheckBox chkSortMoves 
         Caption         =   "根据历史表对走法进行排序"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   255
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.Frame Frame1 
         Caption         =   "比较符号"
         Height          =   1035
         Left            =   1635
         TabIndex        =   4
         Top             =   1140
         Width           =   1440
         Begin VB.OptionButton optGreater 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   6
            Top             =   345
            Width           =   525
         End
         Begin VB.OptionButton optGreateOrEqual 
            Caption         =   ">="
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   165
            TabIndex        =   5
            Top             =   615
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "排序算法"
         Height          =   1035
         Left            =   60
         TabIndex        =   1
         Top             =   1140
         Width           =   1440
         Begin VB.OptionButton optSelectSort 
            Caption         =   "选择排序法"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   75
            TabIndex        =   3
            Top             =   585
            Width           =   1260
         End
         Begin VB.OptionButton optQuickSort 
            Caption         =   "快速排序法"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   75
            TabIndex        =   2
            Top             =   330
            Value           =   -1  'True
            Width           =   1305
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "随机度："
      Height          =   195
      Left            =   4140
      TabIndex        =   22
      Top             =   2355
      Width           =   720
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'http://www.vbgood.com/viewthread.php?tid=89344&extra=page%3D1&page=10
'VBProFan

Private Sub chkSearchQuiesc_Click()
  chkSortByMvvLva.Enabled = chkSearchQuiesc.Value = vbChecked
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuSettings_Click
  Cancel = 1
End Sub

Private Sub optLimitDepth_Click()
  With txtDepthLimit
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
  End With
End Sub

Private Sub txtDepthLimit_GotFocus()
  optLimitDepth.Value = True
End Sub

Private Sub txtTimeLimit_GotFocus()
  optLimitTime.Value = True
End Sub


