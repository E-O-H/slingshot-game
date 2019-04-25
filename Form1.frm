VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SlingShot config"
   ClientHeight    =   5145
   ClientLeft      =   6450
   ClientTop       =   4680
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6735
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   10
      Text            =   "-2"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Text            =   "1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Text            =   "5"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   4
      Text            =   "300"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始游戏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1920
      TabIndex        =   2
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   1
      Text            =   "4"
      Top             =   0
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6720
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label6 
      Caption         =   $"Form1.frx":0000
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label Label5 
      Caption         =   "引力公式R指数"
      BeginProperty Font 
         Name            =   "经典繁颜体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "万有引力强度"
      BeginProperty Font 
         Name            =   "经典繁颜体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "大小地图缩放比例"
      BeginProperty Font 
         Name            =   "经典繁颜体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "MAXTIME"
      BeginProperty Font 
         Name            =   "经典繁颜体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "最大星球数量"
      BeginProperty Font 
         Name            =   "经典繁颜体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PNM = Text1.Text
MT = Text2.Text
n = Text3.Text
G = Text4.Text
index = Text5.Text
Form2.Show
Form1.Hide
Form3.Line1.X1 = 9600 - 9600 / n
Form3.Line1.X2 = 9600 - 9600 / n
Form3.Line1.Y1 = 6000 - 6000 / n
Form3.Line1.Y2 = 12000 / n + 6000 - 6000 / n

Form3.Line2.X1 = 9600 - 9600 / n
Form3.Line2.X2 = 19200 / n + 9600 - 9600 / n
Form3.Line2.Y1 = 6000 - 6000 / n
Form3.Line2.Y2 = 6000 - 6000 / n

Form3.Line3.X1 = 19200 / n + 9600 - 9600 / n
Form3.Line3.X2 = 19200 / n + 9600 - 9600 / n
Form3.Line3.Y1 = 6000 - 6000 / n
Form3.Line3.Y2 = 12000 / n + 6000 - 6000 / n

Form3.Line4.X1 = 9600 - 9600 / n
Form3.Line4.X2 = 19200 / n + 9600 - 9600 / n
Form3.Line4.Y1 = 12000 / n + 6000 - 6000 / n
Form3.Line4.Y2 = 12000 / n + 6000 - 6000 / n

Unload Form1
End Sub

