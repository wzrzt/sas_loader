VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "启动 SAS9.4"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6060
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6060
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Left            =   2400
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   2880
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "秒内关闭"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "此窗口将在"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d, t, a
Dim e As Integer


Private Sub Form_Load()
Me.Top = 1200
Me.Left = 1500
a = 10
Label2.Caption = a
d = Date
t = Format(Now, "HHmmss")
Label4.Caption = "今天是": Label5.Caption = d
If Dir(App.Path & "\sas.exe") <> "" Then '存在
Date = "2015-05-01"
Shell Dir(App.Path & "\sas.exe"), vbNormalFocus  '打开某exe
Else
Me.Visible = True
Label6.Caption = "sas.exe找不到，请检查路径"
e = MsgBox("请将此文件放于sas安装目录中的sas.exe旁边!", vbOKOnly, "ERROR")

End If



'要执行的代码

Timer1.Interval = 10000
Timer2.Interval = 1000

   End Sub
 
Private Sub Timer1_Timer()
If t >= 235950 Then d = d + 1
If Date <> d Then Date = d

Timer2.Interval = 0
End
End Sub

Private Sub Timer2_Timer()
a = a - 1
Label2.Caption = a
End Sub

