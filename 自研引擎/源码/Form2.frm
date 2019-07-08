VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "设置"
   ClientHeight    =   3240
   ClientLeft      =   6705
   ClientTop       =   2115
   ClientWidth     =   7545
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   7545
   Begin VB.CheckBox Check6 
      Caption         =   "能否管理地形(选为假地形将不会被鼠标管理)"
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CheckBox Check5 
      Caption         =   "粒子标记"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "灯光标记"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   4320
      TabIndex        =   17
      Text            =   "100"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "校准"
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Z"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Y"
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "X"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   2400
      Width           =   6855
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4320
      TabIndex        =   10
      Text            =   "100"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   4320
      TabIndex        =   8
      Text            =   "100"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1320
      TabIndex        =   5
      Text            =   "1048"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1320
      TabIndex        =   4
      Text            =   "1048"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Text            =   "1048"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "总宽长高宽度是米*2为单位"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "格子高度"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "格子长度"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "格子宽度"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "网格高度"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "网格长度"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "网格宽度"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = CSng(Text1.Text)
z = CSng(Text2.Text)
y = CSng(Text3.Text)
xx = CSng(Text4.Text)
zz = CSng(Text5.Text)
yy = CSng(Text6.Text)
If Check2.Value = 1 Then xwg = True Else xwg = False
If Check3.Value = 1 Then ywg = True Else ywg = False
If Check4.Value = 1 Then zwg = True Else zwg = False
If Check1.Value = 1 Then kg2 = True Else kg2 = False
If Check5.Value = 1 Then kg3 = True Else kg3 = False
If Check6.Value = 1 Then kg4 = True Else kg4 = False
'debug.Print tkwg
Unload Form2
End Sub
Private Sub Command2_Click()
Text1.Text = CStr(map.getter.GetLandRealWidth / 2)
Text2.Text = CStr(map.getter.GetLandRealHeight / 2)
Text3.Text = "1000"
End Sub
Private Sub Form_Load()
Text1.Text = CStr(x)
Text2.Text = CStr(z)
Text3.Text = CStr(y)
Text4.Text = CStr(xx)
Text5.Text = CStr(zz)
Text6.Text = CStr(yy)
Check2.Value = IIf(xwg = True, 1, 0)
Check3.Value = IIf(ywg = True, 1, 0)
Check4.Value = IIf(zwg = True, 1, 0)
Check1.Value = IIf(kg2 = True, 1, 0)
Check5.Value = IIf(kg3 = True, 1, 0)
Check6.Value = IIf(kg4 = True, 1, 0)
End Sub
