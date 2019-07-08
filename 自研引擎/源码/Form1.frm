VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TV3D可视化编辑器--小熊制作，版权所有，翻版必究,,小熊QQ:1066562980"
   ClientHeight    =   13035
   ClientLeft      =   3375
   ClientTop       =   390
   ClientWidth     =   22995
   LinkTopic       =   "Form1"
   ScaleHeight     =   13035
   ScaleWidth      =   22995
   Begin VB.CommandButton Command34 
      Caption         =   "浏览"
      Height          =   225
      Left            =   20760
      TabIndex        =   163
      Top             =   9360
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Caption         =   "浏览"
      Height          =   225
      Left            =   20760
      TabIndex        =   162
      Top             =   9840
      Width           =   735
   End
   Begin VB.CommandButton Command32 
      Caption         =   "浏览"
      Height          =   225
      Left            =   20760
      TabIndex        =   161
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Caption         =   "关于"
      Height          =   255
      Left            =   19560
      TabIndex        =   157
      Top             =   12480
      Width           =   495
   End
   Begin VB.CommandButton Command31 
      Caption         =   "浏览"
      Height          =   225
      Left            =   20760
      TabIndex        =   155
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "新工程"
      Height          =   255
      Left            =   18360
      TabIndex        =   149
      Top             =   12480
      Width           =   1095
   End
   Begin VB.CommandButton Command30 
      Caption         =   "删除"
      Height          =   255
      Left            =   22200
      TabIndex        =   148
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Caption         =   "添加"
      Height          =   255
      Left            =   21240
      TabIndex        =   147
      Top             =   7080
      Width           =   855
   End
   Begin VB.ListBox List7 
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   20040
      List            =   "Form1.frx":0002
      TabIndex        =   145
      Top             =   7320
      Width           =   2895
   End
   Begin VB.ListBox List6 
      Height          =   2220
      ItemData        =   "Form1.frx":0004
      Left            =   20040
      List            =   "Form1.frx":0006
      TabIndex        =   142
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   20880
      Top             =   7920
   End
   Begin VB.CommandButton Command27 
      Caption         =   "停止"
      Height          =   255
      Left            =   22320
      TabIndex        =   140
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      Caption         =   "暂停"
      Height          =   255
      Left            =   21480
      TabIndex        =   139
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command25 
      Caption         =   "创建工程目录"
      Height          =   375
      Left            =   14760
      TabIndex        =   138
      Top             =   12000
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "播放"
      Height          =   255
      Left            =   20760
      TabIndex        =   137
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "删除"
      Height          =   255
      Left            =   22080
      TabIndex        =   136
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "添加"
      Height          =   255
      Left            =   20040
      TabIndex        =   135
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command22 
      Caption         =   "校准"
      Height          =   255
      Left            =   22080
      TabIndex        =   133
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   132
      Top             =   9960
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   131
      Top             =   7800
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Gohere"
      Enabled         =   0   'False
      Height          =   255
      Left            =   22200
      TabIndex        =   130
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox pxx 
      Height          =   270
      Left            =   21960
      TabIndex        =   126
      Text            =   "1"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox pyy 
      Height          =   270
      Left            =   21960
      TabIndex        =   125
      Text            =   "1"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox pzzz 
      Height          =   270
      Left            =   21960
      TabIndex        =   124
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox pxxx 
      Height          =   270
      Left            =   21000
      TabIndex        =   123
      Text            =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox pzz 
      Height          =   270
      Left            =   21960
      TabIndex        =   122
      Text            =   "1"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox pyyy 
      Height          =   270
      Left            =   20520
      TabIndex        =   121
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox pp 
      Height          =   270
      Left            =   21960
      TabIndex        =   120
      Text            =   "1"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "设置"
      Enabled         =   0   'False
      Height          =   255
      Left            =   21120
      TabIndex        =   115
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "添加"
      Enabled         =   0   'False
      Height          =   255
      Left            =   20040
      TabIndex        =   114
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox pz 
      Height          =   270
      Left            =   20520
      TabIndex        =   110
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox py 
      Height          =   270
      Left            =   20520
      TabIndex        =   109
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox px 
      Height          =   270
      Left            =   20520
      TabIndex        =   108
      Text            =   "0"
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox List5 
      Height          =   1680
      ItemData        =   "Form1.frx":0008
      Left            =   20040
      List            =   "Form1.frx":000A
      TabIndex        =   107
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command15 
      Caption         =   "添加"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   105
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "添加"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   104
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "设置材质"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   103
      Top             =   8520
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Gohere"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   102
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Gohere"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   101
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "设置材质"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   100
      Top             =   6360
      Width           =   615
   End
   Begin VB.ListBox List4 
      Height          =   1860
      ItemData        =   "Form1.frx":000C
      Left            =   0
      List            =   "Form1.frx":000E
      TabIndex        =   98
      Top             =   8520
      Width           =   2295
   End
   Begin VB.ListBox List3 
      Height          =   1860
      ItemData        =   "Form1.frx":0010
      Left            =   0
      List            =   "Form1.frx":0012
      TabIndex        =   95
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   1320
      Top             =   6600
   End
   Begin VB.CommandButton Command9 
      Caption         =   "设置材质"
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      TabIndex        =   94
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox cz 
      Height          =   270
      Left            =   720
      TabIndex        =   91
      Text            =   "材质1"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "删除材质"
      Enabled         =   0   'False
      Height          =   300
      Left            =   1920
      TabIndex        =   90
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "创建材质"
      Enabled         =   0   'False
      Height          =   300
      Left            =   0
      TabIndex        =   89
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox zs 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2160
      TabIndex        =   88
      Text            =   "0"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox xsb 
      Enabled         =   0   'False
      Height          =   270
      Left            =   480
      TabIndex        =   86
      Text            =   "0"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox ssr 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2160
      TabIndex        =   79
      Text            =   "0"
      Top             =   2880
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   1140
      ItemData        =   "Form1.frx":0014
      Left            =   0
      List            =   "Form1.frx":0016
      TabIndex        =   78
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CheckBox prt 
      Caption         =   "PRT表面   减少散射R"
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox btm 
      Height          =   270
      Left            =   2280
      TabIndex        =   76
      Text            =   "255"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox r3 
      Height          =   270
      Left            =   840
      TabIndex        =   61
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox g3 
      Height          =   270
      Left            =   840
      TabIndex        =   60
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox b3 
      Height          =   270
      Left            =   840
      TabIndex        =   59
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox tm3 
      Height          =   270
      Left            =   840
      TabIndex        =   58
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox r4 
      Height          =   270
      Left            =   2160
      TabIndex        =   57
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox g4 
      Height          =   270
      Left            =   2160
      TabIndex        =   56
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox b4 
      Height          =   270
      Left            =   2160
      TabIndex        =   55
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox tm4 
      Height          =   270
      Left            =   2160
      TabIndex        =   54
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox gg 
      Height          =   270
      Left            =   840
      TabIndex        =   53
      Text            =   "81"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox ssg 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2160
      TabIndex        =   52
      Text            =   "0"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox ssb 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2160
      TabIndex        =   51
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox xsr 
      Enabled         =   0   'False
      Height          =   270
      Left            =   480
      TabIndex        =   50
      Text            =   "0"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox xsg 
      Enabled         =   0   'False
      Height          =   270
      Left            =   480
      TabIndex        =   49
      Text            =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox tm2 
      Height          =   270
      Left            =   2160
      TabIndex        =   48
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox G 
      Height          =   270
      Left            =   600
      TabIndex        =   46
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Bbjk 
      Height          =   270
      Left            =   600
      TabIndex        =   45
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox R2 
      Height          =   270
      Left            =   2160
      TabIndex        =   44
      Text            =   "0"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox G2 
      Height          =   270
      Left            =   2160
      TabIndex        =   43
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox B2 
      Height          =   270
      Left            =   2160
      TabIndex        =   42
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox TM 
      Height          =   270
      Left            =   600
      TabIndex        =   41
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox R 
      Height          =   270
      Left            =   600
      TabIndex        =   38
      Text            =   "0"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "纹理"
      Height          =   375
      Left            =   13920
      TabIndex        =   34
      Top             =   12000
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   22200
      Top             =   11040
   End
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   375
      Left            =   19080
      TabIndex        =   33
      Top             =   12120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "设置"
      Height          =   375
      Left            =   13080
      TabIndex        =   31
      Top             =   12000
      Width           =   735
   End
   Begin VB.CheckBox Check6 
      Caption         =   "灰度图"
      Height          =   255
      Left            =   14400
      TabIndex        =   30
      Top             =   12480
      Width           =   855
   End
   Begin VB.CheckBox Check5 
      Caption         =   "平坦"
      Enabled         =   0   'False
      Height          =   375
      Left            =   17640
      TabIndex        =   29
      Top             =   12480
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   12360
      TabIndex        =   28
      Text            =   "8"
      Top             =   12480
      Width           =   420
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   270
      Left            =   15360
      TabIndex        =   26
      Text            =   "C:\Documents and Settings\琉璃\桌面\heightmap.JPG"
      Top             =   12480
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   13560
      TabIndex        =   25
      Text            =   "32"
      Top             =   12480
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   10920
      TabIndex        =   22
      Text            =   "8"
      Top             =   12480
      Width           =   375
   End
   Begin VB.CheckBox Check4 
      Caption         =   "俯视工具F1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   12000
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "相对"
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   12480
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   9480
      TabIndex        =   18
      Text            =   "1"
      Top             =   12480
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "递增"
      Height          =   255
      Left            =   7920
      TabIndex        =   17
      Top             =   12480
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "固定高度"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   12480
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   8880
      TabIndex        =   15
      Text            =   "1"
      Top             =   12000
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Left            =   6960
      TabIndex        =   13
      Text            =   "5"
      Top             =   12480
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   6960
      TabIndex        =   8
      Text            =   "20"
      Top             =   12000
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   10800
      TabIndex        =   5
      Text            =   "2"
      Top             =   12000
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   14160
      TabIndex        =   4
      Text            =   "fsbzwy"
      Top             =   11520
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "载入"
      Height          =   375
      Left            =   16560
      TabIndex        =   3
      Top             =   11520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3960
      TabIndex        =   1
      Text            =   "C:\Documents and Settings\琉璃\桌面\E3DPackage.pak"
      Top             =   11520
      Width           =   9255
   End
   Begin VB.PictureBox Picture1 
      Height          =   11415
      Left            =   2880
      ScaleHeight     =   11355
      ScaleWidth      =   17115
      TabIndex        =   0
      Top             =   0
      Width           =   17175
      Begin VB.Label Label54 
         Caption         =   "Label54"
         Height          =   375
         Left            =   17160
         TabIndex        =   143
         Top             =   7080
         Width           =   15
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存"
      Enabled         =   0   'False
      Height          =   375
      Left            =   18240
      TabIndex        =   20
      Top             =   11520
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2040
      ItemData        =   "Form1.frx":0018
      Left            =   0
      List            =   "Form1.frx":001A
      TabIndex        =   35
      Top             =   10680
      Width           =   2895
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   16320
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   164
      Top             =   8280
      Width           =   1200
   End
   Begin VB.Label Label64 
      Caption         =   "结束脚本"
      Height          =   255
      Left            =   20040
      TabIndex        =   160
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label Label63 
      Caption         =   "渲染脚本"
      Height          =   255
      Left            =   20040
      TabIndex        =   159
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label casdsd 
      Caption         =   "逻辑脚本"
      Height          =   255
      Left            =   20040
      TabIndex        =   158
      Top             =   9360
      Width           =   735
   End
   Begin VB.Label Label62 
      Caption         =   "开始脚本"
      Height          =   255
      Left            =   20040
      TabIndex        =   156
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label Label61 
      Caption         =   "销毁地图后：Destroy"
      Height          =   255
      Left            =   21480
      TabIndex        =   154
      Top             =   9840
      Width           =   2295
   End
   Begin VB.Label Label60 
      Caption         =   "渲染：Render.lua"
      Height          =   255
      Left            =   21480
      TabIndex        =   153
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label Label59 
      Caption         =   "逻辑：Logic.lua"
      Height          =   255
      Left            =   21480
      TabIndex        =   152
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label Label58 
      Caption         =   "载入地图后：Main"
      Height          =   255
      Left            =   21480
      TabIndex        =   151
      Top             =   9120
      Width           =   2415
   End
   Begin VB.Label Label57 
      Caption         =   "脚本管理器："
      Height          =   255
      Left            =   20040
      TabIndex        =   150
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label Label56 
      Caption         =   "文件管理器："
      Height          =   255
      Left            =   20040
      TabIndex        =   146
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label55 
      Caption         =   "音频采样精度一定要低于或等于8位！(采样大小),MP3是无法播放的"
      Height          =   375
      Left            =   20040
      TabIndex        =   144
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label yyxx 
      Caption         =   "停止"
      Height          =   255
      Left            =   21240
      TabIndex        =   141
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label53 
      Caption         =   "音效管理器："
      Height          =   255
      Left            =   20040
      TabIndex        =   134
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label52 
      Caption         =   "旋转Z"
      Height          =   255
      Left            =   21360
      TabIndex        =   129
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label51 
      Caption         =   "旋转Y"
      Height          =   255
      Left            =   20040
      TabIndex        =   128
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label50 
      Caption         =   "旋转X"
      Height          =   255
      Left            =   20280
      TabIndex        =   127
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label49 
      Caption         =   "缩放平面"
      Height          =   255
      Left            =   21240
      TabIndex        =   119
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label48 
      Caption         =   "缩放Z"
      Height          =   255
      Left            =   21480
      TabIndex        =   118
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label47 
      Caption         =   "缩放Y"
      Height          =   255
      Left            =   21480
      TabIndex        =   117
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label46 
      Caption         =   "缩放X"
      Height          =   255
      Left            =   21480
      TabIndex        =   116
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label45 
      Caption         =   "Z:"
      Height          =   255
      Left            =   20280
      TabIndex        =   113
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label44 
      Caption         =   "Y:"
      Height          =   255
      Left            =   20280
      TabIndex        =   112
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label43 
      Caption         =   "X:"
      Height          =   255
      Left            =   20280
      TabIndex        =   111
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label42 
      Caption         =   "粒子管理器:"
      Height          =   255
      Left            =   20040
      TabIndex        =   106
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label41 
      Caption         =   "角色"
      Height          =   495
      Left            =   0
      TabIndex        =   99
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label Label40 
      Caption         =   "模型"
      Height          =   255
      Left            =   240
      TabIndex        =   97
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label39 
      Caption         =   "材质管理器："
      Height          =   255
      Left            =   0
      TabIndex        =   96
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label38 
      Caption         =   "（双击列表选择材质）"
      Height          =   255
      Left            =   1080
      TabIndex        =   93
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label33 
      Caption         =   "材质名"
      Height          =   255
      Left            =   120
      TabIndex        =   92
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label24 
      Caption         =   "折射率"
      Height          =   255
      Left            =   1320
      TabIndex        =   87
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "材质编辑器："
      Height          =   255
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label37 
      Caption         =   "吸收B"
      Height          =   255
      Left            =   0
      TabIndex        =   84
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label36 
      Caption         =   "吸收G"
      Height          =   255
      Left            =   0
      TabIndex        =   83
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label35 
      Caption         =   "吸收R"
      Height          =   255
      Left            =   0
      TabIndex        =   82
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label34 
      Caption         =   "减少散射B"
      Height          =   255
      Left            =   1320
      TabIndex        =   81
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label32 
      Caption         =   "减少散射G"
      Height          =   255
      Left            =   1320
      TabIndex        =   80
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "不透明度"
      Height          =   255
      Left            =   1560
      TabIndex        =   75
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "高光强度"
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label29 
      Caption         =   "反射TM"
      Height          =   255
      Left            =   1560
      TabIndex        =   73
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label28 
      Caption         =   "反射B"
      Height          =   255
      Left            =   1680
      TabIndex        =   72
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label27 
      Caption         =   "反射G"
      Height          =   255
      Left            =   1680
      TabIndex        =   71
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label DSSDSD 
      Caption         =   "反射R"
      Height          =   255
      Left            =   1680
      TabIndex        =   70
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label26 
      Caption         =   "自发光TM"
      Height          =   255
      Left            =   120
      TabIndex        =   69
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "自发光B"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "自发光G"
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "自发光R"
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "漫射TM"
      Height          =   255
      Left            =   1560
      TabIndex        =   65
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label21 
      Caption         =   "漫射B"
      Height          =   255
      Left            =   1680
      TabIndex        =   64
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "漫射G"
      Height          =   255
      Left            =   1680
      TabIndex        =   63
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label18 
      Caption         =   "漫射R"
      Height          =   255
      Left            =   1680
      TabIndex        =   62
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "透明"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "环境B"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "环境G"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "环境R"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "纹理管理器："
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   10440
      Width           =   1695
   End
   Begin VB.Label bq 
      BackColor       =   &H80000000&
      Caption         =   "无可用动作"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   15960
      TabIndex        =   32
      Top             =   12000
      Width           =   2895
   End
   Begin VB.Label Label11 
      Caption         =   "地形高度"
      Height          =   255
      Left            =   11400
      TabIndex        =   27
      Top             =   12480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "精密度"
      Height          =   255
      Left            =   12840
      TabIndex        =   24
      Top             =   12480
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "地形宽度"
      Height          =   255
      Left            =   10080
      TabIndex        =   23
      Top             =   12480
      Width           =   855
   End
   Begin VB.Label Label8 
      Height          =   1095
      Left            =   4560
      TabIndex        =   14
      Top             =   11880
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "凹凸幅度："
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   12000
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "刷子大小："
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   12000
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "移动速度"
      Height          =   255
      Left            =   9960
      TabIndex        =   10
      Top             =   12000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "密码："
      Height          =   255
      Left            =   13440
      TabIndex        =   9
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":001C
      Height          =   2775
      Left            =   20160
      TabIndex        =   7
      Top             =   10200
      Width           =   2655
   End
   Begin VB.Label Label2 
      Height          =   1095
      Left            =   3000
      TabIndex        =   6
      Top             =   11880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "地图文件："
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   11520
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private land As TVLandscape
Private bDoLoop As Boolean
Private bDoLoop2 As Boolean
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private gb As TVMesh
Private ga As TVMesh
Private gc As Long
Private gd As Long
Private sv As String
Private FRAMES_PER_SECOND As Long
Private Sub Check6_Click()
If Check6.Value Then Text11.Enabled = True: Check5.Enabled = True Else Text11.Enabled = False: Check5.Enabled = False
End Sub
Private Sub load2()
Static ji As Long
Dim n As TextStream
Dim B2() As String
Set fs = CreateObject("Scripting.FileSystemObject")
If ji = 0 Then
If fs.FileExists(Text1.Text) Then
FileCopy Text1.Text, getstrleftb(Text1.Text, "\") + "\" + getstrrightb(Text1.Text, "\") + " 复件." + getstrrightb(Text1.Text, ".")
End If
End If

ji = map.loadmap(scene, Text1.Text, Text2.Text, sv)

If ji = 2 Then
'Command1.Enabled = False
x = map.getter.GetLandRealWidth / 2
z = map.getter.GetLandRealHeight / 2
y = 1000
'map.getter.SetHeight 0, 0, 10
'map.getter.SaveTerrainData "E:\备战作品\易飞翔\sd.ter", TV_LANDSAVE_ALL
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command7.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True
Command14.Enabled = True
Command15.Enabled = True
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
Command21.Enabled = True
Command20.Enabled = True
Form3.Command1.Enabled = True
Form3.Command2.Enabled = True
Form3.Command3.Enabled = True
Form3.Command4.Enabled = True
Form4.Command1.Enabled = True
Form4.Command2.Enabled = True
bDoLoop2 = True
Text1.Enabled = False
Text2.Enabled = False
Form3.loadli
Debug.Print "纹理工作目录  " + getstrleftb(Text1.Text, ".") + "\texture\"
Form4.File1.filename = getstrleftb(Text1.Text, ".") + "\texture\"
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists(getstrleftb(Form1.Text1.Text, ".") + "\texture.dat") Then
Set bb = fs.OpenTextFile(getstrleftb(Form1.Text1.Text, ".") + "\texture.dat")
'If bb.ReadAll = "" Or bb.ReadAll = Chr(10) Or bb.ReadAll = Chr(13) Or bb.ReadAll = Chr(10) + Chr(13) Then
b = Split(bb.ReadAll, Chr(13) + Chr(10), , vbTextCompare)
For c = 0 To UBound(b)
B2 = Split(b(c), "|", , vbTextCompare)
b3 = B2(0) + "\"
    If (UBound(B2) > 1) Then
        If B2(2) = "0" Then
            b3 = b3 + "普通\"
        ElseIf B2(2) = "1" Then
            b3 = b3 + "3D\"
        ElseIf B2(2) = "2" Then
            b3 = b3 + "凹凸\"
        ElseIf B2(2) = "3" Then
            b3 = b3 + "立方\"
        ElseIf B2(2) = "4" Then
            b3 = b3 + "DUVV\"
        ElseIf B2(2) = "5" Then
            b3 = b3 + "Alpha\"
        Else
        End If
    Else
        b3 = b3 + "普通\"
    End If
    If UBound(B2) > 0 Then
        b3 = b3 + B2(1)
    Else
        b3 = b3 + "0"
    End If
List1.AddItem b3
Next
bb.Close
End If
bq.Caption = sv
Else

If ji = 1 Then
bq.Caption = sv
DoEvents
load2
Else
bq.Caption = sv
List1.AddItem ("临时文件，请勿删除！.jpg\普通\0")
fs.CopyFile App.Path + "\a.JPG", getstrleftb(Form1.Text1.Text, ".") + "\texture\临时文件，请勿删除！.jpg", True
End If

End If
End Sub
Private Sub Command1_Click()
Dim bb As TextStream
Dim b() As String
Dim B2() As String
Dim b3 As String
Dim a As Long
Dim b1 As String
load2
If Dir(App.Path & "\maptest", vbDirectory) <> "" Then
 DeleteFolder (App.Path & "\maptest")
 MkDir getstrleftb(Form1.Text1.Text, ".") + "\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\light\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\Mat\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\texture\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\particle\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\sou\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\other\"
b1 = "光工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\light\" + Chr(13) + "材质工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\Mat\" + Chr(13) + "纹理工作目录" + _
getstrleftb(Form1.Text1.Text, ".") + "\texture\" + Chr(13) + "粒子工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\particle\" + Chr(13) + "音乐工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\sou\" _
 + Chr(13) + "文件工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\other\"
Label2.Caption = b1
 Else

 End If
'End If
If sv = "完成！" Then
a = pac.OpenPackage(Form1.Text1.Text)
pac.SetArchivePassword "fsbzwy", a
b = Split(pacstr("mat.dat", a), Chr(13) + Chr(10), , vbTextCompare)
For c = 0 To UBound(b)
B2 = Split(pacstr("Mat\" + b(c) + ".mat", a), Chr(13) + Chr(10), , vbTextCompare)
If CSng(B2(18)) = 0 Then List2.AddItem b(c) + "|0|0|0|0|0|0|0|0" Else List2.AddItem b(c) + "|" + B2(18) + "|" + B2(19) + "|" + B2(20) + "|" + B2(21) + "|" + B2(22) + "|" + B2(23) + "|" + B2(24) + "|" + B2(25)
Next
pac.ClosePackage a
Command9.Enabled = True
Debug.Print "粒子工作目录  " + getstrleftb(Text1.Text, ".") + "\particle\"
Debug.Print "音乐工作目录  " + getstrleftb(Text1.Text, ".") + "\sou\"
MsgBox "打开成功，创建了一个副本防止工程意外破坏"
zr = True
Else
MsgBox sv
End If
End Sub
Private Sub Command10_Click()
If Not List3.ListIndex = -1 Then
getmesh(getstrleftb(List3.List(List3.ListIndex), "|")).SetMaterial (GetMat(getstrlefta(List2.List(List2.ListIndex), "|")))
List3.List(List3.ListIndex) = getstrleftb(List3.List(List3.ListIndex), "|") + "|" + getstrlefta(List2.List(List2.ListIndex), "|")
Else
MsgBox "请选择一个模型！"
End If
End Sub
Private Sub Command11_Click()
Dim a As TV_3DVECTOR
If Not List3.ListIndex = -1 Then
a = getmesh(getstrleftb(List3.List(List3.ListIndex), "|")).GetPosition
scene.GetCamera.SetCamera a.x, a.y, a.z, a.x, a.y, a.z
Else
MsgBox "请选择一个模型！"
End If
End Sub
Private Sub Command12_Click()
Dim a As TV_3DVECTOR
If Not List4.ListIndex = -1 Then
a = GetActor(getstrleftb(List4.List(List4.ListIndex), "|")).GetPosition
scene.GetCamera.SetCamera a.x, a.y, a.z, a.x, a.y, a.z
Else
MsgBox "请选择一个角色！"
End If
End Sub
Private Sub Command13_Click()
If Not List4.ListIndex = -1 Then
GetActor(getstrleftb(List4.List(List4.ListIndex), "|")).SetMaterial (GetMat(getstrlefta(List2.List(List2.ListIndex), "|")))
List4.List(List4.ListIndex) = getstrleftb(List4.List(List4.ListIndex), "|") + "|" + getstrlefta(List2.List(List2.ListIndex), "|")
Else
MsgBox "请选择一个角色！"
End If
End Sub
Private Sub Command14_Click()
Dim a As String
Dim b As String
a = InputBox("请输入一个模型名，模型名必须正确", "添加", "模型")
If a = "" Then Exit Sub
b = InputBox("请输入一个材质名，材质名必须正确", "添加", mat.GetMaterialName(1))
If b = "" Then Exit Sub
command2_click2
List3.AddItem a + "|" + b
getmesh(a).SetMaterial (GetMat(b))
End Sub
Private Sub Command15_Click()
Dim a As String
Dim b As String
a = InputBox("请输入一个角色名，模型名必须正确", "添加", "角色")
If a = "" Then Exit Sub
b = InputBox("请输入一个材质名，材质名必须正确", "添加", mat.GetMaterialName(1))
If b = "" Then Exit Sub
command2_click2
List4.AddItem a + "|" + b
GetActor(a).SetMaterial (GetMat(b))
End Sub
Private Sub Command16_Click()
Dim b As String
Dim c As Long
Dim e As String
'CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
'CommonDialog1.Filter = "WAV波形文件(*.wav)|*.wav|MP3音乐(*.mp3)|*.mp3|OGG音乐(*.ogg)|*.ogg|MIDI音乐(*.mid)|*.mid|FLAC音乐(*.flac)|*.flac|MOD音乐(*.mod)|*.mod|XM音乐(*.xm)|*.xm|IT音乐(*.it)|*.it|S3M音乐(*.s3m)|*.s3m|所有文件(*.*)|*.*"
'CommonDialog1.ShowOpen
'If CommonDialog1.FileTitle = "" Then Exit Sub
'b = CommonDialog1.filename
'e = getstrrightb(b, "\")
'If isinlistex(List6, e) Then MsgBox "含有重复名称的音乐！！": Exit Sub
'Set fi = CreateObject("Scripting.FileSystemObject")
'If Not fi.FileExists(getstrleftb(Form1.Text1.Text, ".") + "\sou\" + e) Then
'FileCopy b, getstrleftb(Form1.Text1.Text, ".") + "\sou\" + e
'End If
command2_click2
List6.AddItem e
sous.AddFile getstrleftb(Form1.Text1.Text, ".") + "\sou\" + e, , TV_SOUNDTYPE_MP3
List6.ListIndex = List6.ListCount - 1
End Sub
Private Sub Command17_Click()
Dim b As String
Dim c As Long
Dim e As String
Dim a As TVParticleSystem
'CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
'CommonDialog1.Filter = "TV3D二进制粒子文件(*.tvp)|*.tvp|所有文件(*.*)|*.*"
'CommonDialog1.ShowOpen
'If CommonDialog1.FileTitle = "" Then Exit Sub
'b = CommonDialog1.filename
'e = getstrrightb(b, "\")
'If isinlist(List5, getstrlefta(e, ".")) = True Then
'Do While isinlist(List5, getstrlefta(e, "."))
'e = "次 " + e
'Loop
'End If
FileCopy b, getstrleftb(Form1.Text1.Text, ".") + "\particle\" + e
Form1.List5.AddItem getstrlefta(e, ".") + "|" + pxx.Text + "|" + pyy.Text + "|" + pzz.Text + "|" + pp.Text
Set a = scene.CreateParticleSystem
a.load b
a.SetGlobalPosition CSng(px.Text), CSng(py.Text), CSng(pz.Text)
a.SetGlobalScale CSng(pxx.Text), CSng(pyy.Text), CSng(pzz.Text), CSng(pp.Text)
a.SetGlobalRotation CSng(pxxx.Text), CSng(pyyy.Text), CSng(pzzz.Text)
map.addpar a, getstrlefta(e, ".")
List5.ListIndex = List5.ListCount - 1
List5_DblClick
'a.Enable False
End Sub
Private Sub Command18_Click()
If Not List5.ListIndex = -1 Then
map.getparformname(getstrlefta(List5.List(List5.ListIndex), "|")).SetGlobalPosition CSng(px.Text), CSng(py.Text), CSng(pz.Text)
map.getparformname(getstrlefta(List5.List(List5.ListIndex), "|")).SetGlobalScale CSng(pxx.Text), CSng(pyy.Text), CSng(pzz.Text), CSng(pp.Text)
map.getparformname(getstrlefta(List5.List(List5.ListIndex), "|")).SetGlobalRotation CSng(pxxx.Text), CSng(pyyy.Text), CSng(pzzz.Text)
List5.List(List5.ListIndex) = getstrlefta(List5.List(List5.ListIndex), "|") + "|" + pxx.Text + "|" + pyy.Text + "|" + pzz.Text + "|" + pp.Text
Else
MsgBox "请选择一个粒子！"
End If
End Sub
Private Sub Command19_Click()
If Not List5.ListIndex = -1 Then
scene.GetCamera.SetCamera CSng(px.Text), CSng(py.Text), CSng(pz.Text), CSng(px.Text), CSng(py.Text), CSng(pz.Text)
Else
MsgBox "请选择一个粒子！"
End If
End Sub
Private Sub command2_click2()
Dim q As String
If zr Then
q = getstrleftb(Text1.Text, ".") + "\"
map.getter.SaveTerrainData q + "land.ter", TV_LANDSAVE_ALL
'Set a = CreateObject("File System Object.TextStream")
Form3.savelig q
savetex
savemat
savepar
savesou
saveoth
savejb
MsgBox "为了您的安全,已经保存成功！将在目录下创建文件夹，如果有错误可以使用复件"
Else
End If
's = pac.OpenPackage("C:\Documents and Settings\琉璃\桌面\E3DPackage2.pak")
'Debug.Print "s"
'pac.SetArchivePassword "fsbzwy", s
'Debug.Print "ss"
'Debug.Print scene.SaveTVS(s)
'pac.ClosePackage s
'Debug.Print "sas"
'debug.Print getflaleft(Text1.Text) + "\land.ter"

End Sub
Private Sub Command2_Click()
MsgBox "请确保原目录所有必要的子目录已建立好"
Dim q As String
q = getstrleftb(Text1.Text, ".") + "\"
map.getter.SaveTerrainData q + "land.ter", TV_LANDSAVE_ALL
'Set a = CreateObject("File System Object.TextStream")
Form3.savelig q
savetex
savemat
savepar
savesou
saveoth
savejb
's = pac.OpenPackage("C:\Documents and Settings\琉璃\桌面\E3DPackage2.pak")
'Debug.Print "s"
'pac.SetArchivePassword "fsbzwy", s
'Debug.Print "ss"
'Debug.Print scene.SaveTVS(s)
'pac.ClosePackage s
'Debug.Print "sas"
'debug.Print getflaleft(Text1.Text) + "\land.ter"
MsgBox "已经保存成功！将在目录下创建文件夹，如果有错误可以使用复件"
End Sub
Sub savejb()
'If jbwj.Text <> "" Then
'Set fs = CreateObject("Scripting.FileSystemObject")
'If fs.FileExists(jbwj.Text) Then
'fs.CopyFile jbwj.Text, getstrleftb(Text1.Text, ".") + "\game.lua", True
'End If
'End If
End Sub
Sub saveoth()
Dim b() As String
Dim bb As TextStream
Dim a As Long
If List7.ListCount <> 0 Then
Set fs = CreateObject("Scripting.FileSystemObject")
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\other.dat", True)
For a = 0 To List7.ListCount - 1
If a = List7.ListCount - 1 Then
bb.Write List7.List(a)
Else
bb.WriteLine List7.List(a)
End If
Next
bb.Close
End If
End Sub
Sub savesou()
Dim b() As String
Dim bb As TextStream
Dim a As Long
If List6.ListCount <> 0 Then
Set fs = CreateObject("Scripting.FileSystemObject")
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\sou.dat", True)
For a = 0 To List6.ListCount - 1
If a = List6.ListCount - 1 Then
bb.Write List6.List(a)
Else
bb.WriteLine List6.List(a)
End If
Next
bb.Close
End If
End Sub
Sub savepar()
Dim b() As String
Dim bb As TextStream
Dim a As Long
If List5.ListCount <> 0 Then
Set fs = CreateObject("Scripting.FileSystemObject")
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\particle.dat", True)
For a = 0 To List5.ListCount - 1
List5.ListIndex = a
List5_DblClick
If a = List5.ListCount - 1 Then
bb.Write getstrlefta(List5.List(a), "|") + ".tvp|" + px.Text + "|" + py.Text + "|" + pz.Text + "|" + pxx.Text + "|" + pyy.Text + "|" + pzz.Text + "|" + pp.Text + "|" + pxxx.Text + "|" + pyyy.Text + "|" + pzzz.Text
Else
bb.WriteLine getstrlefta(List5.List(a), "|") + ".tvp|" + px.Text + "|" + py.Text + "|" + pz.Text + "|" + pxx.Text + "|" + pyy.Text + "|" + pzz.Text + "|" + pp.Text + "|" + pxxx.Text + "|" + pyyy.Text + "|" + pzzz.Text
End If
Next
bb.Close
End If
End Sub
Sub savemat()
Dim b() As String
Dim bb As TextStream
If List2.ListCount <> 0 Then
Set fs = CreateObject("Scripting.FileSystemObject")
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\mat.dat", True)
For c = 0 To List2.ListCount - 1
If c = List2.ListCount - 1 Then
b = Split(List2.List(c), "|", , vbTextCompare)
bb.Write b(0)
Else
b = Split(List2.List(c), "|", , vbTextCompare)
bb.WriteLine b(0)
End If
Next
bb.Close
For c = 0 To List2.ListCount - 1
List2.ListIndex = c
List2_DblClick
b = Split(List2.List(c), "|", , vbTextCompare)
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\mat\" + b(0) + ".mat", True)
bb.WriteLine R.Text
bb.WriteLine G.Text
bb.WriteLine Bbjk.Text
bb.WriteLine TM.Text
bb.WriteLine R2.Text
bb.WriteLine G2.Text
bb.WriteLine B2.Text
bb.WriteLine tm2.Text
bb.WriteLine r3.Text
bb.WriteLine g3.Text
bb.WriteLine b3.Text
bb.WriteLine tm3.Text
bb.WriteLine r4.Text
bb.WriteLine g4.Text
bb.WriteLine b4.Text
bb.WriteLine tm4.Text
bb.WriteLine gg.Text
bb.WriteLine btm.Text
If prt.Value = 1 Then bb.WriteLine CStr(prt.Value): bb.WriteLine CStr(ssr.Text): bb.WriteLine CStr(ssg.Text): bb.WriteLine CStr(ssb.Text): bb.WriteLine CStr(xsr.Text): bb.WriteLine CStr(xsg.Text): bb.WriteLine CStr(xsb.Text): bb.Write CStr(zs.Text) Else bb.Write CStr(prt.Value)
bb.Close
Next
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\mat\Mesh.dat", True)
For c = 0 To List3.ListCount - 1
If c = List3.ListCount - 1 Then
bb.Write List3.List(c)
Else
bb.WriteLine List3.List(c)
End If
Next
bb.Close
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\mat\NPC.dat", True)
For c = 0 To List4.ListCount - 1
If c = List4.ListCount - 1 Then
bb.Write List4.List(c)
Else
bb.WriteLine List4.List(c)
End If
Next
bb.Close
End If
End Sub
Private Sub Command20_Click()
If Not List3.ListIndex = -1 Then
List3.RemoveItem List3.ListIndex
getmesh(getstrlefta(List3.List(List3.ListIndex), "|")).SetMaterial mat.GetMaterialName(1)
Else
MsgBox "请选择一个模型"
End If
End Sub
Private Sub Command21_Click()
If Not List4.ListIndex = -1 Then
List4.RemoveItem List4.ListIndex
GetActor(getstrlefta(List4.List(List4.ListIndex), "|")).SetMaterial mat.GetMaterialName(1)
Else
MsgBox "请选择一个角色！"
End If
End Sub
Private Sub Command22_Click()
px.Text = scene.GetCamera.GetPosition.x
py.Text = scene.GetCamera.GetPosition.y
pz.Text = scene.GetCamera.GetPosition.z
End Sub
Private Sub Command23_Click()
If Not List6.ListIndex = -1 Then
Kill getstrleftb(Form1.Text1.Text, ".") + "\sou\" + List6.List(List6.ListIndex)
sous.Remove getstrlefta(List6.List(List6.ListIndex), ".")
List6.RemoveItem List6.ListIndex
Else
MsgBox "请选择一个音乐！"
End If
End Sub
Private Sub Command24_Click()
If Not List6.ListIndex = -1 Then
Debug.Print getstrlefta(List6.List(List6.ListIndex), ".")
'sous(getstrlefta(List6.List(List6.ListIndex), ".")).Play
Else
MsgBox "请选择一个音乐！"
End If
End Sub
Private Sub Command25_Click()
Dim b As String
MsgBox "如果这是新工程，请创建全部必要的工作目录，如果是旧工程，请不要创建，否则程序会出错！！"
If MsgBox("是否创建所有的工程目录？", vbOKCancel) = vbOK Then
MkDir getstrleftb(Form1.Text1.Text, ".") + "\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\light\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\Mat\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\texture\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\particle\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\sou\"
MkDir getstrleftb(Form1.Text1.Text, ".") + "\other\"
b = "光工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\light\" + Chr(13) + "材质工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\Mat\" + Chr(13) + "纹理工作目录" + _
getstrleftb(Form1.Text1.Text, ".") + "\texture\" + Chr(13) + "粒子工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\particle\" + Chr(13) + "音乐工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\sou\" _
 + Chr(13) + "文件工作目录" + getstrleftb(Form1.Text1.Text, ".") + "\other\"
Clipboard.SetText b
MsgBox "已创建完毕！！将路径拷贝到了剪贴板" + Chr(13) + b
End If
End Sub
Private Sub Command26_Click()
'sous(getstrlefta(List6.List(List6.ListIndex), ".")).Pause
End Sub
Private Sub Command27_Click()
'sous(getstrlefta(List6.List(List6.ListIndex), ".")).Stop_
End Sub
Private Sub Command28_Click()
MsgBox "TV3D可视化编辑器--小熊制作，版权所有，翻版必究,小熊QQ:1066562980"
End Sub
Private Sub Command29_Click()
'CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
'CommonDialog1.Filter = "所有文件(*.*)|*.*"
'CommonDialog1.ShowOpen
'If CommonDialog1.FileTitle = "" Then Exit Sub
'List7.AddItem (getstrrightb(CommonDialog1.FileTitle, "\"))
'Set fi = CreateObject("Scripting.FileSystemObject")
'If Not fi.FileExists(getstrleftb(Form1.Text1.Text, ".") + "\other\" + getstrrightb(CommonDialog1.FileTitle, "\")) Then
'FileCopy CommonDialog1.FileTitle, getstrleftb(Form1.Text1.Text, ".") + "\other\" + getstrrightb(CommonDialog1.FileTitle, "\")
'End If
End Sub
Sub qw()
Dim G As Long
Dim gg As Long
''debug.Print Form2.Text1.Text, Form2.Text2.Text, Form2.Text3.Text, Form2.Text4.Text, Form2.Text5.Text, Form2.Text6.Text
If xwg Then
    For gg = Fix(-(y / yy)) - 1 To Fix(y / yy) + 1
    For G = Fix(-(x / xx)) - 1 To Fix(x / xx) + 1
    scr2D.Draw_Line3D G * xx, yy * gg, -z, G * xx, yy * gg, z, red
    Next
    Next
End If
If zwg Then
    For gg = Fix(-(y / yy)) - 1 To Fix(y / yy) + 1
    For G = Fix(-(z / zz)) - 1 To Fix(z / zz) + 1
    scr2D.Draw_Line3D -x, yy * gg, G * zz, x, yy * gg, G * zz, red
    Next
    Next
End If
If ywg Then
    For gg = Fix(-(x / xx)) - 1 To Fix(x / xx) + 1
    For G = Fix(-(z / zz)) - 1 To Fix(z / zz) + 1
    scr2D.Draw_Line3D xx * gg, -y, zz * G, xx * gg, y, zz * G, red
    Next
    Next
End If

'for gg=clng(
'scr2D.darw_li
End Sub
Private Sub Command3_Click()
Command25_Click
Dim a(0) As String
a(0) = ""
Command1.Enabled = False
map.createland scene, ""
map.createsky
map.createmeshnpcarray 0, 0, scene, a, a
If Check6.Value = 1 Then
map.getter.GenerateTerrain Text11.Text, CLng(Text9.Text), CLng(Text7.Text), CLng(Text10.Text), 0, 0, 0, IIf(Check5.Value = 1, True, False): scene.SetCamera 0, 50 + map.getter.GetHeight(0, 0), 0, 0, 0, 0
'debug.Print map.getter.GetLandRealWidth, map.getter.GetLandRealHeight
map.getter.SetPosition -map.getter.GetLandRealWidth() / 2, 0, -map.getter.GetLandRealHeight() / 2
'debug.Print map.getter.GetLandRealWidth, map.getter.GetLandRealHeight
map.getter.SetTexture gc
Else
map.getter.CreateEmptyTerrain CLng(Text9.Text), CLng(Text7.Text), CLng(Text10.Text), 0, 0, 0: 'debug.Print "创建空地形": scene.SetCamera 0, 50, 0, 0, 0, 0
map.getter.SetPosition -map.getter.GetLandRealWidth() / 2, 0, -map.getter.GetLandRealHeight() / 2
map.getter.SetTexture gc
End If
x = map.getter.GetLandRealWidth / 2
z = map.getter.GetLandRealHeight / 2
y = 1000
xwg = True
ywg = True
zwg = True
bDoLoop2 = True
Command2.Enabled = True
Command3.Enabled = False
Command7.Enabled = True
Command8.Enabled = True
Form3.Command1.Enabled = True
Form3.Command2.Enabled = True
Form3.Command3.Enabled = True
Form3.Command4.Enabled = True
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
Form4.Command1.Enabled = True
Form4.Command2.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Command9.Enabled = True
MsgBox "你可以保存地形，将其复制到地图包中然后再用此软件编辑"
MsgBox "创建新工程成功！"
scene.GetCamera.SetCamera -200, 100, 0, 200, 0, 0
End Sub
Private Sub Command30_Click()
'List7.RemoveItem (List7.ListIndex)
'Set fi = CreateObject("Scripting.FileSystemObject")
'If fi.FileExists(getstrleftb(Form1.Text1.Text, ".") + "\other\" + List7.List(List7.ListIndex)) Then
'Kill getstrleftb(Form1.Text1.Text, ".") + "\other\" + getstrrightb(CommonDialog1.FileTitle, "\")
'End If
End Sub
Private Sub Command31_Click()
'CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
'CommonDialog1.Filter = "lua脚本文件(*.lua)|*.lua|所有文件(*.*)|*.*"
'CommonDialog1.ShowOpen
'If CommonDialog1.FileTitle = "" Then Exit Sub
'Set fi = CreateObject("Scripting.FileSystemObject")
'fi.CopyFile CommonDialog1.filename, getstrleftb(Text1.Text, "\") + "\Main.lua", True
'MsgBox "已经拷贝到工程目录下！"
End Sub
Private Sub Command32_Click()
'CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
'CommonDialog1.Filter = "lua脚本文件(*.lua)|*.lua|所有文件(*.*)|*.*"
'CommonDialog1.ShowOpen
'If CommonDialog1.FileTitle = "" Then Exit Sub
'Set fi = CreateObject("Scripting.FileSystemObject")
'fi.CopyFile CommonDialog1.filename, getstrleftb(Text1.Text, "\") + "\Render.lua", True
'MsgBox "已经拷贝到工程目录下！"
End Sub
Private Sub Command33_Click()
'CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
'CommonDialog1.Filter = "lua脚本文件(*.lua)|*.lua|所有文件(*.*)|*.*"
'CommonDialog1.ShowOpen
'If CommonDialog1.FileTitle = "" Then Exit Sub
'Set fi = CreateObject("Scripting.FileSystemObject")
'fi.CopyFile CommonDialog1.filename, getstrleftb(Text1.Text, "\") + "\Estroy.lua", True
'MsgBox "已经拷贝到工程目录下！"
End Sub
Private Sub Command34_Click()
'CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
'CommonDialog1.Filter = "lua脚本文件(*.lua)|*.lua|所有文件(*.*)|*.*"
'CommonDialog1.ShowOpen
'If CommonDialog1.FileTitle = "" Then Exit Sub
'Set fi = CreateObject("Scripting.FileSystemObject")
'fi.CopyFile CommonDialog1.filename, getstrleftb(Text1.Text, "\") + "\Logic.lua", True
'MsgBox "已经拷贝到工程目录下！"
End Sub
Private Sub Command4_Click()
Form2.Show
End Sub
Private Sub Command5_Click()
If zr = True Then map.destroy
End
End Sub
Private Sub Command6_Click()
Form4.Show
Form4.File1.filename = getstrleftb(Text1.Text, ".") + "\texture\"
End Sub
Private Sub savetex()
Dim b() As String
Dim bb As TextStream
If List1.ListCount <> 0 Then
Set fs = CreateObject("Scripting.FileSystemObject")
Set bb = fs.CreateTextFile(getstrleftb(Form1.Text1.Text, ".") + "\texture.dat", True)
For c = 0 To List1.ListCount - 1
b = Split(List1.List(c), "\", , vbTextCompare)
If b(1) = "普通" Then
b(1) = "0"
ElseIf b(1) = "3D" Then
b(1) = "1"
ElseIf b(1) = "凹凸" Then
b(1) = "2"
ElseIf b(1) = "立方" Then
b(1) = "3"
ElseIf b(1) = "DUVU" Then
b(1) = "4"
ElseIf b(1) = "Alpha" Then
b(1) = "5"
Else
End If
If c = List1.ListCount - 1 Then bb.Write b(0) + "|" + b(2) + "|" + b(1) Else bb.WriteLine b(0) + "|" + b(2) + "|" + b(1)
Next
bb.Close
End If
End Sub
Private Sub Command7_Click()
Dim a As Long
a = mat.CreateMaterialQuick(CSng(R.Text), CSng(G.Text), CSng(Bbjk.Text), CSng(TM.Text), cz.Text)
mat.SetAmbient a, CSng(R.Text), CSng(G.Text), CSng(Bbjk.Text), CSng(TM.Text)
mat.SetDiffuse a, CSng(R2.Text), CSng(G2.Text), CSng(B2.Text), CSng(tm2.Text)
mat.SetEmissive a, CSng(r3.Text), CSng(g3.Text), CSng(b3.Text), CSng(tm3.Text)
mat.SetSpecular a, CSng(r4.Text), CSng(g4.Text), CSng(b4.Text), CSng(tm4.Text)
mat.SetPower a, CSng(gg.Text)
mat.SetOpacity a, CSng(btm.Text)
mat.EnablePRTSubSurface a, IIf(prt.Value = 1, True, False)
If prt.Value = 1 Then
mat.SetPRTSubSurfReducedScattering a, CSng(ssr.Text), CSng(ssg.Text), CSng(ssb.Text)
mat.SetPRTSubSurfAbsorption a, CSng(xsr.Text), CSng(xsg.Text), CSng(xsb.Text)
mat.SetPRTSubSurfRefractionIndexRatio a, CSng(zs.Text)
List2.AddItem cz.Text + "|" + CStr(prt.Value) + "|" + ssr.Text + "|" + ssg.Text + "|" + ssb.Text + "|" + xsr.Text + "|" + xsg.Text + "|" + xsb.Text + "|" + zs.Text
Else
ssr.Text = "0"
ssg.Text = "0"
ssb.Text = "0"
xsr.Text = "0"
xsg.Text = "0"
xsb.Text = "0"
zs.Text = "0"
List2.AddItem cz.Text + "|" + CStr(prt.Value) + "|" + ssr.Text + "|" + ssg.Text + "|" + ssb.Text + "|" + xsr.Text + "|" + xsg.Text + "|" + xsb.Text + "|" + zs.Text
End If
List2.ListIndex = List2.ListCount - 1
List2_DblClick
End Sub
Private Sub Command8_Click()
Dim b() As String
If Not List2.ListIndex = -1 Then
b = Split(List2.List(List2.ListIndex), "|", , vbTextCompare)
mat.DeleteMaterial GetMat(b(0))
List2.RemoveItem List2.ListIndex
List2.ListIndex = 0
List2_DblClick
Else
MsgBox "请选择一个材质！"
End If
End Sub
Private Sub Command9_Click()
Dim a As Long
Dim b() As String
If Not List2.ListIndex = -1 Then
b = Split(List2.List(List2.ListIndex), "|", , vbTextCompare)
a = GetTex(b(0))
mat.SetAmbient a, CSng(R.Text), CSng(G.Text), CSng(Bbjk.Text), CSng(TM.Text)
mat.SetDiffuse a, CSng(R2.Text), CSng(G2.Text), CSng(B2.Text), CSng(tm2.Text)
mat.SetEmissive a, CSng(r3.Text), CSng(g3.Text), CSng(b3.Text), CSng(tm3.Text)
mat.SetSpecular a, CSng(r4.Text), CSng(g4.Text), CSng(b4.Text), CSng(tm4.Text)
mat.SetPower a, CSng(gg.Text)
mat.SetOpacity a, CSng(btm.Text)
mat.EnablePRTSubSurface a, IIf(prt.Value = 1, True, False)
If prt.Value = 1 Then
mat.SetPRTSubSurfReducedScattering a, CSng(ssr.Text), CSng(ssg.Text), CSng(ssb.Text)
mat.SetPRTSubSurfAbsorption a, CSng(xsr.Text), CSng(xsg.Text), CSng(xsb.Text)
mat.SetPRTSubSurfRefractionIndexRatio a, CSng(zs.Text)
List2.List(List2.ListIndex) = cz.Text + "|" + CStr(prt.Value) + "|" + ssr.Text + "|" + ssg.Text + "|" + ssb.Text + "|" + xsr.Text + "|" + xsg.Text + "|" + xsb.Text + "|" + zs.Text
Else
ssr.Text = "0"
ssg.Text = "0"
ssb.Text = "0"
xsr.Text = "0"
xsg.Text = "0"
xsb.Text = "0"
zs.Text = "0"
List2.List(List2.ListIndex) = cz.Text + "|" + CStr(prt.Value) + "|" + ssr.Text + "|" + ssg.Text + "|" + ssb.Text + "|" + xsr.Text + "|" + xsg.Text + "|" + xsb.Text + "|" + zs.Text
End If
List2_DblClick
Else
MsgBox "请选择一个材质！"
End If
End Sub
Private Sub Form_Load()
If RegSvr32("c:\windows\system32\TVb3D65.dll", False) = False Then
RegSvr32 App.Path & "\TVb3D65.dll", False
End If
Dim aaaa As Long
iniall
Text1.Text = App.Path + "\maptest.map"
Set gb = scene.CreateMeshBuilder("ghjkl")
gb.LoadTVM App.Path + "\箭头.TVM"
ReDim ge(0 To 0)
Set ge(0) = scene.CreateBillboard(tex.LoadTexture(App.Path + "\a.jpg", , , , black), 0, 0, 0, 12, 12, "qq", True)
ge(0).SetBillboardType TV_BILLBOARD_FREEROTATION
ge(0).SetAlphaTest True, 64
Set gee = scene.CreateBillboard(tex.LoadTexture(App.Path + "\a.jpg", , , , black), 0, 0, 0, 12, 12, "q", True)
gee.SetBillboardType TV_BILLBOARD_FREEROTATION
gee.SetAlphaTest True, 64
Set ga = scene.CreateMeshBuilder("ghjkl2")
ga.SetColor red
Set gef = scene.CreateBillboard(tex.LoadTexture(App.Path + "\b.jpg", , , , black), 0, 0, 0, 12, 12, "bb", True)
gc = tex.LoadTexture(App.Path + "\箭头.png")
xx = 100
zz = 100
yy = 100
xwg = False
ywg = False
zwg = False
kg2 = True
kg3 = True
kg4 = True
gd = lig.CreatePointLight(cre3d(0, 0, 0), 255, 0, 0, 200)
'gb.SetLightingMode TV_LIGHTING_MANAGED
bDoLoop = True
Form1.Show
Form3.Show
Form4.Show
a
End Sub
Private Sub a()
  Static SKIP_TICKS As Long
  Static next_game_tick As Single
  Static sleep_time As Long
  FRAMES_PER_SECOND = 60
  SKIP_TICKS = 1000 / FRAMES_PER_SECOND
  next_game_tick = tv.TickCount
 Do While bDoLoop
    If GetFocus = Picture1.hwnd Then
    If bDoLoop2 Then B拜拜
        tv.Clear False
            If bDoLoop2 Then
            map.Render
            qw
            gb.Render
            ga.Render
            map.SubRender
            If kg2 Then
                For ass = 0 To UBound(ge)
                ge(ass).Render
                Next
            End If
            If kg3 And zr Then
                For ass = 0 To UBound(pa)
                pa(ass).Render
                Next
            End If
            Else
            End If
        tv.RenderToScreen
        If inp.IsKeyPressed(TV_KEY_ESCAPE) Then bDoLoop = False
    Else
    Sleep 100
    End If
    DoEvents
    next_game_tick = next_game_tick + SKIP_TICKS
    sleep_time = next_game_tick - tv.TickCount
    If sleep_time >= 0 Then Sleep sleep_time
  Loop
desall
End
End Sub
Sub B拜拜()
Static x As Long
Static y As Long
Static x2 As Single
Static y2 As Single
Static a As Long
Static b As Boolean
Static c As Boolean
Static d As Boolean
Static x3 As Long
Static y3 As Long
Dim w As Single
Dim e As TV_3DVECTOR
Dim f As TV_2DVECTOR
Dim f2 As TV_2DVECTOR
Static aa As Boolean
Static bb As Boolean
Static bbaq As Single
Static asdw As Long
Dim aqw As Boolean
Dim ligt As TV_LIGHT
Dim vbb As Long
Static xzz As Boolean
Dim ligq As TV_LIGHT
Static jkl As Boolean
Static jkl2 As Boolean
Dim str As String
Dim stre As String
Dim pat As TVParticleSystem
Dim wgn As Boolean
inp.GetMouseState x, y, retroll:=a, retbutton1:=aa, retbutton2:=b, retbutton3:=bb
 ''debug.Print bb
 inp.GetAbsMouseState x3, y3
 If b Then
 If c = False Then c = True: d = Not (d)
 If aa = False Then bbaq = 0
 c = True
 Else
 c = False
 End If
 x2 = x2 - (y / 5)
 y2 = y2 - (x / 5)
 If inp.IsKeyPressed(TV_KEY_F1) Then Check4.Value = 1 Else Check4.Value = 0
 If aa And bb Then Text6.Text = CStr(CLng(Text6.Text) + a / 100): aqw = True Else If bb Then Text4.Text = CStr(CLng(Text4.Text) + a / 100) Else If Check4.Value = 0 Then Text3.Text = CStr(CLng(Text3.Text) + a / 100)
 If x2 > 100 Then x2 = 100
 If x2 < -100 Then x2 = -100
 If d Then scene.GetCamera.SetRotation -x2, -y2, 0
 lig.EnableLight gd, Not (xzz)
 If xzz = False Then
 If inp.IsKeyPressed(TV_KEY_W) Then scene.GetCamera.MoveRelative CSng(Text3.Text), 0, 0, True
 If inp.IsKeyPressed(TV_KEY_S) Then scene.GetCamera.MoveRelative -CSng(Text3.Text), 0, 0, True
 If inp.IsKeyPressed(TV_KEY_A) Then scene.GetCamera.MoveRelative 0, 0, -CSng(Text3.Text), True
 If inp.IsKeyPressed(TV_KEY_D) Then scene.GetCamera.MoveRelative 0, 0, CSng(Text3.Text), True
 If inp.IsKeyPressed(TV_KEY_Q) Then scene.GetCamera.MoveRelative 0, CSng(Text3.Text), 0, True
 If inp.IsKeyPressed(TV_KEY_E) Then scene.GetCamera.MoveRelative 0, -CSng(Text3.Text), 0, True
 e = map.getter.MousePick(x3, y3).GetCollisionImpact
 End If
 bq.Caption = "无可用动作"
 For piu = 0 To UBound(ge)
 ge(piu).ShowBoundingBox False
 Next
 xzz = False
 Set col = scene.MousePick(x3, y3, 0, TV_TESTTYPE_ACCURATETESTING)
 lig.SetLightPosition gd, col.GetCollisionImpact.x, col.GetCollisionImpact.y + 20, col.GetCollisionImpact.z
    If col.GetCollisionObjectType = TV_OBJECT_MESH And kg4 = True Then
    If col.GetCollisionMesh.GetMeshName = "ghjkl" Or col.GetCollisionMesh.GetMeshName = "ghjkl2" Then wgn = True
    End If
    If map.getter.MousePick(x3, y3).IsCollision = True And kg4 = True Or wgn = True Then
    bq.Caption = "可以管理地形"
    Label8.Caption = "地形X:" & e.x & Chr(13) & "地形Y:" & e.y & Chr(13) & "地形Z:" & e.z
    scr2D.Draw_Box3D cre3d(e.x - 10, e.y - 1, e.z - 10), cre3d(e.x + 10, e.y + 1, e.z + 10), red
    ga.Resetmesh
    ga.AddFloor gc, e.x - CLng(Text4.Text) / 2, e.z - CLng(Text4.Text) / 2, e.x + CLng(Text4.Text) / 2, e.z + CLng(Text4.Text) / 2, e.y
    If CSng(Text6.Text) < 0 Then gb.SetRotation 0, 0, 180: gb.SetPosition e.x, e.y + 60, e.z Else gb.SetRotation 0, 0, 0: gb.SetPosition e.x, e.y, e.z
    'scr2D.Draw_Box3D cre3d(e.x - CLng(Text4.Text) / 2, e.y - 2, e.z - CLng(Text4.Text) / 2), cre3d(e.x + CLng(Text4.Text) / 2, e.y + 2, e.z + CLng(Text4.Text) / 2), red
    If Check4.Value = 1 Then asdw = asdw + a: scene.GetCamera.SetCamera e.x, e.y + asdw + 20, e.z, e.x, e.y, e.z: Sleep 20
    If aa And Not (aqw) Then
        If Check1.Value = 1 Then
            For qqq = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                For qqq2 = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                map.getter.SetHeight e.x + qqq, e.z + qqq2, CSng(Text5.Text) / 100
                Next
            Next
        ElseIf Check2.Value = 1 Then
            bbaq = bbaq + CSng(Text8.Text) / 10000
            If Check3.Value = 1 Then
                For qqq = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                    For qqq2 = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                    map.getter.SetHeight e.x + qqq, e.z + qqq2, bbaq, brelative:=True
                    Next
                Next
            Else
                For qqq = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                    For qqq2 = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                    map.getter.SetHeight e.x + qqq, e.z + qqq2, bbaq
                    Next
                Next
            End If
        Else
            For qqq = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                For qqq2 = -CLng(Text4.Text) / 2 To CLng(Text4.Text) / 2
                map.getter.SetHeight e.x + qqq, e.z + qqq2, CSng(Text6.Text) / 100, brelative:=True
                Next
            Next
        End If
    End If
'    For w = 0 To 360
'        f = lineang(CSng(Text4.Text), w, e.X, e.z)
'        f2 = lineang(CSng(Text4.Text) + 1, w, e.X, e.z)
'        scr2D.Draw_Line3D f.X, map.getter.GetHeight(f.X, f.Y) + 10, f.Y, f2.X, map.getter.GetHeight(f2.X, f2.Y) + 10, f2.Y
'    Next

 ElseIf col.GetCollisionObjectType = TV_OBJECT_MESH Then
    str = col.GetCollisionMesh.GetMeshName
    bq.Caption = "不能管理的模型:" + str + "，可以添加入列表"
    If isinlist(List3, str) Then
    bq.Caption = "可以管理模型:" + str
        If aa = True Then
        bq.Caption = "管理模型中:" + str + "，可以改变其材质"
        str2 = str
        List3.ListIndex = inlist(List3, str)
        End If
    ElseIf InStrRev(str, "bb|") <> 0 Then
        bq.Caption = "可以管理粒子" + getstrrightb(str, "|")
        If aa Then
            If isinlist(List5, getstrrightb(str, "|")) Then
            List5.ListIndex = inlist(List5, getstrrightb(str, "|"))
            bq.Caption = "正在管理粒子,wsadqe控制位置,zxcv缩放,空格按格子移动,单位除以100" + getstrrightb(str, "|")
            Set pat = map.getparformname(getstrrightb(str, "|"))
            If inp.IsKeyPressed(TV_KEY_SPACE) Then
                If inp.IsKeyPressed(TV_KEY_W) Then
                    If jkl2 = False Then
                                    Debug.Print "sdsad"
                        map.getparformname(getstrrightb(str, "|")).SetGlobalPosition getgz(pat.GetGlobalPosition.x, 1, True), pat.GetGlobalPosition.y, pat.GetGlobalPosition.z
                        jkl2 = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_S) Then
                    If jkl2 = False Then
                        map.getparformname(getstrrightb(str, "|")).SetGlobalPosition getgz(pat.GetGlobalPosition.x, 1, False), pat.GetGlobalPosition.y, pat.GetGlobalPosition.z
                        jkl2 = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_A) Then
                    If jkl2 = False Then
                        map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, pat.GetGlobalPosition.y, getgz(pat.GetGlobalPosition.z, 3, True)
                        jkl2 = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_D) Then
                    If jkl2 = False Then
                        map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, pat.GetGlobalPosition.y, getgz(pat.GetGlobalPosition.z, 3, False)
                        jkl2 = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_Q) Then
                    If jkl2 = False Then
                        map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, getgz(pat.GetGlobalPosition.y, 2, True), pat.GetGlobalPosition.z
                        jkl2 = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_E) Then
                    If jkl2 = False Then
                        map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, getgz(pat.GetGlobalPosition.y, 2, False), pat.GetGlobalPosition.z
                        jkl2 = True
                    End If
                Else
                End If
            Else
                jkl2 = False
                If inp.IsKeyPressed(TV_KEY_W) Then map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x + CSng(Text3.Text), pat.GetGlobalPosition.y, pat.GetGlobalPosition.z
                If inp.IsKeyPressed(TV_KEY_S) Then map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x - CSng(Text3.Text), pat.GetGlobalPosition.y, pat.GetGlobalPosition.z
                If inp.IsKeyPressed(TV_KEY_A) Then map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, pat.GetGlobalPosition.y, pat.GetGlobalPosition.z + CSng(Text3.Text)
                If inp.IsKeyPressed(TV_KEY_D) Then map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, pat.GetGlobalPosition.y, pat.GetGlobalPosition.z - CSng(Text3.Text)
                If inp.IsKeyPressed(TV_KEY_Q) Then map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, pat.GetGlobalPosition.y + CSng(Text3.Text), pat.GetGlobalPosition.z
                If inp.IsKeyPressed(TV_KEY_E) Then map.getparformname(getstrrightb(str, "|")).SetGlobalPosition pat.GetGlobalPosition.x, pat.GetGlobalPosition.y - CSng(Text3.Text), pat.GetGlobalPosition.z
            End If
            If inp.IsKeyPressed(TV_KEY_Z) Then pxx.Text = CSng(pxx.Text) + CSng(Text3.Text) / 100
            If inp.IsKeyPressed(TV_KEY_X) Then pyy.Text = CSng(pyy.Text) + CSng(Text3.Text) / 100
            If inp.IsKeyPressed(TV_KEY_C) Then pzz.Text = CSng(pzz.Text) + CSng(Text3.Text) / 100
            If inp.IsKeyPressed(TV_KEY_V) Then pp.Text = CSng(pp.Text) + CSng(Text3.Text) / 100
            map.getparformname(getstrrightb(str, "|")).SetGlobalScale CSng(pxx.Text), CSng(pyy.Text), CSng(pzz.Text), CSng(pp.Text)
            List5.List(List5.ListIndex) = getstrlefta(List5.List(List5.ListIndex), "|") + "|" + pxx.Text + "|" + pyy.Text + "|" + pzz.Text + "|" + pp.Text
            List5_DblClick
            xzz = True
            Else
            bq.Caption = "无效粒子" + getstrrightb(str, "|")
            End If
        End If
    ElseIf lig.IsLightActive(GetLight(getstrrightb(str, "\"))) And InStrRev(str, "qq\") <> 0 Then
        bq.Caption = "可以管理光:" + getstrrightb(str, "\")
        If aa = True Then
            bq.Caption = "正在管理光:" + getstrrightb(str, "\") + "，请按wsadqe键移动，空格按格子移动"
            col.GetCollisionMesh.ShowBoundingBox True, red
            xzz = True
            If inp.IsKeyPressed(TV_KEY_SPACE) Then
                If inp.IsKeyPressed(TV_KEY_W) Then
                    If jkl = False Then
                        lig.SetLightPosition GetLight(getstrrightb(str, "\")), getgz(ligq.position.x, 1, True), ligq.position.y, ligq.position.z
                        Form3.List1_DblClick
                        jkl = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_S) Then
                    If jkl = False Then
                        lig.SetLightPosition GetLight(getstrrightb(str, "\")), getgz(ligq.position.x, 1, False), ligq.position.y, ligq.position.z
                        Form3.List1_DblClick
                        jkl = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_A) Then
                    If jkl = False Then
                        lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, ligq.position.y, getgz(ligq.position.z, 3, True)
                        Form3.List1_DblClick
                        jkl = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_D) Then
                    If jkl = False Then
                        lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, ligq.position.y, getgz(ligq.position.z, 3, False)
                        Form3.List1_DblClick
                        jkl = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_Q) Then
                    If jkl = False Then
                        lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, getgz(ligq.position.y, 2, True), ligq.position.z
                        Form3.List1_DblClick
                        jkl = True
                    End If
                ElseIf inp.IsKeyPressed(TV_KEY_E) Then
                    If jkl = False Then
                        lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, getgz(ligq.position.y, 2, False), ligq.position.z
                        Form3.List1_DblClick
                        jkl = True
                    End If
                Else
                End If
            Else
                jkl = False
                lig.GetLight GetLight(getstrrightb(str, "\")), ligq
                If inp.IsKeyPressed(TV_KEY_W) Then lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x + CSng(Text3.Text), ligq.position.y, ligq.position.z
                If inp.IsKeyPressed(TV_KEY_S) Then lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x - CSng(Text3.Text), ligq.position.y, ligq.position.z
                If inp.IsKeyPressed(TV_KEY_A) Then lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, ligq.position.y, ligq.position.z + CSng(Text3.Text)
                If inp.IsKeyPressed(TV_KEY_D) Then lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, ligq.position.y, ligq.position.z - CSng(Text3.Text)
                If inp.IsKeyPressed(TV_KEY_Q) Then lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, ligq.position.y + CSng(Text3.Text), ligq.position.z
                If inp.IsKeyPressed(TV_KEY_E) Then lig.SetLightPosition GetLight(getstrrightb(str, "\")), ligq.position.x, ligq.position.y - CSng(Text3.Text), ligq.position.z
                Form3.List1_DblClick
                Form3.findlist getstrrightb(str, "\")
            End If
        End If
    Else
        If aa Then
        hi = MsgBox("是否添加这个模型到材质列表？", 1, "提示")
        If hi = 1 Then
        stre = InputBox("请输入一个材质名，材质名必须正确", "提示", mat.GetMaterialName(1))
        List3.AddItem str + "|" + stre
        getmesh(str).SetMaterial (GetMat(stre))
        MsgBox "添加成功！", , "提示"
        End If
        str2 = str
        End If
    End If
 ElseIf col.GetCollisionObjectType = TV_OBJECT_ACTOR Then
    str = col.GetCollisionActor.GetName
    bq.Caption = "不能管理的角色:" + str
    If isinlist(List4, str) Then
        bq.Caption = "可以管理角色:" + str
        If aa = True Then
            bq.Caption = "管理角色中:" + str + "，可以改变其材质"
            str3 = str
            List4.ListIndex = inlist(List4, str)
        End If
    Else
        If aa Then
        hi = MsgBox("是否添加这个角色到材质列表？", 1, "提示")
        If hi = 1 Then
        stre = InputBox("请输入一个材质名，材质名必须正确", "提示", mat.GetMaterialName(1))
        List4.AddItem str + "|" + stre
        GetActor(str).SetMaterial (GetMat(stre))
        MsgBox "添加成功！", , "提示"
        End If
        str3 = str
        End If
    End If
 Else
 End If
 'If pz.GetCollisionObjectType = TV_OBJECT_MESH Then
 For vbb = 0 To lig.GetCount - 1
     lig.GetLight vbb + 1, ligt
     ge(vbb).SetPosition ligt.position.x, ligt.position.y, ligt.position.z
 Next
 For vbb = 0 To List5.ListCount - 1
     pa(vbb).SetPosition map.getpar(vbb).GetGlobalPosition.x, map.getpar(vbb).GetGlobalPosition.y, map.getpar(vbb).GetGlobalPosition.z
 Next
 Label2.Caption = "摄像机X:" & scene.GetCamera.GetPosition.x & Chr(13) & "摄像机Y:" & scene.GetCamera.GetPosition.y & Chr(13) & "摄像机Z:" & scene.GetCamera.GetPosition.z  'Chr(10) + Chr(13)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bDoLoop = False
    bDoLoop2 = False
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then Text5.Enabled = True Else Text5.Enabled = False
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then Text8.Enabled = True Else Text8.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
If zr = True Then If zr = True Then map.destroy
End
End Sub

Private Sub List2_DblClick()
Dim a As Long
Dim b() As String
b = Split(List2.List(List2.ListIndex), "|", , vbTextCompare)
a = GetMat(b(0))
R.Text = CStr(mat.GetAmbient(a).R)
G.Text = CStr(mat.GetAmbient(a).G)
Bbjk.Text = CStr(mat.GetAmbient(a).b)
TM.Text = CStr(mat.GetAmbient(a).a)
R2.Text = CStr(mat.GetDiffuse(a).R)
G2.Text = CStr(mat.GetDiffuse(a).G)
B2.Text = CStr(mat.GetDiffuse(a).b)
tm2.Text = CStr(mat.GetDiffuse(a).a)
r3.Text = CStr(mat.GetEmissive(a).R)
g3.Text = CStr(mat.GetEmissive(a).G)
b3.Text = CStr(mat.GetEmissive(a).b)
tm3.Text = CStr(mat.GetEmissive(a).a)
r4.Text = CStr(mat.GetSpecular(a).R)
g4.Text = CStr(mat.GetSpecular(a).G)
b4.Text = CStr(mat.GetSpecular(a).b)
tm4.Text = CStr(mat.GetSpecular(a).a)
gg.Text = CStr(mat.GetPower(a))
btm.Text = CStr(mat.GetOpacity(a))
prt.Value = IIf(b(1) = "1", 1, 0)
If b(1) = 1 Then
ssr.Text = b(2)
ssg.Text = b(3)
ssb.Text = b(4)
xsr.Text = b(5)
xsg.Text = b(6)
xsb.Text = b(7)
zs.Text = b(8)
Else
ssr.Text = "0"
ssg.Text = "0"
ssb.Text = "0"
xsr.Text = "0"
xsg.Text = "0"
xsb.Text = "0"
zs.Text = "0"
End If
End Sub
Private Sub List5_DblClick()
Dim a() As String
a = Split(List5.List(List5.ListIndex), "|", , vbTextCompare)
px.Text = CStr(map.getparformname(a(0)).GetGlobalPosition.x)
py.Text = CStr(map.getparformname(a(0)).GetGlobalPosition.y)
pz.Text = CStr(map.getparformname(a(0)).GetGlobalPosition.z)
pxx.Text = a(1)
pyy.Text = a(2)
pzz.Text = a(3)
pp.Text = a(4)
pxxx.Text = CStr(map.getparformname(a(0)).GetGlobalRotation.x)
pyyy.Text = CStr(map.getparformname(a(0)).GetGlobalRotation.y)
pzzz.Text = CStr(map.getparformname(a(0)).GetGlobalRotation.z)
End Sub
Private Sub List6_DblClick()
Command24_Click
End Sub
Private Sub Timer1_Timer()
Dim a As TVMesh
Dim y As TVMesh
Dim ggg() As Long
Dim ggg2() As Long
If bDoLoop2 Then
ReDim ge(lig.GetCount - 1)
For gh = 1 To lig.GetCount
Set a = gee.Duplicate("qq" & "\" & lig.GetLightName(gh))
Set ge(gh - 1) = a
Next
If List5.ListCount > 0 Then
ReDim pa(List5.ListCount - 1)
For gh = 1 To List5.ListCount
Set y = gef.Duplicate("bb" + "|" + getstrlefta(List5.List(gh - 1), "|"))
Set pa(gh - 1) = y
'Debug.Print ("bb" + "|" + getstrlefta(List5.List(gh - 1), "|"))
Next
End If
End If
End Sub
Function getgz(asng As Single, adir As Long, jia As Boolean) As Single 'xyz
Dim q As Single
If adir = 1 Then
    If jia = True Then
        q = (Int(asng / xx) + 1) * xx
    Else
        q = (Int(asng / xx)) * xx
    End If
ElseIf adir = 2 Then
    If jia = True Then
        q = (Int(asng / yy) + 1) * yy
    Else
        q = (Int(asng / yy)) * yy
    End If
ElseIf adir = 3 Then
    If jia = True Then
        q = (Int(asng / zz) + 1) * zz
    Else
        q = (Int(asng / zz)) * zz
    End If
Else
End If
getgz = q
End Function
Private Sub Timer2_Timer()
If prt.Value = 1 Then
ssr.Enabled = True
ssg.Enabled = True
ssb.Enabled = True
xsr.Enabled = True
xsg.Enabled = True
xsb.Enabled = True
zs.Enabled = True
Else
ssr.Enabled = False
ssg.Enabled = False
ssb.Enabled = False
xsr.Enabled = False
xsg.Enabled = False
xsb.Enabled = False
zs.Enabled = False
End If
End Sub
Private Sub Timer3_Timer()
'If Not List6.ListIndex = -1 Then
'    Select Case sous(getstrleftb(List6.List(List6.ListIndex), ".")).PlayState
'        Case TV_PLAYSTATE_UNDEFINED
'            yyxx.Caption = "未知"
'        Case TV_PLAYSTATE_PLAYING
'            yyxx.Caption = "播放中"
'        Case TV_PLAYSTATE_PAUSED
'            yyxx.Caption = "暂停中"
'        Case TV_PLAYSTATE_STOPPED
'            yyxx.Caption = "停止"
'        Case TV_PLAYSTATE_ENDED
'            yyxx.Caption = "结束"
'    End Select
'Else
'yyxx.Caption = "未选择"
'End If
End Sub
