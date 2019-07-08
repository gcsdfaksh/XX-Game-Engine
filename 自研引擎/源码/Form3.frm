VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "GetLightProperties"
   ClientHeight    =   8850
   ClientLeft      =   240
   ClientTop       =   0
   ClientWidth     =   5790
   FillStyle       =   0  'Solid
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command22 
      Caption         =   "校准"
      Height          =   255
      Left            =   1800
      TabIndex        =   81
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   6000
   End
   Begin VB.CheckBox Check6 
      Caption         =   "属性"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   4560
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox r3 
      Height          =   270
      Left            =   2040
      TabIndex        =   69
      Text            =   "0"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox g3 
      Height          =   270
      Left            =   2040
      TabIndex        =   68
      Text            =   "0"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox b3 
      Height          =   270
      Left            =   2040
      TabIndex        =   67
      Text            =   "0"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox r4 
      Height          =   270
      Left            =   3600
      TabIndex        =   66
      Text            =   "0"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox g4 
      Height          =   270
      Left            =   3600
      TabIndex        =   65
      Text            =   "0"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox b4 
      Height          =   270
      Left            =   3600
      TabIndex        =   64
      Text            =   "0"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox g2 
      Height          =   270
      Left            =   600
      TabIndex        =   60
      Text            =   "0"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox b2 
      Height          =   270
      Left            =   600
      TabIndex        =   59
      Text            =   "0"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox r2 
      Height          =   270
      Left            =   600
      TabIndex        =   58
      Text            =   "0"
      Top             =   3480
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form3.frx":0000
      Left            =   0
      List            =   "Form3.frx":0002
      TabIndex        =   54
      Top             =   6360
      Width           =   5175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Go here"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   53
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   3360
      TabIndex        =   51
      Text            =   "512"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3360
      TabIndex        =   49
      Text            =   "0"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "阴影常量"
      Height          =   255
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4920
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "立体纹理"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   46
      Text            =   "纹理1"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CheckBox gyx 
      Caption         =   "光映像"
      Height          =   255
      Left            =   3240
      TabIndex        =   45
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox yz 
      Caption         =   "投影"
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox dg 
      Height          =   1095
      Left            =   4800
      TabIndex        =   43
      Text            =   "灯光"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox qd 
      Height          =   270
      Left            =   3600
      TabIndex        =   40
      Text            =   "0"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox tg 
      Caption         =   "托管"
      Height          =   255
      Left            =   1440
      TabIndex        =   39
      Top             =   4560
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.OptionButton zz 
      Caption         =   "直线光"
      Height          =   375
      Left            =   4800
      TabIndex        =   38
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton jj 
      Caption         =   "聚光灯"
      Height          =   375
      Left            =   4800
      TabIndex        =   37
      Top             =   720
      Width           =   975
   End
   Begin VB.OptionButton dd 
      Caption         =   "点光源"
      Height          =   375
      Left            =   4800
      TabIndex        =   36
      Top             =   240
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox b 
      Height          =   270
      Left            =   3600
      TabIndex        =   35
      Text            =   "0"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox r 
      Height          =   270
      Left            =   3600
      TabIndex        =   32
      Text            =   "0"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   30
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "创建"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox sj3 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      TabIndex        =   26
      Text            =   "0"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox sj2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      TabIndex        =   25
      Text            =   "0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox sj1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      TabIndex        =   22
      Text            =   "0"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox fxz 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      TabIndex        =   17
      Text            =   "0"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox fxy 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      TabIndex        =   16
      Text            =   "0"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox fxx 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      TabIndex        =   15
      Text            =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox bj 
      Height          =   270
      Left            =   3600
      TabIndex        =   10
      Text            =   "100"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox sj 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3600
      TabIndex        =   9
      Text            =   "0"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox theta 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3600
      TabIndex        =   8
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox g 
      Height          =   270
      Left            =   3600
      TabIndex        =   7
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox phi 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3600
      TabIndex        =   6
      Text            =   "0"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox y 
      Height          =   270
      Left            =   480
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox z 
      Height          =   270
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox x 
      Height          =   270
      Left            =   480
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label40 
      Caption         =   "范围"
      Height          =   735
      Left            =   4440
      TabIndex        =   80
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label39 
      Caption         =   "b"
      Height          =   255
      Left            =   2280
      TabIndex        =   79
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label38 
      Caption         =   "g"
      Height          =   255
      Left            =   1680
      TabIndex        =   78
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label35 
      Caption         =   "r"
      Height          =   255
      Left            =   1080
      TabIndex        =   77
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label37 
      Caption         =   "阴影常量"
      Height          =   375
      Left            =   3600
      TabIndex        =   76
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label36 
      Caption         =   "立体纹理"
      Height          =   495
      Left            =   2880
      TabIndex        =   75
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label31 
      Caption         =   "名称"
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label30 
      Caption         =   "漫射b"
      Height          =   255
      Left            =   3000
      TabIndex        =   72
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label29 
      Caption         =   "漫射g"
      Height          =   255
      Left            =   3000
      TabIndex        =   71
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label28 
      Caption         =   "漫射r"
      Height          =   255
      Left            =   3000
      TabIndex        =   70
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label27 
      Caption         =   "反射b"
      Height          =   375
      Left            =   1560
      TabIndex        =   63
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label26 
      Caption         =   "反射g"
      Height          =   255
      Left            =   1560
      TabIndex        =   62
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label25 
      Caption         =   "反射r"
      Height          =   255
      Left            =   1560
      TabIndex        =   61
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label24 
      Caption         =   "b"
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label23 
      Caption         =   "g"
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label22 
      Caption         =   "r"
      Height          =   375
      Left            =   240
      TabIndex        =   55
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "范围"
      Height          =   255
      Left            =   2880
      TabIndex        =   52
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "无0 摄影全球1 全球阴影图2 摄影对象3 强度不可更改 请双击列表进入光 还有衰减获取不了 环境色有缺陷"
      Height          =   2055
      Left            =   4680
      TabIndex        =   50
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "名称"
      Height          =   375
      Left            =   4800
      TabIndex        =   42
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "强度"
      Height          =   255
      Left            =   3120
      TabIndex        =   41
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "环境B"
      Height          =   375
      Left            =   3000
      TabIndex        =   34
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "环境G"
      Height          =   375
      Left            =   3000
      TabIndex        =   33
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "托管0为假1为真 类型有1点光源2聚光灯3直射光"
      Height          =   975
      Left            =   1920
      TabIndex        =   31
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "环境R"
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "衰减3"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "衰减2"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "衰减1"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "theta"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "衰减"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "半径"
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "方向Z："
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "方向Y："
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "方向X："
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "phi"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Z："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Y："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub loadli()
Dim a As Long
Dim b() As String
Dim b2() As String
If lig.GetCount > 0 Then
a = pac.OpenPackage(Form1.Text1.Text)
pac.SetArchivePassword "fsbzwy", a
b = Split(pacstr("light.dat", a), Chr(13) + Chr(10), , vbTextCompare)

For q = 1 To lig.GetCount - 1
b2 = Split(pacstr("light\" + b(q - 1) + ".lit", a), Chr(13) + Chr(10), , vbTextCompare)
ReDim Preserve b2(0 To UBound(b2) + 20)
For zzz = 0 To UBound(b2)
If b2(zzz) = "" Then b2(zzz) = "无"
Next
Select Case b2(0)
Case "1"
If b2(15) = "无" Then b2(15) = "0"
If b2(16) = "无" Then b2(16) = "0"
If b2(17) = "无" Then b2(17) = "0"
List1.List(q - 1) = lig.GetLightName(q) + "|" + b2(15) + "|" + b2(16) + "|" + b2(17) + "|" + b2(12) + "|" + b2(13) + "|" + b2(14)
Case "2"
If b2(25) = "无" Then b2(25) = "0"
If b2(26) = "无" Then b2(26) = "0"
If b2(27) = "无" Then b2(27) = "0"
List1.List(q - 1) = lig.GetLightName(q) + "|" + b2(25) + "|" + b2(26) + "|" + b2(27) + "|" + b2(22) + "|" + b2(23) + "|" + b2(24)
Case "3"
If b2(14) = "无" Then b2(14) = "0"
If b2(15) = "无" Then b2(15) = "0"
If b2(16) = "无" Then b2(16) = "0"
List1.List(q - 1) = lig.GetLightName(q) + "|" + b2(14) + "|" + b2(15) + "|" + b2(16) + "|" + b2(11) + "|" + b2(12) + "|" + b2(13)
Case Else
End Select
Next
List1.ListIndex = List1.ListCount - 1
List1_DblClick
pac.ClosePackage a
End If
End Sub
Private Sub Command1_Click()
Dim a As Long
Dim v As Boolean
Dim q() As String
If List1.ListIndex <> -1 Then q = Split(List1.List(List1.ListIndex), "|", , vbTextCompare): If dg.Text = q(0) Then MsgBox ("不能创建一个同名的灯光！！！"): v = True
If v = False Then
If dd.Value = True Then
    a = lig.CreatePointLight(cre3d(CSng(x.Text), CSng(y.Text), CSng(z.Text)), CSng(r.Text), CSng(g.Text), CSng(b.Text), CSng(bj.Text), dg.Text, qd.Text)
    'RemoveItem
    aaa a, True
    'lig.SetLight a
ElseIf jj.Value = True Then
    a = lig.CreateSpotLight(cre3d(CSng(x.Text), CSng(y.Text), CSng(z.Text)), cre3d(CSng(fxx.Text), CSng(fxy.Text), CSng(fxz.Text)), CSng(r.Text), CSng(g.Text), CSng(b.Text), CSng(bj.Text), CSng(phi.Text), CSng(theta.Text), dg.Text, qd.Text)
    aaa a, True
ElseIf zz.Value = True Then
    a = lig.CreateDirectionalLight(cre3d(CSng(fxx.Text), CSng(fxy.Text), CSng(fxz.Text)), CSng(r.Text), CSng(g.Text), CSng(b.Text), dg.Text, qd.Text)
    aaa a, True
Else
End If
End If
End Sub
Sub findlist(str As String)
Dim wsx As Long
Dim q() As String
For wsx = 0 To List1.ListCount - 1
If List1.ListIndex <> -1 Then
    q = Split(List1.List(wsx), "|", , vbTextCompare)
End If
If q(0) = str Then List1.ListIndex = wsx
Next
End Sub
Sub aaa(a As Long, Optional 创建表格 As Boolean)
If IsMissing(创建表格) Or 创建表格 = False Then
List1.List(List1.ListIndex) = dg.Text & "|" & r2.Text & "|" & g2.Text & "|" & b2.Text & "|" & IIf(Check1.Value = 1, Text1.Text, "无") & "|" & IIf(Check2.Value = 1, Text2.Text & "|" & Text3.Text, "无|无")
If dd.Value = True Then
    lig.SetLightPosition a, CSng(x.Text), CSng(y.Text), CSng(z.Text)
    lig.SetLightRange a, CSng(bj.Text)
    lig.SetLightAmbientColor a, CSng(r.Text), CSng(g.Text), CSng(b.Text)
ElseIf jj.Value = True Then
    lig.SetLightPosition a, CSng(x.Text), CSng(y.Text), CSng(z.Text)
    lig.SetLightRange a, CSng(bj.Text)
    lig.SetLightAmbientColor a, CSng(r.Text), CSng(g.Text), CSng(b.Text)
    lig.SetLightDirection a, CSng(fxx.Text), CSng(fxy.Text), CSng(fxz.Text)
    lig.SetLightSpotFalloff a, CSng(sj.Text)
    lig.SetLightAttenuation a, CSng(sj1.Text), CSng(sj2.Text), CSng(sj3.Text)
    lig.SetLightSpotAngles a, CSng(phi.Text), CSng(theta.Text)
ElseIf jj.Value = True Then
    lig.SetLightAmbientColor a, CSng(r.Text), CSng(g.Text), CSng(b.Text)
    lig.SetLightDirection a, CSng(fxx.Text), CSng(fxy.Text), CSng(fxz.Text)
Else
End If
lig.SetLightColor a, CSng(r2.Text), CSng(b2.Text), CSng(g2.Text)
lig.SetLightSpecularColor a, CSng(r3.Text), CSng(g3.Text), CSng(b3.Text)
lig.SetLightDiffuseColor a, CSng(r4.Text), CSng(g4.Text), CSng(b4.Text)
If Check6.Value = 1 Then lig.SetLightProperties a, IIf(tg.Value = 1, True, False), IIf(yz.Value = 1, True, False), IIf(gyx.Value = 1, True, False)
If Check1.Value = 1 Then lig.SetLightCubeMap a, GetTex(Text1.Text)
If Check2.Value = 1 Then lig.SetProjectiveShadowsProperties a, CLng(Text2.Text), CLng(Text3.Text)
Else
By = List1.ListCount
List1.AddItem dg.Text & "|" & r2.Text & "|" & g2.Text & "|" & b2.Text & "|" & IIf(Check1.Value = 1, Text1.Text, "无") & "|" & IIf(Check2.Value = 1, Text2.Text & "|" & Text3.Text, "无|无")
List1.ListIndex = By
lig.SetLightColor a, CSng(r2.Text), CSng(b2.Text), CSng(g2.Text)
lig.SetLightSpecularColor a, CSng(r3.Text), CSng(g3.Text), CSng(b3.Text)
lig.SetLightDiffuseColor a, CSng(r4.Text), CSng(g4.Text), CSng(b4.Text)
If Check6.Value = 1 Then lig.SetLightProperties a, IIf(tg.Value = 1, True, False), IIf(yz.Value = 1, True, False), IIf(gyx.Value = 1, True, False)
If Check1.Value = 1 Then lig.SetLightCubeMap a, GetTex(Text1.Text)
If Check2.Value = 1 Then lig.SetProjectiveShadowsProperties a, CLng(Text2.Text), CLng(Text3.Text)
End If
End Sub
Private Sub Command2_Click()
Dim q() As String
If List1.ListIndex <> -1 Then
    q = Split(List1.List(List1.ListIndex), "|", , vbTextCompare)
    lig.DeleteLight GetLight(q(0))
    List1.RemoveItem List1.ListIndex
Else
MsgBox "请选择一个光！"
End If
List1.ListIndex = 0
List1_DblClick
End Sub
Private Sub Command22_Click()
x.Text = scene.GetCamera.GetPosition.x
y.Text = scene.GetCamera.GetPosition.y
z.Text = scene.GetCamera.GetPosition.z
End Sub
Private Sub Command3_Click()
Dim q() As String
If List1.ListIndex <> -1 Then q = Split(List1.List(List1.ListIndex), "|", , vbTextCompare) Else MsgBox "请选择一个光！": Exit Sub
If dd.Value = True Then
    'RemoveItem
    aaa GetLight(q(0))
    'lig.SetLight a
ElseIf jj.Value = True Then
    aaa GetLight(q(0))
ElseIf zz.Value = True Then
    aaa GetLight(q(0))
Else
End If
List1_DblClick
End Sub
Private Sub Command4_Click()
If List1.ListIndex = -1 Then MsgBox "请选择一个光！": Exit Sub
scene.GetCamera.SetCamera CSng(x.Text), CSng(y.Text), CSng(z.Text), CSng(x.Text), CSng(y.Text), CSng(z.Text)
'ge.SetPosition CSng(x.Text), CSng(y.Text), CSng(z.Text)
kg = True
End Sub
Private Sub dd_Click()
fxx.Enabled = False
fxy.Enabled = False
fxz.Enabled = False
sj1.Enabled = False
sj2.Enabled = False
sj3.Enabled = False
phi.Enabled = False
theta.Enabled = False
sj.Enabled = False
x.Enabled = True
y.Enabled = True
z.Enabled = True
bj.Enabled = True
End Sub
Private Sub jj_Click()
fxx.Enabled = True
fxy.Enabled = True
fxz.Enabled = True
sj1.Enabled = True
sj2.Enabled = True
sj3.Enabled = True
phi.Enabled = True
theta.Enabled = True
sj.Enabled = True
x.Enabled = True
y.Enabled = True
z.Enabled = True
bj.Enabled = True
End Sub
Public Sub List1_DblClick()
Dim a As TV_LIGHT
Dim q() As String
Dim aa As Boolean
Dim aa2 As Boolean
Dim aa3 As Boolean
If List1.ListIndex <> -1 Then
    q = Split(List1.List(List1.ListIndex), "|", , vbTextCompare)
    dg.Text = q(0)
    lig.GetLight GetLight(dg.Text), a
    Select Case a.Type
    Case TV_LIGHT_POINT
        dd.Value = True
        x.Text = CStr(a.position.x)
        y.Text = CStr(a.position.y)
        z.Text = CStr(a.position.z)
        bj.Text = CStr(a.range)
        r.Text = CStr(a.Ambient.r)
        g.Text = CStr(a.Ambient.g)
        b.Text = CStr(a.Ambient.b)
        r3.Text = CStr(a.specular.r)
        g3.Text = CStr(a.specular.g)
        b3.Text = CStr(a.specular.b)
        r4.Text = CStr(a.diffuse.r)
        g4.Text = CStr(a.diffuse.g)
        b4.Text = CStr(a.diffuse.b)
        If q(4) <> "无" Then Text1.Text = q(4): Check1.Value = 1 Else Check1.Value = 0
        If q(5) <> "无" Then Text2.Text = q(5): Text3.Text = q(6): Check2.Value = 1 Else Check2.Value = 0
        r2.Text = q(1)
        g2.Text = q(2)
        b2.Text = q(3)
        lig.GetLightProperties GetLight(q(0)), aa, aa2, aa3
        If aa = True Then tg.Value = 1 Else tg.Value = 0
        If aa2 = True Then yz.Value = 1 Else yz.Value = 0
        If aa3 = True Then gyx.Value = 1 Else gyx.Value = 0
    Case TV_LIGHT_SPOT
        jj.Value = True
        x.Text = CStr(a.position.x)
        y.Text = CStr(a.position.y)
        z.Text = CStr(a.position.z)
        fxx.Text = CStr(a.direction.x)
        fxy.Text = CStr(a.direction.y)
        fxz.Text = CStr(a.direction.z)
        bj.Text = CStr(a.range)
        r.Text = CStr(a.Ambient.r)
        g.Text = CStr(a.Ambient.g)
        b.Text = CStr(a.Ambient.b)
        r3.Text = CStr(a.specular.r)
        g3.Text = CStr(a.specular.g)
        b3.Text = CStr(a.specular.b)
        r4.Text = CStr(a.diffuse.r)
        g4.Text = CStr(a.diffuse.g)
        b4.Text = CStr(a.diffuse.b)
        sj.Text = CStr(a.fFallOff)
        sj1.Text = CStr(a.attenuation.x)
        sj2.Text = CStr(a.attenuation.y)
        sj3.Text = CStr(a.attenuation.z)
        phi.Text = CStr(a.phi)
        theta.Text = CStr(a.theta)
        If q(4) <> "无" Then Text1.Text = q(4): Check1.Value = 1 Else Check1.Value = 0
        If q(5) <> "无" Then Text2.Text = q(5): Text3.Text = q(6): Check2.Value = 1 Else Check2.Value = 0
        r2.Text = q(1)
        g2.Text = q(2)
        b2.Text = q(3)
        lig.GetLightProperties GetLight(q(0)), aa, aa2, aa3
        If aa = True Then tg.Value = 1 Else tg.Value = 0
        If aa2 = True Then yz.Value = 1 Else yz.Value = 0
        If aa3 = True Then gyx.Value = 1 Else gyx.Value = 0
    Case TV_LIGHT_DIRECTIONAL
        zz.Value = True
        fxx.Text = CStr(a.direction.x)
        fxy.Text = CStr(a.direction.y)
        fxz.Text = CStr(a.direction.z)
        r.Text = CStr(a.Ambient.r)
        g.Text = CStr(a.Ambient.g)
        b.Text = CStr(a.Ambient.b)
        r3.Text = CStr(a.specular.r)
        g3.Text = CStr(a.specular.g)
        b3.Text = CStr(a.specular.b)
        r4.Text = CStr(a.diffuse.r)
        g4.Text = CStr(a.diffuse.g)
        b4.Text = CStr(a.diffuse.b)
        If q(4) <> "无" Then Text1.Text = q(4): Check1.Value = 1 Else Check1.Value = 0
        If q(5) <> "无" Then Text2.Text = q(5): Text3.Text = q(6): Check2.Value = 1 Else Check2.Value = 0
        r2.Text = q(1)
        g2.Text = q(2)
        b2.Text = q(3)
        lig.GetLightProperties GetLight(q(0)), aa, aa2, aa3
        If aa = True Then tg.Value = 1 Else tg.Value = 0
        If aa2 = True Then yz.Value = 1 Else yz.Value = 0
        If aa3 = True Then gyx.Value = 1 Else gyx.Value = 0
    Case Else
    End Select
End If
End Sub
Private Sub Timer1_Timer()
If List1.List(0) = "" And List1.ListCount = 2 Then List1.RemoveItem (0)
End Sub
Sub savelig(qq As String)
'Dim fs As Scripting
Dim a As TextStream
Dim bb As TextStream
Dim q() As String
If List1.ListCount <> 0 Then
Set fs = CreateObject("Scripting.FileSystemObject")
Set bb = fs.CreateTextFile(qq + "light.dat", True)
For qw = 1 To List1.ListCount
If qw = List1.ListCount Then
q = Split(List1.List(qw - 1), "|", , vbTextCompare)
bb.Write (q(0))
Else
q = Split(List1.List(qw - 1), "|", , vbTextCompare)
bb.WriteLine (q(0))
End If
Next
bb.Close
For qw = 1 To List1.ListCount
List1.ListIndex = qw - 1
q = Split(List1.List(qw - 1), "|", , vbTextCompare)
Set a = fs.CreateTextFile(qq + "light\" + q(0) + ".lit", True)
List1_DblClick
If dd.Value = True Then
    a.WriteLine ("1")
    a.WriteLine (x.Text)
    a.WriteLine (y.Text)
    a.WriteLine (z.Text)
    a.WriteLine (r.Text)
    a.WriteLine (g.Text)
    a.WriteLine (b.Text)
    a.WriteLine (bj.Text)
    a.WriteLine (qd.Text)
    a.WriteLine (CStr(tg.Value))
    a.WriteLine (CStr(yz.Value))
    a.WriteLine (CStr(gyx.Value))
    If Check1.Value = 1 Then a.WriteLine (Text1.Text) Else a.WriteLine ("")
    If Check2.Value = 1 Then a.WriteLine (Text2.Text): a.WriteLine (Text3.Text) Else a.WriteLine (""): a.WriteLine ("")
    a.WriteLine (r2.Text)
    a.WriteLine (g2.Text)
    a.WriteLine (b2.Text)
    a.WriteLine (r3.Text)
    a.WriteLine (g3.Text)
    a.WriteLine (b3.Text)
    a.WriteLine (r4.Text)
    a.WriteLine (g4.Text)
    a.Write (b4.Text)
ElseIf jj.Value = True Then
    a.WriteLine ("2")
    a.WriteLine (x.Text)
    a.WriteLine (y.Text)
    a.WriteLine (z.Text)
    a.WriteLine (fxx.Text)
    a.WriteLine (fxy.Text)
    a.WriteLine (fxz.Text)
    a.WriteLine (r.Text)
    a.WriteLine (g.Text)
    a.WriteLine (b.Text)
    a.WriteLine (bj.Text)
    a.WriteLine (phi.Text)
    a.WriteLine (theta.Text)
    a.WriteLine (qd.Text)
    a.WriteLine (CStr(tg.Value))
    a.WriteLine (CStr(yz.Value))
    a.WriteLine (CStr(gyx.Value))
    a.WriteLine (sj1.Text)
    a.WriteLine (sj2.Text)
    a.WriteLine (sj3.Text)
    a.WriteLine (sj.Text)
    If Check1.Value = 1 Then a.WriteLine (Text1.Text) Else a.WriteLine ("")
    If Check2.Value = 1 Then a.WriteLine (Text2.Text): a.WriteLine (Text3.Text) Else a.WriteLine (""): a.WriteLine ("")
    a.WriteLine (r2.Text)
    a.WriteLine (g2.Text)
    a.WriteLine (b2.Text)
    a.WriteLine (r3.Text)
    a.WriteLine (g3.Text)
    a.WriteLine (b3.Text)
    a.WriteLine (r4.Text)
    a.WriteLine (g4.Text)
    a.Write (b4.Text)
ElseIf zz.Value = True Then
    a.WriteLine ("3")
    a.WriteLine (fxx.Text)
    a.WriteLine (fxy.Text)
    a.WriteLine (fxz.Text)
    a.WriteLine (r.Text)
    a.WriteLine (g.Text)
    a.WriteLine (b.Text)
    a.WriteLine (qd.Text)
    a.WriteLine (CStr(tg.Value))
    a.WriteLine (CStr(yz.Value))
    a.WriteLine (CStr(gyx.Value))
    If Check1.Value = 1 Then a.WriteLine (Text1.Text) Else a.WriteLine ("")
    If Check2.Value = 1 Then a.WriteLine (Text2.Text): a.WriteLine (Text3.Text) Else a.WriteLine (""): a.WriteLine ("")
    a.WriteLine (r2.Text)
    a.WriteLine (g2.Text)
    a.WriteLine (b2.Text)
    a.WriteLine (r3.Text)
    a.WriteLine (g3.Text)
    a.WriteLine (b3.Text)
    a.WriteLine (r4.Text)
    a.WriteLine (g4.Text)
    a.Write (b4.Text)
Else
End If
a.Close
Next
End If
End Sub
Private Sub zz_Click()
fxx.Enabled = True
fxy.Enabled = True
fxz.Enabled = True
sj1.Enabled = False
sj2.Enabled = False
sj3.Enabled = False
phi.Enabled = False
theta.Enabled = False
sj.Enabled = False
x.Enabled = False
y.Enabled = False
z.Enabled = False
bj.Enabled = False
End Sub
