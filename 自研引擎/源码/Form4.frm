VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "纹理文件夹"
   ClientHeight    =   2925
   ClientLeft      =   285
   ClientTop       =   7950
   ClientWidth     =   4725
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4725
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2160
      TabIndex        =   9
      Text            =   "0"
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Alpha"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option5 
      Caption         =   "DUVU"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Caption         =   "立方"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "凹凸"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "3D"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "普通"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1890
      Left            =   240
      OLEDragMode     =   1  'Automatic
      System          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除纹理"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加纹理"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
  
   Begin VB.Label Label1 
      Caption         =   "透明色："
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim b As String
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNFileMustExist
CommonDialog1.Filter = "BMP位图(*.bmp)|*.bmp|JPG图片(*.jpg)|*.jpg|PNG图片(*.png)|*.png|GIF图片(*.gif)|*.gif|DDS立方图(*.dds)|*.dds|所有文件(*.*)|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileTitle = "" Then Exit Sub
b = CommonDialog1.FileName
FileCopy b, getstrleftb(Form1.Text1.Text, ".") + "\texture\" + getstrrightb(b, "\")
If Option1.Value Then
    Form1.List1.AddItem getstrrightb(b, "\") + "\" + "普通" + "\" + Text1.Text
    tex.LoadTexture getstrrightb(b, "\"), getstrrightb(b, "\"), CSng(Text1.Text), True
ElseIf Option2.Value Then
    Form1.List1.AddItem getstrrightb(b, "\") + "\" + "3D" + "\" + Text1.Text
    tex.LoadVolumeTexture getstrrightb(b, "\"), getstrrightb(b, "\"), CSng(Text1.Text), True
ElseIf Option3.Value Then
    Form1.List1.AddItem getstrrightb(b, "\") + "\" + "凹凸" + "\" + Text1.Text
    tex.LoadBumpTexture getstrrightb(b, "\"), getstrrightb(b, "\"), , , True
ElseIf Option4.Value Then
    Form1.List1.AddItem getstrrightb(b, "\") + "\" + "立方" + "\" + Text1.Text
    tex.LoadCubeTexture getstrrightb(b, "\"), getstrrightb(b, "\"), , CSng(Text1.Text), True
ElseIf Option5.Value Then
    Form1.List1.AddItem getstrrightb(b, "\") + "\" + "DUVU" + "\" + Text1.Text
    tex.LoadDUDVTexture getstrrightb(b, "\"), getstrrightb(b, "\")
ElseIf Option6.Value Then
    Form1.List1.AddItem getstrrightb(b, "\") + "\" + "Alpha" + "\" + Text1.Text
    tex.LoadAlphaTexture getstrrightb(b, "\"), getstrrightb(b, "\")
Else
End If
Unload Form4
End Sub
Private Sub Command2_Click()
If File1.ListIndex = -1 Then
MsgBox "请选择一个要删除的纹理！"
Else
Kill getstrleftb(Form1.Text1.Text, ".") + "\texture\" + File1.List(File1.ListIndex)
Form1.List1.RemoveItem File1.ListIndex
tex.DeleteTexture GetTex(File1.List(File1.ListIndex))
Unload Form4
End If
End Sub

