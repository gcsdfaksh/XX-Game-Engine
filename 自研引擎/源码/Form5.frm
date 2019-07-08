VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "采样说明"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form5"
   ScaleHeight     =   4440
   ScaleWidth      =   5970
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   240
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "以上是一个能被音效引擎加载的音乐标准格式，当音频采样大小大于8时，音乐将不能被加载,MP3格式其实本来就不能加载的"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   3735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
