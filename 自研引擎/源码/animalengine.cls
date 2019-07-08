VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "animalengine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private ani(0 To 10) As animal
'Sub loadEX(index, posa() As TV_3DVECTOR, pica() As Long, timea() As Long, unitypos() As Long, Optional unitytime As Boolean)
'Dim a As Long
'If index > UBound(ani) Then ReDim Preserve ani(0 To index)
'ReDim ani(index).pic(0 To UBound(pic))
'ReDim ani(index).pos(0 To UBound(pic))
'ReDim ani(index).time(0 To UBound(pic))
'ani(index).nowfra = 0
'ani(index).pic = pica
'If unitypos(0) = 1 Then
'    ani(index).pos = posa
'Else
'    For a = 0 To UBound(pic)
'        ani(index).pos(a).x = unitypos(1)
'        ani(index).pos(a).y = unitypos(2)
'    Next
'End If
'If IsMissing(unitytime) Or unitytime = False Then
'    ani(index).time = timea
'Else
'    For a = 0 To UBound(pica)
'        ani(index).time(a) = timea(0)
'    Next
'End If
'End Sub
'Sub load(index, posa() As TV_3DVECTOR, pica() As Long, timea() As Long, unitypos() As Long, Optional unitytime As Boolean, Optional alphaa As Long)
'Dim a As Long
'If index > UBound(ani) Then ReDim Preserve ani(0 To index)
'ReDim ani(index).pic(0 To UBound(pic))
'ReDim ani(index).pos(0 To UBound(pic))
'ReDim ani(index).time(0 To UBound(pic))
'ani(index).nowfra = 0
'For a = 0 To UBound(pica)
'    If IsMissing(alpha) Then
'        ani(index).pic(a) = tex.LoadTexture(pic(a), pic(a), , , alphaa)
'    Else
'        ani(index).pic(a) = tex.LoadTexture(pic(a), pic(a), , , alpha)
'    End If
'Next
'If unitypos(0) = 1 Then
'    ani(index).pos = posa
'Else
'    For a = 0 To UBound(pic)
'        ani(index).pos(a).x = unitypos(1)
'        ani(index).pos(a).y = unitypos(2)
'    Next
'End If
'If IsMissing(unitytime) Or unitytime = False Then
'    ani(index).time = timea
'Else
'    For a = 0 To UBound(pica)
'        ani(index).time(a) = timea(0)
'    Next
'End If
'End Sub
'Sub copy(oldindex As Long, newindex As Long, Optional copyfraandtime As Boolean) '为假或空时不复制时间和帧
'ReDim ani(newindex).pic(0 To UBound(ani(oldindex).pic))
'ani(newindex).pic = ani(oldindex).pic
'ReDim ani(newindex).pos(0 To UBound(ani(oldindex).pos))
'ani(newindex).pos = ani(oldindex).pos
'ReDim ani(newindex).time(0 To UBound(ani(oldindex).time))
'ani(newindex).time = ani(oldindex).time
'If IsMissing(copyfra) Or copyfra = False Then
'Else
'    ani(newindex).nowfra = ani(oldindex).nowfra
'    ani(newindex).nowtime = ani(oldindex).nowtime
'End If
'End Sub
'Sub show(index As Long, left As Single, top As Single, width As Single, height As Single, showtype) '0显示后停止1显示后销毁2循环
'If ani(index).nowtime > ani(index).time(ani(index).nowfra) Then
'    If ani(index).nowfra = UBound(ani(index).time) Then
'        If showtype = 1 Then
'            ani(index).des = True
'        ElseIf showtype = 2 Then
'            ani(index).nowfra = 0
'        Else
'        End If
'    Else
'        ani(index).nowfra = ani(index).nowfra + 1
'    End If
'    ani(index).nowtime = 0
'Else
'    ani(index).nowtime = ani(index).nowtime + 1
'End If
'If ani(index).des = False Then
'    scr2D.Draw_Texture ani(index).pic(ani(index).nowfra), left + ani(index).pos(ani(index).nowfra).x, top + ani(index).pos(ani(index).nowfra).y, left + ani(index).pos(ani(index).nowfra).x + width, top + ani(index).pos(ani(index).nowfra).y + height
'End Sub
'Sub reset(index As Long, Optional clear As Boolean) '假或空重置(重新可以播放动画)，否则为清空
'Dim a As Long
'If clear = False Or IsMissing(clear) = True Then
'    ani(index).nowfra = 0
'    ani(index).nowtime = 0
'    ani(index).des = False
'Else
'    For a = 0 To UBound(ani(index).pic)
'        ani(index).pos(a).x = 0
'        ani(index).pos(a).y = 0
'        ani(index).pic(a) = 0
'        ani(index).time(a) = 10
'    Next
'    ani(index).nowfra = 0
'    ani(index).nowtime = 0
'    ani(index).des = False
'End If
'End Sub
'
'
