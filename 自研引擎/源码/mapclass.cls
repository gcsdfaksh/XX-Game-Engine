VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mapclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private npc() As TVActor
Private mesh() As TVMesh
Private land As TVLandscape
Private sky As TVAtmosphere
Private meshcount As Long
Private npccount As Long
Private texa() As Long
Private par() As TVParticleSystem
Private liga() As Long
Private mata() As Long
Private soulist() As String
Private othlist() As String
Private meshload As Boolean
Private npcload As Boolean
Private iniload As Boolean
Private parload As Boolean
Private souload As Boolean
Private texload As Boolean
Private matload As Boolean
Private lightload As Boolean
Private landload As Boolean
Private othload As Boolean
Private jbload As Boolean
Function isload() As Boolean
isload = iniload
End Function
Function loadmap(scene As TVScene, file As String, password As String, Optional plan As String) As Long '0未找到，1加载中，2加载成功，4未卸载
Dim index As Long
Dim meshdat() As String
Dim meshdat2() As String
Dim npcdat() As String
Dim npcdat2() As String
Dim landdat() As String
Dim matdat() As String
Dim meshmat() As String
Dim meshmat2() As String
Dim npcmat() As String
Dim npcmat2() As String
Dim textdat() As String
Dim textdat2() As String
Dim pardat() As String
Dim pardat2() As String
Dim a As Long
Dim aa() As String
Dim d As Long
Dim soudat() As String
Dim fil As Folder
Dim fill As Folder
Dim othdat() As String
Dim i As Long
Dim s As String
Static load As Long
'///
Set fi = CreateObject("Scripting.FileSystemObject")

If load = 0 Then
If iniload = True Then
If Not IsMissing(plan) Then plan = "地图未卸载！"
load = 0
loadmap = 4
Exit Function
Else
If fi.FileExists(file) Then
If Not IsMissing(plan) Then plan = "打开地图......"
If load = 0 Then load = 1: loadmap = 1: Exit Function
Else
If Not IsMissing(plan) Then plan = "未找到地图！"
iniload = False
load = 0
loadmap = 0
Exit Function
End If
End If
End If

If load = 1 Then
index = pac.OpenPackage(file)
pac.SetArchivePassword password, index
If pac.FileExists("texture.dat", index) Then
If Not IsMissing(plan) Then plan = "读取纹理......"
End If
load = 2: loadmap = 1: Exit Function
End If

If load = 2 Then
If pac.FileExists("mat.dat", index) Or pac.FileExists("light.dat", index) Then
If Not IsMissing(plan) Then plan = "读取材质与灯光中......"
End If
If pac.FileExists("texture.dat", index) Then
textdat = Split(pacstr("texture.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim texa(0 To UBound(textdat))
For a = 0 To UBound(textdat)
    If InStrRev(textdat(a), "|") <> 0 Then
        textdat2 = Split(textdat(a), "|", , vbTextCompare)
        If UBound(textdat2) = 0 Then
            texa(a) = loadtex("texture\" + textdat2(0), index)
        ElseIf UBound(textdat2) = 1 Then
            texa(a) = loadtex("texture\" + textdat2(0), index, CLng(textdat2(1)))
        ElseIf UBound(textdat2) = 2 Then
            texa(a) = loadtex("texture\" + textdat2(0), index, CLng(textdat2(1)), CLng(textdat2(2)))
        Else
            texa(a) = loadtex("texture\" + textdat2(0), index)
        End If
    Else
        texa(a) = loadtex("texture\" + textdat(a), index)
    End If
Next
texload = True
Else
texload = False
End If
End If


If load = 2 Then load = 3: loadmap = 1: Exit Function
If load = 3 Then
If pac.FileExists("land.ini", index) Then
If Not IsMissing(plan) Then plan = "创建地形与天空中......"
End If
If pac.FileExists("mat.dat", index) Then
matdat = Split(pacstr("mat.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim mata(0 To UBound(matdat))
For a = 0 To UBound(matdat)
    mata(a) = loadmat("Mat\" + matdat(a) + ".mat", index)
Next
matload = True
Else
matload = False
End If
End If

If load = 3 Then load = 4: loadmap = 1: Exit Function
If load = 4 Then
If pac.FileExists("light.dat", index) Then
aa = Split(pacstr("light.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim liga(0 To UBound(aa))
For d = 0 To UBound(aa)
    liga(d) = loadlight("light\" + aa(d) + ".lit", index)
Next
lightload = True
Else
lightload = False
End If
End If

If load = 4 Then load = 5: loadmap = 1: Exit Function
If load = 5 Then
If pac.FileExists("Mesh.dat", index) Then
If Not IsMissing(plan) Then plan = "载入模型......"
End If
If pac.FileExists("land.ini", index) Then
Set land = scene.CreateLandscape()
landdat() = Split(pacstr("land.ini", index), Chr(13) + Chr(10), , vbTextCompare)
land.LoadTerrainData pac.GetFile("land.ter", index)
land.SetScale CSng(landdat(2)), CSng(landdat(3)), CSng(landdat(4))
land.SetTexture tex.LoadTexture(pac.GetFile(landdat(1), index), "tera")
land.SetSplattingTexture tex.LoadTexture(pac.GetFile(landdat(7), index)), tex.LoadTexture(pac.GetFile(landdat(6), index)), -1
If UBound(landdat) > 27 Then land.SetMaterial GetMat(landdat(28))
scene.SetCamera CSng(landdat(15)), CSng(landdat(16)), CSng(landdat(17)), CSng(landdat(18)), CSng(landdat(19)), CSng(landdat(20))
Set sky = New TVAtmosphere
sky.SkyBox_SetTexture tex.LoadTexture(pac.GetFile("front.jpg", index), "fronta"), tex.LoadTexture(pac.GetFile("back.jpg", index), "backa") _
, tex.LoadTexture(pac.GetFile("left.jpg", index), "lefta"), tex.LoadTexture(pac.GetFile("right.jpg", index), "righta"), tex.LoadTexture( _
pac.GetFile("top.jpg", index), "topa"), tex.LoadTexture(pac.GetFile("bottom.jpg", index), "bottoma")
sky.SkyBox_Enable True
If landdat(22) = "真" Then
    sky.Sun_Enable True
    sky.Sun_SetTexture tex.LoadTexture(pac.GetFile("太阳.jpg", index), "suna")
    sky.Sun_SetBillboardSize landdat(23)
    sky.Sun_SetPosition landdat(24), landdat(25), landdat(26)
Else
    sky.Sun_Enable False
End If
landload = True
Else
landload = False
End If
End If

If load = 5 Then load = 6: loadmap = 1: Exit Function
If load = 6 Then
If pac.FileExists("NPC.dat", index) Then
If Not IsMissing(plan) Then plan = "载入角色......"
End If
If pac.FileExists("Mesh.dat", index) Then
meshdat() = Split(pacstr("Mesh.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim mesh(0 To UBound(meshdat))
For a = 0 To UBound(meshdat)
    meshdat2() = Split(meshdat(a), "|", , vbTextCompare)
    Set mesh(a) = scene.CreateMeshBuilder(meshdat2(0))
    Debug.Print meshdat2(0) + "名称", meshdat2(1) + "序号", meshdat2(2) + "文件", meshdat2(3) + "X", meshdat2(4) + "Y", meshdat2(5) + "Z", meshdat2(6) + "缩放X", meshdat2(7) + "缩放Y", meshdat2(8) + "缩放Z", meshdat2(9) + "旋转X", meshdat2(10) + "旋转Y", meshdat2(11) + "旋转Z", "模型列表"
    With mesh(a)
        .LoadTVM ("mesh\" + getstrleftb(meshdat2(2), ".") + ".tvm")
        If getstrrightb(meshdat2(2), ".") <> "tvm" Then .SetTexture (tex.LoadTexture(pac.GetFile("mesh\" + meshdat2(2), index), , , , TV_COLORKEY_BLACK))
        .SetPosition CSng(meshdat2(3)), CSng(meshdat2(4)), CSng(meshdat2(5))
        .SetScale CSng(meshdat2(6)), CSng(meshdat2(7)), CSng(meshdat2(8))
        .SetRotation CSng(meshdat2(9)), CSng(meshdat2(10)), CSng(meshdat2(11))
    End With
Next
meshload = True
Else
meshload = False
End If
End If

If load = 6 Then load = 7: loadmap = 1: Exit Function
If load = 7 Then
If pac.FileExists("particle.dat", index) Then
If Not IsMissing(plan) Then plan = "载入粒子......"
End If
If pac.FileExists("NPC.dat", index) Then
npcdat() = Split(pacstr("NPC.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim npc(0 To UBound(npcdat))
For a = 0 To UBound(npcdat)
    npcdat2() = Split(npcdat(a), "|", , vbTextCompare)
    Set npc(a) = scene.CreateActor(npcdat2(0))
    Debug.Print npcdat2(0) + "名称", npcdat2(1) + "序号", npcdat2(2) + "文件", npcdat2(3) + "X", npcdat2(4) + "Y", npcdat2(5) + "Z", npcdat2(6) + "缩放X", npcdat2(7) + "缩放Y", npcdat2(8) + "缩放Z", npcdat2(9) + "旋转X", npcdat2(10) + "旋转Y", npcdat2(11) + "旋转Z", npcdat2(12) + "动画ID", "角色列表"
    With npc(a)
        .LoadTVA ("NPC\" + getstrleftb(npcdat2(2), ".") + ".tva")
        If getstrrightb(npcdat2(2), ".") <> "tva" Then .SetTexture (tex.LoadTexture(pac.GetFile("NPC\" + npcdat2(2), index), , , , TV_COLORKEY_BLACK))
        .SetPosition CSng(npcdat2(3)), CSng(npcdat2(4)), CSng(npcdat2(5))
        .SetScale CSng(npcdat2(6)), CSng(npcdat2(7)), CSng(npcdat2(8))
        .SetRotation CSng(npcdat2(9)), CSng(npcdat2(10)), CSng(npcdat2(11))
        .SetAnimationID CLng(npcdat2(12))
        .SetAnimationLoop True
        .PlayAnimation
    End With
Next
npcload = True
Else
npcload = False
End If
End If

If load = 7 Then load = 8: loadmap = 1: Exit Function
If load = 8 Then
If pac.FileExists("Mat\Mesh.dat", index) Or pac.FileExists("Mat\NPC.dat", index) Then
If Not IsMissing(plan) Then plan = "释放材质到模型与角色......"
End If
If pac.FileExists("particle.dat", index) Then
pardat = Split(pacstr("particle.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim par(0 To UBound(pardat))
For a = 0 To UBound(pardat)
    pardat2 = Split(pardat(a), "|", , vbTextCompare)
    Set par(a) = scene.CreateParticleSystem(getstrlefta(pardat(a), "."))
    With par(a)
    .load pac.GetFile("particle\" + getstrlefta(pardat(a), "|"))
    .SetGlobalPosition CSng(pardat2(1)), CSng(pardat2(2)), CSng(pardat2(3))
    .SetGlobalScale CSng(pardat2(4)), CSng(pardat2(5)), CSng(pardat2(6)), CSng(pardat2(7))
    .SetGlobalRotation CSng(pardat2(8)), CSng(pardat2(9)), CSng(pardat2(10))
    .Enable True
    End With
    Debug.Print getstrlefta(pardat2(0), ".") + "名称", pardat2(1) + "X", pardat2(2) + "Y", pardat2(3) + "Z", pardat2(4) + "缩放X", pardat2(5) + "缩放Y", pardat2(6) + "缩放Z", pardat2(7) + "缩放平面", pardat2(8) + "旋转X", pardat2(9) + "旋转Y", pardat2(10) + "旋转Z", "粒子列表"
    Form1.List5.AddItem getstrlefta(pardat2(0), ".") + "|" + pardat2(4) + "|" + pardat2(5) + "|" + pardat2(6) + "|" + pardat2(7) '多出来的
Next
parload = True
Else
parload = False
End If
End If

If load = 8 Then load = 9: loadmap = 1: Exit Function
If load = 9 Then
If pac.FileExists("sou.dat", index) Then
If Not IsMissing(plan) Then plan = "载入音效中......"
End If
If pac.FileExists("Mat\Mesh.dat", index) Then
meshmat = Split(pacstr("Mat\Mesh.dat", index), Chr(13) + Chr(10), , vbTextCompare)
For a = 0 To UBound(meshmat)
    meshmat2 = Split(meshmat(a), "|", , vbTextCompare)
    Debug.Print meshmat2(0) + "模型", meshmat2(1) + "材质", "模型材质列表"
    glo.getmesh(meshmat2(0)).SetMaterial GetMat(meshmat2(1))
    Form1.List3.AddItem meshmat2(0) + "|" + meshmat2(1) '多出来的
Next
End If
End If

If load = 9 Then load = 10: loadmap = 1: Exit Function
If load = 10 Then
If pac.FileExists("Mat\NPC.dat", index) Then
npcmat = Split(pacstr("Mat\NPC.dat", index), Chr(13) + Chr(10), , vbTextCompare)
For a = 0 To UBound(npcmat)
    npcmat2 = Split(npcmat(a), "|", , vbTextCompare)
    Debug.Print npcmat2(0) + "角色", npcmat2(1) + "材质", "角色材质列表"
    glo.GetActor(npcmat2(0)).SetMaterial GetMat(npcmat2(1))
    Form1.List4.AddItem npcmat2(0) + "|" + npcmat2(1) '多出来的
Next
End If
End If

If load = 10 Then load = 11: loadmap = 1: Exit Function
If load = 11 Then
If pac.FileExists("other.dat", index) Then
If Not IsMissing(plan) Then plan = "解压资源文件......"
End If
If pac.FileExists("sou.dat", index) Then
If fi.FolderExists(App.Path + "\sou\") Then
Set fil = fi.GetFolder(App.Path + "\sou\")
fil.Delete True
MkDir App.Path + "\sou\"
Else
MkDir App.Path + "\sou\"
End If
soudat = Split(pacstr("sou.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim soulist(0 To UBound(soudat))
For a = o To UBound(soudat)
pac.ExtractFile "sou\" + soudat(a), App.Path + "\sou\" + soudat(a), index
'sous.AddFile App.Path + "\sou\" + soudat(a)
Form1.List6.AddItem soudat(a) '多出来的
soulist(a) = getstrleftb(soudat(a), ".")
Next
Dim abc As FileSystemObject
Set abc = New FileSystemObject
abc.DeleteFolder App.Path + "\sou", True
lig.SetSpecularLighting True
souload = True
Else
souload = False
End If
End If

If load = 11 Then load = 12: loadmap = 1: Exit Function
If load = 12 Then
If pac.FileExists("Main.lua", index) Then
If Not IsMissing(plan) Then plan = "加载脚本......"
End If
If pac.FileExists("other.dat", index) Then
If fi.FolderExists(App.Path + "\other\") Then
Set fill = fi.GetFolder(App.Path + "\other\")
fill.Delete True
MkDir App.Path + "\other\"
Else
MkDir App.Path + "\other\"
End If
othdat = Split(pacstr("other.dat", index), Chr(13) + Chr(10), , vbTextCompare)
ReDim othlist(o To UBound(othdat))
For a = 0 To UBound(othdat)
pac.ExtractFile "other\" + othdat(a), App.Path + "\other\" + othdat(a), index
SetAttr App.Path + "\other\" + othdat(a), vbReadOnly
othlist(a) = App.Path + "\other\" + othdat(a)
Next
othload = True
Else
othload = False
End If
End If

If load = 12 Then load = 13: loadmap = 1: Exit Function
If load = 13 Then
'runlua '调试
If Not IsMissing(plan) Then plan = "结束载入......"
If pac.FileExists("Main.lua", index) Then
pac.ExtractFile "Main.lua", App.Path + "\Main.lua", index
'runlua
Kill App.Path + "\Main.lua"
'Kill App.Path + "\game.lua"
If pac.FileExists("Logic.lua", index) Then pac.ExtractFile "Logic.lua", App.Path + "\Logic.lua", index
If pac.FileExists("Render.lua", index) Then pac.ExtractFile "Render.lua", App.Path + "\Render.lua", index
If pac.FileExists("Destroy.lua", index) Then pac.ExtractFile "Destroy.lua", App.Path + "\Destroy.lua", index
jbload = True
Else
jbload = False
End If
End If

If load = 13 Then load = 14: loadmap = 1: Exit Function
If load = 14 Then
If Not IsMissing(plan) Then plan = "完成！"
meshcount = UBound(mesh)
npccount = UBound(npc)
pac.ClosePackage index
iniload = True
load = 0
loadmap = 2
End If

End Function
Sub SubLogic()
'runlua2
End Sub
Sub SubRender()
'runlua3
End Sub
Sub destroy()
Dim index As Long
If iniload = True Then
If landload = True Then
land.destroy
Set sky = New TVAtmosphere
End If

If meshload = True Then
For index = 0 To meshcount
    mesh(index).destroy
Next
End If

If npcload = True Then
For index = 0 To npccount
    npc(index).destroy
Next
End If

If parload = True Then
For index = 0 To UBound(par)
    par(index).destroy
Next
End If

If lightload = True Then
For index = 0 To UBound(liga)
lig.DeleteLight liga(index)
Next
End If

If matload = True Then
For index = 0 To UBound(mata)
mat.DeleteMaterial mata(index)
Next
End If

If texload = True Then
For index = 0 To UBound(texa)
tex.DeleteTexture texa(index)
Next
End If

'If souload = True Then
'For index = 0 To UBound(soulist)
'sous.Remove soulist(index)
'Next
'End If

If othload = True Then
Set fi = CreateObject("Scripting.FileSystemObject")
For index = 0 To UBound(othlist)
If fi.FileExists(othlist(index)) Then Kill othlist(index)
Next
End If

If jbload = True Then
'runlua4
If fi.FileExists(App.Path + "\Destroy.lua") Then Kill App.Path + "\Destroy.lua"
If fi.FileExists("Logic.lua", index) Then Kill App.Path + "\Logic.lua"
If fi.FileExists("Render.lua", index) Then Kill App.Path + "\Render.lua"
End If
End If
End Sub
Sub Render()
Dim index As Long
sky.Atmosphere_Render
For index = 0 To meshcount
    mesh(index).Render
Next
For index = 0 To npccount
    npc(index).Render
Next
If zr Then '多出来的
For index = 0 To UBound(par)
    par(index).Update
    par(index).Render
Next
End If
land.Render
End Sub
Function getmeshformname(name As String, Optional isinlist As Boolean) As TVMesh
Dim a As Long
For a = 0 To meshcount
If mesh(a).GetMeshName = name Then Set getmeshformname = mesh(a): If Not IsMissing(isinlist) Then isinlist = True: Exit Function
Next
Set getmeshformname = mesh(0)
If Not IsMissing(isinlist) Then isinlist = False
End Function
Function getnpcformname(name As String, Optional isinlist As Boolean) As TVActor
Dim a As Long
For a = 0 To npccount
If npc(a).GetName = name Then Set getnpcformname = npc(a): If Not IsMissing(isinlist) Then isinlist = True: Exit Function
Next
Set getnpcformname = npc(0)
If Not IsMissing(isinlist) Then isinlist = False
End Function
Function getparformname(name As String, Optional isinlist As Boolean) As TVParticleSystem
Dim a As Long
For a = 0 To UBound(par)
If par(a).GetName = name Then Set getparformname = par(a): If Not IsMissing(isinlist) Then isinlist = True: Exit Function
Next
Set getparformname = par(0)
If Not IsMissing(isinlist) Then isinlist = False
End Function
Public Function getter() As TVLandscape
Set getter = land
End Function
Public Function getpar(index As Long) As TVParticleSystem
Set getpar = par(index)
End Function
Function getsky() As TVAtmosphere
Set getsky = sky
End Function
Function getmesh(index As Long) As TVMesh
Set getmesh = mesh(index)
End Function
Function getnpc(index As Long) As TVActor
Set getnpc = npc(index)
End Function
Sub Resetland(scene As TVScene, str As String)
land.destroy
Set land = scene.CreateLandscape(str)
End Sub
Sub Resetsky()
Set sky = New TVAtmosphere
End Sub
Sub Resetmesh(scene As TVScene, index As Long, str As String)
mesh(index).destroy
Set mesh(index) = scene.CreateMeshBuilder(str)
End Sub
Sub Resetnpc(scene As TVScene, index As Long, str As String)
npc(index).destroy
Set npc(index) = scene.CreateActor(str)
End Sub
Sub Resetpar(scene As TVScene, index As Long, str As String)
par(index).destroy
Set par(index) = scene.CreateParticleSystem(str)
End Sub
Sub Resetall(scene As TVScene)
Static index As Long
Set sky = New TVAtmosphere
For index = 0 To meshcount
    mesh(index).Resetmesh
Next
For index = 0 To npccount
    npc(index).destroy
    Set npc(index) = scene.CreateActor
Next
For index = 0 To UBound(par)
    par(index).destroy
    Set par(index) = scene.CreateParticleSystem
Next
land.destroy
Set land = scene.CreateLandscape
End Sub
Sub createland(scene As TVScene, str As String)
Set land = scene.CreateLandscape(str)
End Sub
Sub createsky()
Set sky = New TVAtmosphere
End Sub
Sub createmeshnpcarray(arraymesh As Long, arraynpc As Long, scene As TVScene, meshstr() As String, npcstr() As String)
Static index As Long
ReDim mesh(0 To arraymesh)
ReDim npc(0 To arrayanpc)
meshcount = arraymesh
npccount = arraynpc
For index = 0 To meshcount
    Set mesh(index) = scene.CreateMeshBuilder(meshstr(index))
Next
For index = 0 To npccount
    Set npc(index) = scene.CreateActor(npcstr(index))
Next
End Sub
Function getmeshcount() As Long
getmeshcount = meshcount
End Function
Function getnpccount() As Long
getnpccount = npccount
End Function
Function getparcount() As Long
getparcount = UBound(par)
End Function
Sub addmesh(mesha As TVMesh, Optional name As String)
ReDim Preserve mesh(0 To UBound(mesh) + 1)
meshcount = UBound(mesh)
If IsMissing(name) Then
Set mesh(meshcount) = mesha.Duplicate
Else
Set mesh(meshcount) = mesha.Duplicate(name)
End If
End Sub
Sub addnpc(npca As TVActor, Optional name As String)
ReDim Preserve npc(0 To UBound(npc) + 1)
npccount = UBound(npc)
If IsMissing(name) Then
Set npc(npccount) = npca.Duplicate
Else
Set npc(npccount) = npca.Duplicate(name)
End If
End Sub
Sub addpar(ByVal para As TVParticleSystem, Optional name As String)
ReDim Preserve par(0 To UBound(par) + 1)
If IsMissing(name) Then
Set par(UBound(par)) = para
par(UBound(par)).Enable True
Else
para.SetName name
Set par(UBound(par)) = para
par(UBound(par)).Enable True
End If
End Sub

