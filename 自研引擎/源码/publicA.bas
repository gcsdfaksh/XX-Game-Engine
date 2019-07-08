Attribute VB_Name = "publicA"
Public tv As TVEngine
Public scene As TVScene
Public inp As TVInputEngine
Public map As mapclass
Public pac As TVPackage
Public tex As TVTextureFactory
Public scr2D As TVScreen2DImmediate
Public Const white As Long = -1, black As Long = -16777216, black3 As Long = 1275068416, black5 As Long = 2130706432, black7 As Long = -1308622848, black9 As Long = -452984832, yellow As Long = _
-256, green As Long = -16646399, blue As Long = -16776961, alpha As Long = 0, red As Long = -65536
Public mat As TVMaterialFactory
Public glo As TVGlobals
Public lig As TVLightEngine
Public col As TVCollisionResult
'Public sou As TVSoundEngine
'Public sous As TVSounds
Type animal
    pic(0 To 0) As Long
    pos(0 To 0) As TV_2DVECTOR
    time(0 To 0) As Long
    nowtime As Long
    nowfra As Long
    des As Boolean
End Type
Sub iniall() '可做更改
Set tv = New TVEngine
Set inp = New TVInputEngine
Set scene = New TVScene
Set pac = New TVPackage
Set tex = New TVTextureFactory
Set map = New mapclass
Set scr2D = New TVScreen2DImmediate
Set glo = New TVGlobals
Set mat = New TVMaterialFactory
Set glo = New TVGlobals
Set lig = New TVLightEngine
Set col = New TVCollisionResult
Set par = New TVParticleSystem
'Set sou = New TVSoundEngine
tv.AllowMultithreading True
tv.Init3DWindowed Form1.Picture1.hWnd, True
tv.GetViewport.SetAutoResize True
tv.DisplayFPS True
tv.SetAngleSystem TV_ANGLE_DEGREE
tex.SetTextureMode TV_TEXTUREMODE_PALETTIZED8BITS
scene.SetViewFrustum 60, 2000
inp.Initialize True, True
scr2D.Settings_SetTextureFilter TV_FILTER_TRILINEAR
'sou.Init Form1.hWnd
'Set sous = sou.CreateSounds
End Sub
Sub desall() '可做更改
If zr = True Then map.destroy
Set inp = Nothing
Set scene = Nothing
Set tex = Nothing
Set pac = Nothing
Set map = Nothing
Set scr2D = Nothing
Set glo = Nothing
Set mat = Nothing
Set glo = Nothing
Set col = Nothing
Set par = Nothing
Set sou = Nothing
Set sous = Nothing
tv.ReleaseAll
End Sub
Function pacstr(file As String, ID As Long) As String
Dim a As Integer
Dim dat As String
Dim file2 As String
Dim NextLine As String
a = FreeFile
If InStrRev(file, "\") <> 0 Then file2 = getstrrightb(file, "\") Else file2 = file
pac.ExtractFile file, App.Path + "\" + file2, ID
Open App.Path + "\" + file2 For Input As a
dat = StrConv(InputB(LOF(a), a), vbUnicode)
Close a
'debug.Print dat, "adasd"
Kill App.Path + "\" + file2
pacstr = dat
End Function
Function getstrlefta(stringa As String, stringb As String) As String
getstrlefta = Left(stringa, InStr(stringa, stringb) - 1)
End Function
Function getstrrighta(stringa As String, stringb As String) As String
getstrrighta = Right(stringa, Len(stringa) - InStr(stringa, stringb))
End Function
Function getstrleftb(stringa As String, stringb As String) As String
getstrleftb = Left(stringa, InStrRev(stringa, stringb) - 1)
End Function
Function getstrrightb(stringa As String, stringb As String) As String
getstrrightb = Right(stringa, Len(stringa) - InStrRev(stringa, stringb))
End Function
Function cre3d(x As Single, y As Single, z As Single) As TV_3DVECTOR
Dim q As TV_3DVECTOR
q.x = x
q.y = y
q.z = z
cre3d = q
End Function
Function cre2d(x As Single, y As Single) As TV_2DVECTOR
Dim q As TV_2DVECTOR
q.x = x
q.y = y
cre2d = q
End Function
Function loadmat(file As String, index As Long) As Long
Dim a As Long
Dim b() As String
Dim c() As Long
Dim d As Long
b = Split(pacstr(file, index), Chr(13) + Chr(10), , vbTextCompare)
ReDim c(0 To UBound(b))
For d = 0 To UBound(b)
    c(d) = CSng(b(d))
    Debug.Print c(d), ""
Next
With mat
If InStrRev(file, "\") <> 0 Then a = .CreateMaterial(getstrleftb(getstrrightb(file, "\"), ".")) Else a = .CreateMaterial(getstrleftb(file, "."))
.SetAmbient a, c(0), c(1), c(2), c(3)
.SetDiffuse a, c(4), c(5), c(6), c(7)
.SetEmissive a, c(8), c(9), c(10), c(11)
.SetSpecular a, c(12), c(13), c(14), c(15)
.SetPower a, c(16)
.SetOpacity a, c(17)
If c(18) = 1 Then .EnablePRTSubSurface a, True: .SetPRTSubSurfReducedScattering a, c(19), c(20), c(21): .SetPRTSubSurfAbsorption a, c(22), c(23), c(24): .SetPRTSubSurfRefractionIndexRatio a, c(25)
End With
loadmat = a
End Function
Function loadlight(file As String, index As Long) As Long
Dim j As Long
Dim k As String
Dim a() As String
If InStrRev(file, "\") <> 0 Then k = getstrrightb(getstrleftb(file, "."), "\") Else k = getstrleftb(file, ".")
a = Split(pacstr(file, index), Chr(13) + Chr(10), , vbTextCompare)
Select Case a(0)
Case "1"
j = lig.CreatePointLight(cre3d(CSng(a(1)), CSng(a(2)), CSng(a(3))), CSng(a(4)), CSng(a(5)), CSng(a(6)), CSng(a(7)), k, CSng(a(8)))
lig.SetLightProperties j, IIf(a(9) = "1", True, False), IIf(a(10) = "1", True, False), IIf(a(11) = "1", True, False)
If UBound(a) > 11 Then If (a(12) <> "") Then lig.SetLightCubeMap j, GetTex(a(12))
If UBound(a) > 12 Then If a(12) <> "" Then lig.SetProjectiveShadowsProperties j, CLng(a(13)), CLng(a(14))
If UBound(a) > 14 Then If a(15) <> "" Then lig.SetLightColor j, CSng(a(15)), CSng(a(16)), CSng(a(17))
If UBound(a) > 17 Then If a(18) <> "" Then lig.SetLightSpecularColor j, CSng(a(18)), CSng(a(19)), CSng(a(20))
If UBound(a) > 20 Then If a(21) <> "" Then lig.SetLightDiffuseColor j, CSng(a(21)), CSng(a(22)), CSng(a(23))
Case "2"
j = lig.CreateSpotLight(cre3d(CSng(a(1)), CSng(a(2)), CSng(a(3))), cre3d(CSng(a(4)), CSng(a(5)), CSng(a(6))), CSng(a(7)), CSng(a(8)), CSng(a(9)), CSng(a(10)), CSng(a(12)), CSng(a(13)), k, CSng(a(14)))
lig.SetLightProperties j, IIf(a(15) = "1", True, False), IIf(a(16) = "1", True, False), IIf(a(17) = "1", True, False)
If UBound(a) > 17 Then If a(12) <> "" Then lig.SetLightAttenuation j, CSng(a(18)), CSng(a(19)), CSng(a(20)): lig.SetLightSpotFalloff j, CSng(a(21))
If UBound(a) > 21 Then If a(22) <> "" Then lig.SetLightCubeMap j, GetTex(a(22))
If UBound(a) > 22 Then If a(23) <> "" Then lig.SetProjectiveShadowsProperties j, CLng(a(23)), CLng(a(24))
If UBound(a) > 24 Then If a(25) <> "" Then lig.SetLightColor j, CSng(a(25)), CSng(a(26)), CSng(a(27))
If UBound(a) > 27 Then If a(28) <> "" Then lig.SetLightSpecularColor j, CSng(a(28)), CSng(a(29)), CSng(a(30))
If UBound(a) > 30 Then If a(31) <> "" Then lig.SetLightDiffuseColor j, CSng(a(31)), CSng(a(32)), CSng(a(33))
Case "3"
j = lig.CreateDirectionalLight(cre3d(CSng(a(1)), CSng(a(2)), CSng(a(3))), CSng(a(4)), CSng(a(5)), CSng(a(6)), k, CSng(a(7)))
lig.SetLightProperties j, IIf(a(8) = "1", True, False), IIf(a(9) = "1", True, False), IIf(a(10) = "1", True, False)
If UBound(a) > 10 Then If a(11) <> "" Then lig.SetLightCubeMap j, GetTex(a(11))
If UBound(a) > 11 Then If a(12) <> "" Then lig.SetProjectiveShadowsProperties j, CLng(a(12)), CLng(a(13))
If UBound(a) > 13 Then If a(14) <> "" Then lig.SetLightColor j, CSng(a(14)), CSng(a(15)), CSng(a(16))
If UBound(a) > 16 Then If a(17) <> "" Then lig.SetLightSpecularColor j, CSng(a(17)), CSng(a(18)), CSng(a(19))
If UBound(a) > 19 Then If a(20) <> "" Then lig.SetLightDiffuseColor j, CSng(a(20)), CSng(a(21)), CSng(a(22))
Case Else
End Select
loadlight = j
End Function
Function loadtex(file As String, index As Long, Optional aalpha As Long, Optional atype As Long) As Long '0普通纹理,13D纹理，2凹凸纹理，3立方纹理，4DUVU纹理，4Alpha纹理，
Dim j As Long
Dim k As String
Dim l As Long
If InStrRev(file, "\") <> 0 Then k = getstrrightb(getstrleftb(file, "."), "\") Else k = getstrleftb(file, ".")
If IsMissing(aahpla) Then l = 0 Else l = aahpla
If IsMissing(atype) Then
    j = tex.LoadTexture(pac.GetFile(file, index), k, ecolorkey:=l)
Else
    Select Case atype
    Case 1
        j = tex.LoadVolumeTexture(pac.GetFile(file, index), k, l, True)
    Case 2
        j = tex.LoadBumpTexture(pac.GetFile(file, index), k, , , True)
    Case 3
        j = tex.LoadCubeTexture(pac.GetFile(file, index), k, , l, True)
    Case 4
        j = tex.LoadDUDVTexture(pac.GetFile(file, index), k)
    Case 5
        j = tex.LoadAlphaTexture(pac.GetFile(file, index), k)
    Case Else
        j = tex.LoadTexture(pac.GetFile(file, index), k, ecolorkey:=l)
    End Select
End If
LoadTexture = j
End Function
