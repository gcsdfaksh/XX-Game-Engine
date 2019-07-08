Attribute VB_Name = "luajb"
'Public Function Lua_VB_InputBox(ByVal hState As Long) As Long
'Dim s As String, s2 As String, s3 As String
'Dim m As Long
'm = lua_gettop(hState)
'If m >= 1 Then s = lua_tostring(hState, 1)
'If m >= 2 Then s2 = lua_tostring(hState, 2) Else s2 = App.Title
'If m >= 3 Then s3 = lua_tostring(hState, 3)
''///TODO:etc.
's = InputBox(s, s2, s3)
''///
'lua_pushstring hState, s
'Lua_VB_InputBox = 1
'End Function
'
'Public Function Lua_VB_MsgBox(ByVal hState As Long) As Long
'Dim s As String, s3 As String
'Dim i As Long
'Dim m As Long
'm = lua_gettop(hState)
'If m >= 1 Then s = lua_tostring(hState, 1)
'If m >= 2 Then i = lua_tointeger(hState, 2)
'If m >= 3 Then s3 = lua_tostring(hState, 3) Else s3 = App.Title
''///TODO:etc.
'i = MsgBox(s, i, s3)
''///
'lua_pushinteger hState, i
'Lua_VB_MsgBox = 1
'End Function
'Public Function testa(ByVal hState As Long) As Long
'Dim m As Long
'Dim i As Single
'm = lua_gettop(hState)
'If m >= 1 Then i = lua_tonumber(hState, 1) Else i = 60
'scene.GetCamera.SetViewFrustum i, 1000
'testa = 1
'End Function
'Public Function testb(ByVal hState As Long) As Long
'Dim m As Long
'Dim a As Single
'Dim b As Single
'scene.GetCamera.GetViewFrustum a, b 'SetViewFrustum i, 1000
'lua_pushnumber hState, a
'testb = 1
'End Function
'Public Sub inimain(ByVal hState As Long)
''hState = lua_getglobal(hState, "aaa")
'lua_newtable hState
''///
'lua_pushstring hState, "inputbox"
'lua_pushcclosure hState, AddressOf Lua_VB_InputBox, 0
'lua_rawset hState, -3
'lua_pushstring hState, "msgbox"
'lua_pushcclosure hState, AddressOf Lua_VB_MsgBox, 0
'lua_rawset hState, -3
'lua_pushstring hState, "testa"
'lua_pushcclosure hState, AddressOf testa, 0
'lua_rawset hState, -3
'lua_pushstring hState, "testb"
'lua_pushcclosure hState, AddressOf testb, 0
'lua_rawset hState, -3
''lua_pushstring hState, "testa"
''lua_pushcclosure hState, AddressOf testa, 0
''lua_rawset hState, -1
'
''///
'lua_setglobal hState, "m"
'End Sub
'Public Sub inilogic(ByVal hState As Long)
'lua_newtable hState
''///
''///
'lua_setglobal hState, "l"
'End Sub
'Public Sub inirender(ByVal hState As Long)
'lua_newtable hState
''///
'lua_pushstring hState, "msgbox"
'lua_pushcclosure hState, AddressOf Lua_VB_MsgBox, 0
'lua_rawset hState, -3
''///
'lua_setglobal hState, "r"
'End Sub
'Public Sub inidestroy(ByVal hState As Long)
'lua_newtable hState
''///
''///
'lua_pushstring hState, "msgbox"
'lua_pushcclosure hState, AddressOf Lua_VB_MsgBox, 0
'lua_rawset hState, -3
'lua_setglobal hState, "d"
'End Sub
'Public Function test(ByVal hState As Long) As Long
'Dim i As Long
'Dim s As String
''///
'Debug.Print hState, "sdwd"
'lua_pop hState, 1 'discard user data
''///
'lua_gc hState, LUA_GCSTOP, 0 '  /* stop collector during initialization */
'luaopen_base hState  '  /* open libraries */
''luaL_openlibs hState
''///
'inimain hState
'lua_gc hState, LUA_GCRESTART, 0
''///
'
''i = luaL_loadfile(hState, App.Path + "\primetest.lua")
'i = luaL_loadfile(hState, App.Path + "\Main.lua")
''luaL_loadstring (hState,)
'
'If i Then
'  Select Case i
'  Case LUA_ERRSYNTAX
'    s = "Syntax error in file"
'  Case Else
'    s = "Can't open file"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'i = lua_pcall(hState, 0, 0, 0)
'If i Then
'  Select Case i
'  Case LUA_ERRRUN
'    s = "Runtime error"
'  Case LUA_ERRMEM
'    s = "Memory error"
'  Case Else
'    s = "Unknown error"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'End Function
'Public Function test2(ByVal hState As Long) As Long
'Dim i As Long
'Dim s As String
''///
'lua_pop hState, 1 'discard user data
''///
'lua_gc hState, LUA_GCSTOP, 0 '  /* stop collector during initialization */
'luaopen_base hState  '  /* open libraries */
''luaL_openlibs hState
''///
'inilogic hState
'lua_gc hState, LUA_GCRESTART, 0
''///
'
''i = luaL_loadfile(hState, App.Path + "\primetest.lua")
'i = luaL_loadfile(hState, App.Path + "\Logic.lua")
''luaL_loadstring (hState,)
'
'If i Then
'  Select Case i
'  Case LUA_ERRSYNTAX
'    s = "Syntax error in file"
'  Case Else
'    s = "Can't open file"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'i = lua_pcall(hState, 0, 0, 0)
'If i Then
'  Select Case i
'  Case LUA_ERRRUN
'    s = "Runtime error"
'  Case LUA_ERRMEM
'    s = "Memory error"
'  Case Else
'    s = "Unknown error"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'End Function
'Public Function test3(ByVal hState As Long) As Long
'Dim i As Long
'Dim s As String
''///
'lua_pop hState, 1 'discard user data
''///
'lua_gc hState, LUA_GCSTOP, 0 '  /* stop collector during initialization */
'luaopen_base hState  '  /* open libraries */
''luaL_openlibs hState
''///
'inirender hState
'lua_gc hState, LUA_GCRESTART, 0
''///
'
''i = luaL_loadfile(hState, App.Path + "\primetest.lua")
'i = luaL_loadfile(hState, App.Path + "\Render.lua")
''luaL_loadstring (hState,)
'
'If i Then
'  Select Case i
'  Case LUA_ERRSYNTAX
'    s = "Syntax error in file"
'  Case Else
'    s = "Can't open file"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'i = lua_pcall(hState, 0, 0, 0)
'If i Then
'  Select Case i
'  Case LUA_ERRRUN
'    s = "Runtime error"
'  Case LUA_ERRMEM
'    s = "Memory error"
'  Case Else
'    s = "Unknown error"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'End Function
'Public Function test4(ByVal hState As Long) As Long
'Dim i As Long
'Dim s As String
''///
'lua_pop hState, 1 'discard user data
''///
'lua_gc hState, LUA_GCSTOP, 0 '  /* stop collector during initialization */
'luaopen_base hState  '  /* open libraries */
''luaL_openlibs hState
''///
'inidestroy hState
'lua_gc hState, LUA_GCRESTART, 0
''///
'
''i = luaL_loadfile(hState, App.Path + "\primetest.lua")
'i = luaL_loadfile(hState, App.Path + "\Destroy.lua")
''luaL_loadstring (hState,)
'
'If i Then
'  Select Case i
'  Case LUA_ERRSYNTAX
'    s = "Syntax error in file"
'  Case Else
'    s = "Can't open file"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'i = lua_pcall(hState, 0, 0, 0)
'If i Then
'  Select Case i
'  Case LUA_ERRRUN
'    s = "Runtime error"
'  Case LUA_ERRMEM
'    s = "Memory error"
'  Case Else
'    s = "Unknown error"
'  End Select
'  MsgBox s, vbCritical
'  Exit Function
'End If
'End Function
'Public Sub runlua()
'Dim hState, i As Long
'hState = lua_open
'If hState Then
'  i = lua_cpcall(hState, AddressOf test, 0)
'  If i Then MsgBox "Error"
'  lua_close hState
'  hState = 0
'Else
'  MsgBox "cannot create state: not enough memory", vbCritical
'End If
'End Sub
'Public Sub runlua2()
'Dim hState, i As Long
'hState = lua_open
'If hState Then
'  i = lua_cpcall(hState, AddressOf test2, 0)
'  If i Then MsgBox "Error"
'  lua_close hState
'  hState = 0
'Else
'  MsgBox "cannot create state: not enough memory", vbCritical
'End If
'End Sub
'Public Sub runlua3()
'Dim hState, i As Long
'hState = lua_open
'If hState Then
'  i = lua_cpcall(hState, AddressOf test3, 0)
'  If i Then MsgBox "Error"
'  lua_close hState
'  hState = 0
'Else
'  MsgBox "cannot create state: not enough memory", vbCritical
'End If
'End Sub
'Public Sub runlua4()
'Dim hState, i As Long
'hState = lua_open
'If hState Then
'  i = lua_cpcall(hState, AddressOf test4, 0)
'  If i Then MsgBox "Error"
'  lua_close hState
'  hState = 0
'Else
'  MsgBox "cannot create state: not enough memory", vbCritical
'End If
'End Sub
