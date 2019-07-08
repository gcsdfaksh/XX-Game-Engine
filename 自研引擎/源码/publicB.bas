Attribute VB_Name = "publicB"
Public xx As Long
Public yy As Single
Public zz As Single
Public x As Single
Public y As Single
Public z As Single
Public xwg As Boolean
Public zwg As Boolean
Public ywg As Boolean
Public ggg() As Long
Public kg As Boolean
Public ge() As TVMesh
Public kg2 As Boolean
Public kg3 As Boolean
Public gee As TVMesh
Public gef As TVMesh
Public pa() As TVMesh
Public str2 As String
Public str3 As String
Public zr As Boolean
Public kg4 As Boolean
Function isinlist(lista As ListBox, str As String) As Boolean
Dim c As Long
Dim hj As Boolean
For c = 0 To lista.ListCount - 1
If getstrlefta(lista.List(c), "|") = str Then hj = True: Exit For
Next
isinlist = hj
End Function
Function inlist(lista As ListBox, str As String) As Long
Dim c As Long
Dim hj As Long
For c = 0 To lista.ListCount - 1
If getstrlefta(lista.List(c), "|") = str Then hj = c: Exit For
Next
inlist = hj
End Function
Function isinlistex(lista As ListBox, str As String) As Boolean
Dim c As Long
Dim hj As Boolean
For c = 0 To lista.ListCount - 1
If lista.List(c) = str Then hj = True: Exit For
Next
isinlistex = hj
End Function
Function inlistex(lista As ListBox, str As String) As Long
Dim c As Long
Dim hj As Long
For c = 0 To lista.ListCount - 1
If glista.List(c) = str Then hj = c: Exit For
Next
inlistex = hj
End Function
