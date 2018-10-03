VERSION 5.00
Begin VB.Form tsendvta 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia Datos Web"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8340
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "CancelaProceso..."
      Height          =   975
      Left            =   3720
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Transferencia"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label registro 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado del proceso"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu dlo789923 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsendvta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command2.Visible = True Then Exit Sub
Command2.Visible = True
If Combo1 = "Compras/Ventas" Then
   prepara_ventas1
End If
If Combo1 = "Productos" Then
   envia_formato_orion3
   'envia_productos
End If
Command2.Visible = False
End Sub
Sub prepara_ventas1()
Dim buf As String
Dim xbuf As String
Dim xbuf1 As String
Dim xbuf2 As String

Dim mytablex As Snapshot
Dim i As Integer
Dim xlinea As String
Dim xcount As Long

Dim vr
On Error GoTo cmd1_err
If Len(fechai) <> 10 Then Exit Sub
If Len(fechaf) <> 10 Then Exit Sub
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
xbuf = globalweb + "\01C" + Format(Now, "ddmmyyyy")
xbuf1 = "\rp_orion.v2\001d\01\web\" + "01D" + Format(Now, "ddmmyyyy")
xbuf2 = "\rp_orion.v2\001d\01\web\" + "01F" + Format(Now, "ddmmyyyy")
Open xbuf For Output As #1

buf = "select * from factura where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
 
Close #1
prepara_ventas2
prepara_ventas3
If xcount > 0 Then  'ejecutar el dos
   Call ExecuteCommand("\rp_orion.v2\001d\01\web\subir.bat scanpos " & xbuf & " " & xbuf1 & " " & xbuf2)
End If
MsgBox "Proceso Finalizado " & xcount, 48, "Aviso"
Exit Sub
cmd1_err:
MsgBox "Error " & error$, 48, "Aviso"
Exit Sub

End Sub


Private Sub Command2_Click()
Command2.Visible = False
End Sub

Private Sub dlo789923_Click()
If Command2.Visible = True Then
   Command2.Visible = False
   Exit Sub
End If
tsendvta.Hide
Unload tsendvta
End Sub

Private Sub Form_Load()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
Combo1.Clear
Combo1.AddItem "*"
Combo1.AddItem "Productos"
'Combo1.AddItem "Compras/Ventas"
Combo1.ListIndex = 0
End Sub
Sub prepara_ventas2()
Dim buf As String

Dim mytablex As Snapshot
Dim i As Integer
Dim xlinea As String
Dim xcount As Long
Dim vr
On Error GoTo cmd2_err
If Len(fechai) <> 10 Then Exit Sub
If Len(fechaf) <> 10 Then Exit Sub
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
xbuf = "\rp_orion.v2\001d\01\web\" + "01D" + Format(Now, "ddmmyyyy")
Open xbuf For Output As #1

buf = "select * from detalle where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
If Command2.Visible = False Then Exit Do
registro = "" & xcount
r = DoEvents()
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
 
Close #1
Exit Sub
cmd2_err:
MsgBox "Error " & error$, 48, "Aviso"
Exit Sub
End Sub
Sub prepara_ventas3()
Dim buf As String

Dim mytablex As Snapshot
Dim i As Integer
Dim xlinea As String
Dim xcount As Long
Dim vr
On Error GoTo cmd3_err
If Len(fechai) <> 10 Then Exit Sub
If Len(fechaf) <> 10 Then Exit Sub
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
xbuf = "\rp_orion.v2\001d\01\web\" + "01F" + Format(Now, "ddmmyyyy")
Open xbuf For Output As #1

buf = "select * from fpagov where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
If Command2.Visible = False Then Exit Do
vr = DoEvents()
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
 
Close #1
Exit Sub
cmd3_err:
MsgBox "Error " & error$, 48, "Aviso"
Exit Sub
End Sub
Sub envia_productos()
Dim buf As String
Dim xbuf As String
Dim xbuf1 As String
Dim xbuf2 As String
Dim xbuf3 As String
Dim xbuf4 As String
Dim xbuf5 As String
Dim xbuf6 As String
Dim xbuf7 As String
Dim xbuf8 As String
Dim xbuf9 As String
Dim xbuf10 As String
Dim xbuf11 As String
Dim xyz As String

Dim mytablex As Snapshot
Dim i As Integer
Dim xlinea As String
Dim xcount As Long
Dim vr
On Error GoTo cmd11_err

'----------------------------------------------------------
xbuf = globaldir + "\web\envio\prod"
Open xbuf For Output As #1
buf = "select * from producto  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
If Command2.Visible = False Then Exit Do
vr = DoEvents()
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1

'familia
xbuf = globaldir + "\web\envio\fami"
Open xbuf For Output As #1
buf = "select * from familia  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'subfamilia
xbuf = globaldir + "\web\envio\subfam"

Open xbuf For Output As #1
buf = "select * from subfamilia  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'seccion
xbuf = globaldir + "\web\envio\secc"

Open xbuf For Output As #1
buf = "select * from seccion  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'marca
xbuf = globaldir + "\web\envio\marca"
Open xbuf For Output As #1
buf = "select * from marca  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'categoria
xbuf = globaldir + "\web\envio\cate"
Open xbuf For Output As #1
buf = "select * from categoria  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'linea
xbuf = globaldir + "\web\envio\line"
Open xbuf For Output As #1
buf = "select * from linea  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'color
xbuf = globaldir + "\web\envio\color"
Open xbuf For Output As #1
buf = "select * from color  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'equiva
xbuf = globaldir + "\web\envio\equi"


Open xbuf For Output As #1
buf = "select * from productob  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'codprov
xbuf = globalweb + "\" + globalemp + "codp"
Open xbuf For Output As #1
buf = "select * from codprov  "
Set mytablex = mydbxglo.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
 
'ahora debe enviar los datos
xbuf = globaldir + "\web\envia\prod"
xbuf1 = globaldir + "\web\envia\subfam"
xbuf2 = globaldir + "\web\envia\secc"
xbuf3 = globaldir + "\web\envia\marca"
xbuf4 = globaldir + "\web\envia\cate"
xbuf5 = globaldir + "\web\envia\line"
xbuf6 = globaldir + "\web\envia\colo"
xbuf7 = globaldir + "\web\envia\equi"
xbuf8 = globaldir + "\web\envia\codp"
Call ExecuteCommand("\rp_orion.v2\001d\01\web\subirp.bat scanpos " & xbuf & " " & xbuf1 & " " & xbuf2 & "" & xbuf3 & " " & xbuf4 & " " & xbuf5 & " " & xbuf6 & "" & xbuf7 & " " & xbuf8)
'Call ExecuteCommand("\rp_orion.v2\001d\web\subir.bat scanpos " & xbuf3 & " " & xbuf4 & " " & xbuf5)
'Call ExecuteCommand("\rp_orion.v2\001d\web\subir.bat scanpos " & xbuf6 & " " & xbuf7 & " " & xbuf8)
'Call ExecuteCommand("\rp_orion.v2\001d\web\subir.bat scanpos " & xbuf9 & " " & xbuf10 & " " & xbuf11)
MsgBox "proceso Terminado", 48, "Aviso"
Exit Sub
cmd11_err:
MsgBox "Error envia Productos " & error$, 48, "Aviso"
Exit Sub
End Sub

Sub envia_formato_orion3()
Dim buf As String
Dim xbuf As String
Dim xbuf1 As String
Dim xbuf2 As String
Dim xbuf3 As String
Dim xbuf4 As String
Dim xbuf5 As String
Dim xbuf6 As String
Dim xbuf7 As String
Dim xbuf8 As String
Dim xbuf9 As String
Dim xbuf10 As String
Dim xbuf11 As String
Dim xyz As String
Dim mydbx As Database
Dim mytablex As Snapshot
Dim i As Integer
Dim xlinea As String
Dim found As Integer
Dim xcount As Long
Dim vr
On Error GoTo cmd131_err
found = copia_productos()
If found = 1 Then
   'MsgBox "Puedes seguir"
End If
exporta_todo_orion3
'MsgBox "Hola"
globaldira = globalweb + "\temp1"
Set mydbx = OpenDatabase(globaldira, False, False, "foxpro 2.5;")
xbuf = globalweb + "\" + globalemp + "PROD"
Open xbuf For Output As #1
buf = "select * from producto  "
Set mytablex = mydbx.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'familia

xbuf = globalweb + "\" + globalemp + "FAMI"
Open xbuf For Output As #1
buf = "select * from familia  "
Set mytablex = mydbx.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'subfamilia
xbuf = globalweb + "\" + globalemp + "SUBF"

Open xbuf For Output As #1
buf = "select * from subfamilia  "
Set mytablex = mydbx.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'seccion
xbuf = globalweb + "\" + globalemp + "secc"

Open xbuf For Output As #1
buf = "select * from seccion  "
Set mytablex = mydbx.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'marca
xbuf = globalweb + "\" + globalemp + "marc"
Open xbuf For Output As #1
buf = "select * from marca  "
Set mytablex = mydbx.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'categoria
xbuf = globalweb + "\" + globalemp + "cate"

Open xbuf For Output As #1
buf = "select * from categoria  "
Set mytablex = mydbx.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'xxx----xxxx--xxxx
xbuf = globalweb + "\" + globalemp + "colo"
Open xbuf For Output As #1
buf = "select * from color  "
Set mytablex = mydbx.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
xlinea = ""
If mytablex.EOF Then Exit Do
registro = "" & xcount
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do
For i = 0 To mytablex.Fields.Count - 1
xlinea = xlinea + "" & mytablex.Fields(i) & "|"
Next i
xcount = xcount + 1
Print #1, xlinea
mytablex.MoveNext
Loop
mytablex.Close
Close #1
'ahora debe enviar los datos
xbuf = globalweb + "\" + globalemp + "PROD"
xbuf1 = globalweb + "\" + globalemp + "FAMI"
xbuf2 = globalweb + "\" + globalemp + "SUBF"
xbuf3 = globalweb + "\" + globalemp + "secc"
xbuf4 = globalweb + "\" + globalemp + "marc"
xbuf5 = globalweb + "\" + globalemp + "cate"
xbuf6 = globalweb + "\" + globalemp + ""
xbuf7 = globalweb + "\" + globalemp + ""
xbuf8 = globalweb + "\" + globalemp + ""
xbuf9 = globalweb + "\" + globalemp + "colo"
xbuf10 = globalweb + "\" + globalemp + ""
xbuf11 = globalweb + "\" + globalemp + ""
'MsgBox "PASE POR AQUI"

Call ExecuteCommand("\rp_orion.v2\001d\06\web\subirp.bat scanpos " & xbuf & " " & xbuf1 & " " & xbuf2 & " " & xbuf3 & " " & xbuf4 & " " & xbuf5 & " " & xbuf6 & " " & xbuf7 & " " & xbuf8 & " " & xbuf9 & " " & xbuf10 & " " & xbuf11)
'Call ExecuteCommand("\rp_orion.v2\001d\web\subir.bat scanpos " & xbuf3 & " " & xbuf4 & " " & xbuf5)
'Call ExecuteCommand("\rp_orion.v2\001d\web\subir.bat scanpos " & xbuf6 & " " & xbuf7 & " " & xbuf8)
'Call ExecuteCommand("\rp_orion.v2\001d\web\subir.bat scanpos " & xbuf9 & " " & xbuf10 & " " & xbuf11)
MsgBox "proceso Terminado", 48, "Aviso"
dlo789923_Click
Exit Sub
cmd131_err:
MsgBox "ErrorEnvia Formato " & error$, 48, "Aviso"
Exit Sub

End Sub
Sub exporta_producto_orion3()
Dim mytablex As Table
Dim mydby As Database
Dim mytabley As Snapshot
Dim mydbx As Database
Dim xcount As Integer
Dim vr
On Error GoTo cmd34_err
globaldira = globalweb + "\temp1"
Set mydbx = OpenDatabase(globaldira, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("producto")
Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
buf = "select * from producto  "
Set mytabley = mydby.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
If mytabley.EOF Then Exit Do
registro = "" & xcount
xcount = xcount + 1
vr = DoEvents()
If Command2.Visible = False Then Exit Do
If Command2.Visible = False Then Exit Do

    mytablex.AddNew
    mytablex.Fields("maxdscto") = 0
    mytablex.Fields("flagtalla") = ""
    mytablex.Fields("unidadp") = "" & mytabley.Fields("unidad")
    mytablex.Fields("factorp") = Val("" & mytabley.Fields("factor"))
    mytablex.Fields("descue1") = 0
    mytablex.Fields("descue2") = 0
    mytablex.Fields("descue3") = 0
    mytablex.Fields("descue4") = 0
    mytablex.Fields("descue5") = 0
    mytablex.Fields("descue6") = 0
    mytablex.Fields("descue7") = 0
    mytablex.Fields("descue8") = 0
    mytablex.Fields("descue9") = 0
    mytablex.Fields("descue10") = 0
    mytablex.Fields("fabrica") = ""
    mytablex.Fields("comv1") = 0
    mytablex.Fields("comv2") = 0
    mytablex.Fields("comv3") = 0
    mytablex.Fields("comv4") = 0
    mytablex.Fields("comv5") = 0
    mytablex.Fields("comv6") = 0
    mytablex.Fields("comv7") = 0
    mytablex.Fields("comv8") = 0
    mytablex.Fields("comv9") = 0
    mytablex.Fields("comv10") = 0
    mytablex.Fields("NROBARRA") = 0
    mytablex.Fields("stock") = 0
    mytablex.Fields("cpa") = ""
    mytablex.Fields("cpa1") = ""
    mytablex.Fields("percepcion") = 0
    mytablex.Fields("flagunidad") = ""
    mytablex.Fields("flagremate") = ""
    mytablex.Fields("grafico") = ""
    mytablex.Fields("nprevta") = "9"
    mytablex.Fields("flagcolpo") = ""
    mytablex.Fields("tipocfre") = ""
    'mytablex.Fields("cantcfre") = 0
    'mytablex.Fields("puntcfre") = 0
    mytablex.Fields("oferta") = ""
    mytablex.Fields("presentaci") = ""
    'mytablex.Fields("fechavence") = ""
    
    mytablex.Fields("LOCAL") = ""
    mytablex.Fields("balanza") = ""
    mytablex.Fields("monedac") = "" & mytabley.Fields("monedac")
    
    
    mytablex.Fields("serie") = ""
    mytablex.Fields("grupocf") = ""
    'mytablex.Fields("nodscto") = 0
    'mytablex.Fields("premio") = 0
    'mytablex.Fields("pventaca") = 0
    
    mytablex.Fields("producto") = "" & mytabley.Fields("producto")
    mytablex.Fields("excludscto") = ""
    mytablex.Fields("retencion") = 0
    mytablex.Fields("proveedor") = ""
    mytablex.Fields("proveedor1") = ""
    mytablex.Fields("comision") = Val("" & mytabley.Fields("comision"))
    mytablex.Fields("peso") = 0.0001
    
    mytablex.Fields("servicio") = ""
    mytablex.Fields("Barras") = "" & mytabley.Fields("barras")
    mytablex.Fields("Nac_imp") = "N"
    mytablex.Fields("Familia") = "" & mytabley.Fields("familia")
    mytablex.Fields("Subfamilia") = "" & mytabley.Fields("subfamilia")
    mytablex.Fields("Seccion") = "" & mytabley.Fields("seccion")
    mytablex.Fields("Categoria") = "" & mytabley.Fields("categoria")
    mytablex.Fields("Marca") = "" & mytabley.Fields("marca")
    mytablex.Fields("Descripcio") = "" & mytabley.Fields("descripcio")
    
    mytablex.Fields("Abreviado") = Mid$("" & mytabley.Fields("descorto"), 1, 20)
    mytablex.Fields("Presentaci") = Mid$("" & mytabley.Fields("presenta"), 1, 12)
    
    mytablex.Fields("Igv") = Val("" & mytabley.Fields("igv"))
    
    'mytablex.Fields("isc") = 0
    
    mytablex.Fields("Moneda") = "" & mytabley.Fields("monedav")
    'mytablex.Fields("Costoini") = 0
    
    mytablex.Fields("Costopaqu") = Val("" & mytabley.Fields("costou"))
    mytablex.Fields("Costopaqp") = Val("" & mytabley.Fields("costop"))
    'mytablex.Fields("P1min") = 0
    'mytablex.Fields("P1max") = 0
    'mytablex.Fields("Ppventa1") = 0
    'mytablex.Fields("P2min") = 0
    'mytablex.Fields("P2max") = 0
    'mytablex.Fields("Ppventa2") = 0
    'mytablex.Fields("P3min") = 0
    'mytablex.Fields("P3max") = 0
    'mytablex.Fields("Ppventa3") = 0
    'mytablex.Fields("P4min") = 0
    'mytablex.Fields("P4max") = 0
    'mytablex.Fields("Ppventa4") = 0
    'mytablex.Fields("P5min") = 0
    'mytablex.Fields("P5max") = 0
    'mytablex.Fields("Ppventa5") = 0
    mytablex.Fields("Unidad") = "" & mytabley.Fields("unidad")
    mytablex.Fields("Factor") = Val("" & mytabley.Fields("factor"))

    mytablex.Fields("Unidad1") = "" & mytabley.Fields("unidad1")
    mytablex.Fields("Factor1") = Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("Pventa1") = Val("" & mytabley.Fields("pventa1"))
    
    
    mytablex.Fields("estado") = "" & mytabley.Fields("estado")
    
    mytablex.Fields("Unidad2") = "" & mytabley.Fields("unidad2")
    mytablex.Fields("Factor2") = Val("" & mytabley.Fields("factor2"))
    mytablex.Fields("Pventa2") = Val("" & mytabley.Fields("pventa2"))
    
    mytablex.Fields("Unidad3") = "" & mytabley.Fields("unidad3")
    mytablex.Fields("Factor3") = Val("" & mytabley.Fields("factor3"))
    mytablex.Fields("Pventa3") = Val("" & mytabley.Fields("pventa3"))

    mytablex.Fields("Unidad4") = "" & mytabley.Fields("unidad4")
    mytablex.Fields("Factor4") = Val("" & mytabley.Fields("factor4"))
    mytablex.Fields("Pventa4") = Val("" & mytabley.Fields("pventa4"))
    
    mytablex.Fields("Unidad5") = "" & mytabley.Fields("unidad5")
    mytablex.Fields("Factor5") = Val("" & mytabley.Fields("factor5"))
    mytablex.Fields("Pventa5") = Val("" & mytabley.Fields("pventa5"))
    mytablex.Fields("Unidad6") = "" & mytabley.Fields("unidad6")
    mytablex.Fields("Factor6") = Val("" & mytabley.Fields("factor6"))
    mytablex.Fields("Pventa6") = Val("" & mytabley.Fields("pventa6"))
    
    mytablex.Fields("Unidad7") = "" & mytabley.Fields("unidad7")
    mytablex.Fields("Factor7") = Val("" & mytabley.Fields("factor7"))
    mytablex.Fields("Pventa7") = Val("" & mytabley.Fields("pventa7"))

    mytablex.Fields("Unidad8") = "" & mytabley.Fields("unidad8")
    mytablex.Fields("Factor8") = Val("" & mytabley.Fields("factor8"))
    mytablex.Fields("Pventa8") = Val("" & mytabley.Fields("pventa8"))
    
    mytablex.Fields("Unidad9") = "" & mytabley.Fields("unidad9")
    mytablex.Fields("Factor9") = Val("" & mytabley.Fields("factor9"))
    mytablex.Fields("Pventa9") = Val("" & mytabley.Fields("pventa9"))
    
    mytablex.Fields("margen1") = Val("" & mytabley.Fields("margen1"))
    mytablex.Fields("margen2") = Val("" & mytabley.Fields("margen2"))
    mytablex.Fields("margen3") = Val("" & mytabley.Fields("margen3"))
    mytablex.Fields("margen4") = Val("" & mytabley.Fields("margen4"))
    mytablex.Fields("margen5") = Val("" & mytabley.Fields("margen5"))
    mytablex.Fields("margen6") = Val("" & mytabley.Fields("margen6"))
    mytablex.Fields("margen7") = Val("" & mytabley.Fields("margen7"))
    mytablex.Fields("margen8") = Val("" & mytabley.Fields("margen8"))
    mytablex.Fields("margen9") = Val("" & mytabley.Fields("margen9"))
    mytablex.Fields("margen10") = Val("" & mytabley.Fields("margen10"))
    
    
    'mytablex.Fields("pmargen1") = Val("" & mytabley.Fields("pmargen1"))
    'mytablex.Fields("pmargen2") = Val("" & mytabley.Fields("pmargen2"))
    'mytablex.Fields("pmargen3") = Val("" & mytabley.Fields("pmargen3"))
    'mytablex.Fields("pmargen4") = Val("" & mytabley.Fields("pmargen4"))
    'mytablex.Fields("pmargen5") = Val("" & mytabley.Fields("pmargen5"))
    
    mytablex.Fields("pvdscto1") = Val("" & mytabley.Fields("pventa1"))
    mytablex.Fields("pvdscto2") = Val("" & mytabley.Fields("pventa2"))
    mytablex.Fields("pvdscto3") = Val("" & mytabley.Fields("pventa3"))
    mytablex.Fields("pvdscto4") = Val("" & mytabley.Fields("pventa4"))
    mytablex.Fields("pvdscto5") = Val("" & mytabley.Fields("pventa5"))
    mytablex.Fields("pvdscto6") = Val("" & mytabley.Fields("pventa6"))
    mytablex.Fields("pvdscto7") = Val("" & mytabley.Fields("pventa7"))
    mytablex.Fields("pvdscto8") = Val("" & mytabley.Fields("pventa8"))
    mytablex.Fields("pvdscto9") = Val("" & mytabley.Fields("pventa9"))
    mytablex.Fields("pvdscto10") = Val("" & mytabley.Fields("pventa10"))
    mytablex.Fields("estado") = "1"
    mytablex.Update
mytabley.MoveNext
Loop
mytablex.Close
mytabley.Close
mydbx.Close
mydby.Close

Exit Sub
cmd34_err:
MsgBox "Producto error " + error$, 48, "Aviso"
Exit Sub
End Sub

Sub exporta_familia_orion3()
Dim mytablex As Table
Dim mydby As Database
Dim mytabley As Snapshot
Dim mydbx As Database
Dim xcount As Integer
On Error GoTo cmd91_err
globaldira = globalweb + "\temp1"
Set mydbx = OpenDatabase(globaldira, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("familia")
Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
buf = "select * from familia  "
Set mytabley = mydby.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
If mytabley.EOF Then Exit Do
registro = "" & xcount
xcount = xcount + 1
vr = DoEvents()
If Command2.Visible = False Then Exit Do

    mytablex.AddNew
    mytablex.Fields("familia") = "" & mytabley.Fields("familia")
    mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
    mytablex.Update
mytabley.MoveNext
Loop
mytablex.Close
mytabley.Close
mydby.Close
mydbx.Close
 
Exit Sub
cmd91_err:
MsgBox "errorFamilia " + error$, 48, "Aviso"
Exit Sub


End Sub

Sub exporta_subfamilia_orion3()
Dim mytablex As Table
Dim mydby As Database
Dim mytabley As Snapshot
Dim mydbx As Database
Dim xcount As Integer
On Error GoTo cmd90_err
globaldira = globalweb + "\temp1"
Set mydbx = OpenDatabase(globaldira, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("subfamil")
Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
buf = "select * from subfamilia  "
Set mytabley = mydby.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
If mytabley.EOF Then Exit Do
registro = "" & xcount
xcount = xcount + 1
vr = DoEvents()
If Command2.Visible = False Then Exit Do

    mytablex.AddNew
    mytablex.Fields("subfamilia") = "" & mytabley.Fields("subfamilia")
    mytablex.Fields("familia") = "" & mytabley.Fields("familia")
    mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
    mytablex.Update
mytabley.MoveNext
Loop
mytablex.Close
mytabley.Close
mydby.Close
mydbx.Close
 
Exit Sub
cmd90_err:
MsgBox "errorSubFamilia " + error$, 48, "Aviso"
Exit Sub


End Sub

Sub exporta_seccion_orion3()
Dim mytablex As Table
Dim mydby As Database
Dim mytabley As Snapshot
Dim mydbx As Database
Dim xcount As Integer
On Error GoTo cmd80_err
globaldira = globalweb + "\temp1"
Set mydbx = OpenDatabase(globaldira, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("seccion")
Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
buf = "select * from seccion  "
Set mytabley = mydby.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
If mytabley.EOF Then Exit Do
registro = "" & xcount
xcount = xcount + 1
vr = DoEvents()
If Command2.Visible = False Then Exit Do

    mytablex.AddNew
    mytablex.Fields("seccion") = "" & mytabley.Fields("seccion")
    mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
    mytablex.Update
mytabley.MoveNext
Loop
mytablex.Close
mytabley.Close
mydbx.Close
mydby.Close
 
Exit Sub
cmd80_err:
MsgBox "Seccion " + error$, 48, "Aviso"
Exit Sub

End Sub
Sub exporta_categoria_orion3()
Dim mytablex As Table
Dim mydby As Database
Dim mytabley As Snapshot
Dim mydbx As Database
Dim xcount As Integer
On Error GoTo cmd78_err
globaldira = globalweb + "\temp1"
Set mydbx = OpenDatabase(globaldira, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("categori")
Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
buf = "select * from categoria  "
Set mytabley = mydby.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
If mytabley.EOF Then Exit Do
registro = "" & xcount
xcount = xcount + 1
vr = DoEvents()
If Command2.Visible = False Then Exit Do

    mytablex.AddNew
    mytablex.Fields("categoria") = "" & mytabley.Fields("categoria")
    mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
    mytablex.Update
mytabley.MoveNext
Loop
mytablex.Close
mytabley.Close
mydbx.Close
mydby.Close
 
Exit Sub
cmd78_err:
MsgBox "errorCategoria " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub exporta_marca_orion3()
Dim mytablex As Table
Dim mydby As Database
Dim mytabley As Snapshot
Dim mydbx As Database
Dim xdcount As Integer
On Error GoTo cmd79_err
globaldira = globalweb + "\temp1"
Set mydbx = OpenDatabase(globaldira, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("marca")
Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
buf = "select * from marca  "
Set mytabley = mydby.CreateSnapshot(buf)
xcount = 0
registro = ""
Do
If mytabley.EOF Then Exit Do
registro = "" & xcount
xcount = xcount + 1
vr = DoEvents()
If Command2.Visible = False Then Exit Do
    mytablex.AddNew
    mytablex.Fields("marca") = "" & mytabley.Fields("marca")
    mytablex.Fields("descripcio") = "" & mytabley.Fields("descripcio")
    mytablex.Update
mytabley.MoveNext
Loop
mytablex.Close
mytabley.Close
mydbx.Close
mydby.Close
 
Exit Sub
cmd79_err:
MsgBox "errorMarca " + error$, 48, "Aviso"
Exit Sub
End Sub

Sub exporta_todo_orion3()
exporta_producto_orion3
exporta_marca_orion3
exporta_categoria_orion3
exporta_seccion_orion3
exporta_subfamilia_orion3
exporta_familia_orion3

End Sub
Function copia_productos()
On Error GoTo cmd235_err
borrar_archivo globalweb & "\temp1\producto.dbf"
borrar_archivo globalweb & "\temp1\producto.cdx"
borrar_archivo globalweb & "\temp1\familia.dbf"
borrar_archivo globalweb & "\temp1\familia.cdx"
borrar_archivo globalweb & "\temp1\subfamil.dbf"
borrar_archivo globalweb & "\temp1\subfamil.cdx"
borrar_archivo globalweb & "\temp1\seccion.dbf"
borrar_archivo globalweb & "\temp1\seccion.cdx"
borrar_archivo globalweb & "\temp1\marca.dbf"
borrar_archivo globalweb & "\temp1\marca.cdx"
borrar_archivo globalweb & "\temp1\color.dbf"
borrar_archivo globalweb & "\temp1\color.cdx"
borrar_archivo globalweb & "\temp1\linea.dbf"
borrar_archivo globalweb & "\temp1\linea.cdx"

FileCopy globalweb & "\temp1\tempx\producto.dbf", globalweb & "\temp1\producto.dbf"
FileCopy globalweb & "\temp1\tempx\producto.cdx", globalweb & "\temp1\producto.cdx"
FileCopy globalweb & "\temp1\tempx\familia.dbf", globalweb & "\temp1\familia.dbf"
FileCopy globalweb & "\temp1\tempx\familia.cdx", globalweb & "\temp1\familia.cdx"
FileCopy globalweb & "\temp1\tempx\subfamil.dbf", globalweb & "\temp1\subfamil.dbf"
FileCopy globalweb & "\temp1\tempx\subfamil.cdx", globalweb & "\temp1\subfamil.cdx"
FileCopy globalweb & "\temp1\tempx\seccion.dbf", globalweb & "\temp1\seccion.dbf"
FileCopy globalweb & "\temp1\tempx\seccion.cdx", globalweb & "\temp1\seccion.cdx"
FileCopy globalweb & "\temp1\tempx\marca.dbf", globalweb & "\temp1\marca.dbf"
FileCopy globalweb & "\temp1\tempx\categori.cdx", globalweb & "\temp1\marca.cdx"
FileCopy globalweb & "\temp1\tempx\color.dbf", globalweb & "\temp1\color.dbf"
FileCopy globalweb & "\temp1\tempx\color.cdx", globalweb & "\temp1\color.cdx"
FileCopy globalweb & "\temp1\tempx\linea.dbf", globalweb & "\temp1\linea.dbf"
FileCopy globalweb & "\temp1\tempx\linea.cdx", globalweb & "\temp1\linea.cdx"
'FileCopy globalweb & "\temp1\tempx\proveedo.dbf", globalweb & "\temp\proveedo.dbf"
'FileCopy globalweb & "\temp1\tempx\proveedo.cdx", globalweb & "\temp\proveedo.cdx"
'FileCopy globalweb & "\temp1\tempx\clientes.dbf", globalweb & "\temp\clientes.dbf"
'FileCopy globalweb & "\temp1\tempx\clientes.cdx", globalweb & "\temp\clientes.cdx"
copia_productos = 1
Exit Function
cmd235_err:
MsgBox "Warning copia productos " + error$, 48, "Aviso"
Exit Function
End Function


