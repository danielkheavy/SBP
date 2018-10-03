VERSION 5.00
Begin VB.Form rptcaba 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caja Bancos"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox titulo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   8
      Text            =   "Libro Caja Bancos"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox nrolineas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   7
      Text            =   "45"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox origen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo reporte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label periodo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Origen"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu dju3453 
      Caption         =   "&Buscar"
   End
   Begin VB.Menu ldo434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "rptcaba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function sql_documento(mydbx As Database, mytablex As Snapshot)
Dim buf As String
Dim buf2 As String
If Len(fechai) <> 10 Then Exit Function
If Len(fechaf) <> 10 Then Exit Function
If Not IsDate(fechai) Then Exit Function
If Not IsDate(fechaf) Then Exit Function
buf = "select * from mdh_vou where mes='" & Trim(periodo) & "'"
buf = buf & "  and fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
buf = buf & " and cuenta like '1*'"
If origen <> "*" Then
   buf2 = Trim(extra_loquesea("" & origen))
   buf = buf & " and t like'" & buf2 & "'"
End If
buf = buf & "order by Fecha,t,val(vou)"
Set mytablex = mydbx.CreateSnapshot(buf)
sql_documento = 1
End Function
Sub cabecera_documento()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    'cabecera_tipico "" & menuipos!nempresa, "", "" & "" & usuariopos
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("            Periodo Fechai : " & fechai & " al " & fechaf, 60, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("----Comprobante------", 22, 0, 0)
    found = formateaa("------------Cuenta-Subcuenta-------------", 42, 0, 0)
    found = formateaa("---------Movimiento--", 22, 2, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("     ", 6, 0, 0)
    found = formateaa("Or", 5, 0, 0)
    found = formateaa("Cuenta", 11, 0, 0)
    found = formateaa("Descripcio", 31, 0, 0)
    found = formateaa("Debe ", 11, 0, 1)
    found = formateaa("Haber ", 11, 0, 1)
    found = formateaa("Saldo ", 11, 0, 1)
    found = formateaa("Glosa ", 11, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    

End Sub
Sub cuerpo_programa_documento(mydbx As Database, mytablex As Snapshot)
Dim tmp As String
Dim sw As Integer
Dim buf As String
Dim found As Integer
Dim sdx As Double
Dim tmp1 As String
Dim xsaldo As Double
sdx = 0
sw = 0
suma1 = 0
suma2 = 0
ssuma1 = 0
ssuma2 = 0
xsaldo = 0
Do
If mytablex.EOF Then Exit Do
If sw = 0 Then
   buf = "" & mytablex.Fields("Fecha")
   found = formateaa(buf, 10, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" '& mytablex.Fields("vou")
   found = formateaa(buf, 5, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("t")
   found = formateaa(buf, 4, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = busca_origen("" & mytablex.Fields("t"))
   found = formateaa(buf, 20, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   sw = 1
   suma1 = 0
   suma2 = 0
   suma3 = 0
   tmp = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("t")
End If
If tmp <> "" & mytablex.Fields("fecha") & "" & mytablex.Fields("t") Then
   found = formateaa("----------------------", 85, 2, 1)
   nlineas
   found = formateaa("Total Comprobante-----------> ", 64, 0, 1)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas


   buf = "" & mytablex.Fields("Fecha")
   found = formateaa(buf, 10, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" '& mytablex.Fields("vou")
   found = formateaa(buf, 5, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("t")
   found = formateaa(buf, 4, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = busca_origen("" & mytablex.Fields("t"))
   found = formateaa(buf, 20, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   suma1 = 0
   suma2 = 0
   tmp = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("t")
End If
   found = formateaa("", 16, 0, 0)
   buf = "" & mytablex.Fields("vou")
   found = formateaa(buf, 5, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("cuenta")
   found = formateaa(buf, 10, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("Ncuenta")
   found = formateaa(buf, 30, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("debe")
   buf = Format(buf, "0.00")
   If Val(buf) = 0 Then
      buf = ""
   End If
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("haber")
   buf = Format(buf, "0.00")
   If Val(buf) = 0 Then
      buf = ""
   End If
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   'saldo
   If Val("" & mytablex.Fields("debe")) > 0 Then
   xsaldo = xsaldo + Val("" & mytablex.Fields("debe"))
   End If
   If Val("" & mytablex.Fields("haber")) > 0 Then
   xsaldo = xsaldo - Val("" & mytablex.Fields("debe"))
   End If
   buf = Format(xsaldo, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
      
   buf = "" & mytablex.Fields("glosa")
   found = formateaa(buf, 20, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   suma1 = suma1 + Val("" & mytablex.Fields("debe"))
   suma2 = suma2 + Val("" & mytablex.Fields("haber"))
   ssuma1 = ssuma1 + Val("" & mytablex.Fields("debe"))
   ssuma2 = ssuma2 + Val("" & mytablex.Fields("haber"))
mytablex.MoveNext
Loop
found = formateaa("----------------------", 85, 2, 1)
   nlineas
   found = formateaa("Total Comprobante-----------> ", 64, 0, 1)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
found = formateaa("======================", 85, 2, 1)
   nlineas
   found = formateaa("Total a la Fecha-----------> ", 64, 0, 1)
   buf = Format(ssuma1, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(ssuma2, "0.00")
   found = formateaa(buf, 10, 0, 1)
   found = formateaa("", 1, 2, 0)
   
End Sub

Private Sub dju3453_Click()
Dim found As Integer
Dim mydbx As Database
Dim mydby As Database
Dim mytablex As Snapshot
Dim mytabley As Table
Dim mytablez As Table
Dim buf As String
contlin = 0
suma1 = 0
suma2 = 0
suma3 = 0
ssuma1 = 0
ssuma2 = 0
ssuma3 = 0
Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
found = sql_documento(mydbx, mytablex)
If found = 0 Then
   mytablex.Close
   mydbx.Close
   Exit Sub
End If
    Filename = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    borra_nombre Filename
    Open Filename For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento mydbx, mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    mydbx.Close
    genver.Show 1

End Sub

Private Sub Form_Load()
Dim found As Integer
carga_como
found = busca_parame()
fechai = "01/" & Mid$(periodo, 1, 2) & "/" & Mid$(periodo, 3, 4)
fechaf = Format(Day(Now), "00") & "/" & Mid$(periodo, 1, 2) & "/" & Mid$(periodo, 3, 4)
End Sub
Sub carga_como()
Dim mydbx As Database
Dim mytablex As Table
origen.Clear
origen.AddItem "*"
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("origen")
Do
If mytablex.EOF Then Exit Do
origen.AddItem "" & mytablex.Fields("origen") & "|" & "" & mytablex.Fields("descripcio")
mytablex.MoveNext
Loop
mytablex.Close
mydbx.Close
origen.ListIndex = 0

End Sub
Function busca_parame()
Dim sdx As Double
Dim mydbx As Database
Dim mytablex As Table
periodo = ""
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   periodo = "" & mytablex.Fields("mesconta") & "" & mytablex.Fields("anoconta")
   busca_parame = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function



Private Sub ldo434_Click()
rptcaba.Hide
Unload rptcaba
End Sub
Sub nlineas()
    contlin = contlin + 1
    If contlin > Val(nrolineas) Then
       cabecera_documento
    End If
End Sub
Function extra_loquesea(buf As String) As String
Dim j
Dim buf1 As String
buf1 = ""
If InStr(buf, "|") > 0 Then
j = InStr(buf, "|")
   buf1 = Mid$(buf, 1, j - 1)
End If
extra_loquesea = buf1
End Function
Function busca_origen(buf As String) As String
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("origen")
mytablex.Index = "origen"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_origen = "" & mytablex.Fields("descripcio")
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function


