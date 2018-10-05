VERSION 5.00
Begin VB.Form repvouch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Vouchers"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox mes 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox observa 
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   9
      Text            =   "*"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox numero 
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   8
      Text            =   "*"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox asiento 
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
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "*"
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox tipo 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox moneda 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox voucher 
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   0
      Text            =   "*"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox fechai 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3600
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "45"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes"
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
      TabIndex        =   21
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
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
      TabIndex        =   19
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
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
      TabIndex        =   18
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
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
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
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
      TabIndex        =   16
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asiento"
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
      TabIndex        =   15
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Voucher"
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
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
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
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label4 
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
      TabIndex        =   10
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Menu fdskejer 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu lo324m1 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repvouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fdskejer_Click()
Dim found As Integer


Dim mytablex As Snapshot
Dim mytableY As Table
Dim buf As String
contlin = 0
suma1 = 0
suma2 = 0
suma3 = 0
ssuma1 = 0
ssuma2 = 0
ssuma3 = 0


found = sql_voucher(mytablex)
If found = 0 Then
    
   Exit Sub
End If
    Set mytableY = mydbxglo.OpenTable("dvoucher")
    mytableY.Index = "dvoucher"

    Filename = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & Filename)
    Open Filename For Append As #1
    '------------------------------------
    cabecera_voucher
    cuerpo_programa_voucher mytablex, mytableY
    '------------------------------------
    Close #1
    cerrar_archivo
    mytableY.Close
    mytablex.Close
     
    genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub Form_Load()
Dim i As Integer
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
moneda.Clear
moneda.AddItem "%"
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0
tipo.Clear
tipo.AddItem "%"
tipo.AddItem "C"
tipo.AddItem "D"
tipo.AddItem "A"
tipo.ListIndex = 0

mes.Clear
mes.AddItem "%"
For i = 1 To 12
    mes.AddItem Format(i, "00")
Next i
mes.ListIndex = 0
End Sub

Private Sub lo324m1_Click()
repvouch.Hide
Unload repvouch
End Sub
Function sql_voucher(mytablex As Snapshot)
Dim buf As String
If Len(fechai) <> 10 Then Exit Function
If Len(fechaf) <> 10 Then Exit Function
If Not IsDate(fechai) Then Exit Function
If Not IsDate(fechaf) Then Exit Function
buf = "select * from cvoucher where "
buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
If voucher <> "%" Then
buf = buf & " and numero like '" & voucher & "'"
End If
If asiento <> "%" Then
buf = buf & " and asiento like '" & asiento & "'"
End If
If tipo <> "%" Then
buf = buf & " and tipo like '" & tipo & "'"
End If
If mes <> "%" Then
buf = buf & " and mes like '" & mes & "'"
End If
If moneda <> "%" Then
buf = buf & " and moneda like '" & moneda & "'"
End If
If numero <> "%" Then
buf = buf & " and numero1 like '" & numero & "'"
End If
If observa <> "%" Then
buf = buf & " and concepto like '" & observa & "'"
End If
buf = buf & " order by asiento,str(numero),fecha"
Set mytablex = mydbxglo.CreateSnapshot(buf)
sql_voucher = 1

End Function
Sub cabecera_voucher()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Dia", 4, 0, 0)
    found = formateaa("Cuenta", 9, 0, 0)
    found = formateaa("Relacion", 12, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Debe ", 11, 0, 1)
    found = formateaa("Haber ", 11, 0, 1)
    found = formateaa("Observa", 7, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub
Sub cuerpo_programa_voucher(mytablex As Snapshot, mytableY As Table)
Dim buf As String
Dim found As Integer
Do
   If mytablex.EOF Then Exit Do
   found = formateaa("Asiento:", 9, 0, 0)
   buf = "" & mytablex.Fields("asiento")
   found = formateaa(buf, 2, 0, 0)
   found = formateaa("", 1, 0, 0)
   found = formateaa("Voucher:", 9, 0, 0)
   buf = "" & mytablex.Fields("numero")
   found = formateaa(buf, 11, 0, 0)
   found = formateaa("", 1, 0, 0)
   found = formateaa("Fecha:", 7, 0, 0)
   buf = "" & mytablex.Fields("fecha")
   found = formateaa(buf, 10, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   '---- buscar las cuentas
   suma1 = 0
   suma2 = 0
   mytableY.Seek "=", "" & mytablex.Fields("numero")
   If Not mytableY.NoMatch Then
      imprime_detalle mytablex, mytableY
   End If
'------------------------------------- ------------
   mytablex.MoveNext
Loop
End Sub
Sub nlineas()
    contlin = contlin + 1
    If contlin > Val(nrolineas) Then
       cabecera_voucher
    End If
End Sub
Sub imprime_detalle(mytablex As Snapshot, mytableY As Table)
Dim buf As String
Dim found As Integer
Do
      If mytableY.EOF Then Exit Do
      If "" & mytableY.Fields("numero") = "" & mytablex.Fields("numero") Then
         '------------------------------------------
         buf = "" & mytableY.Fields("dia")
         found = formateaa(buf, 3, 0, 0)
         found = formateaa("", 1, 0, 0)
         buf = "" & mytableY.Fields("cuenta")
         found = formateaa(buf, 8, 0, 0)
         found = formateaa("", 1, 0, 0)
         buf = busca_relacion("" & mytableY.Fields("cuenta"))
         found = formateaa(buf, 11, 0, 0)
         found = formateaa("", 1, 0, 0)
         buf = "" & mytablex.Fields("numero1")
         found = formateaa(buf, 11, 0, 0)
         found = formateaa("", 1, 0, 0)
         buf = "" & mytablex.Fields("fecha")
         found = formateaa(buf, 10, 0, 0)
         found = formateaa("", 1, 0, 0)
         buf = "" & mytableY.Fields("debe")
         buf = Format(Val(buf), "0.00")
         If Val(buf) = 0 Then
            buf = ""
         End If
         found = formateaa(buf, 10, 0, 1)
         found = formateaa("", 1, 0, 0)
         suma1 = suma1 + Val("" & mytableY.Fields("debe"))
         buf = "" & mytableY.Fields("haber")
         buf = Format(Val(buf), "0.00")
         If Val(buf) = 0 Then
            buf = ""
         End If
         found = formateaa(buf, 10, 0, 1)
         found = formateaa("", 1, 0, 0)
         suma2 = suma2 + Val("" & mytableY.Fields("haber"))
         buf = "" & mytableY.Fields("observa")
         found = formateaa(buf, 30, 0, 0)
         found = formateaa("", 1, 2, 0)
         nlineas
         '------------------------------------------
         Else
         Exit Do
      End If
      mytableY.MoveNext
      Loop
         found = formateaa("", 48, 0, 0)
         buf = Format(suma1, "0.00")
         found = formateaa(buf, 10, 0, 1)
         found = formateaa("", 1, 0, 0)
         
         buf = Format(suma1, "0.00")
         found = formateaa(buf, 10, 0, 1)
         found = formateaa("", 1, 2, 0)
         nlineas
End Sub
Function busca_relacion(buf As String) As String

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("cont001")
mytablex.Index = "cont001"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_relacion = "" & mytablex.Fields("relacion")
End If
'------------------------------------- ------------
mytablex.Close
 

End Function


