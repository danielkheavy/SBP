VERSION 5.00
Begin VB.Form repprodu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Produccion"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox veproducto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
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
      TabIndex        =   11
      Top             =   6120
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
      TabIndex        =   4
      Text            =   "45"
      Top             =   5760
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
      TabIndex        =   3
      Top             =   5400
      Width           =   3855
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
      TabIndex        =   2
      Top             =   840
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
      TabIndex        =   1
      Top             =   480
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
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tarjetas"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agrupacion"
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
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7560
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   255
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
      TabIndex        =   9
      Top             =   5760
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
      TabIndex        =   8
      Top             =   5400
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
      TabIndex        =   7
      Top             =   840
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
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label10 
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
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu eki 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu lso3232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repprodu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub eki_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    Dim mytablez As Table

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_cuentaxc(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_cuentaxc
    cuerpo_programa_cuentaxc mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub Form_Load()
    veproducto.AddItem "N"
    veproducto.AddItem "S"
    veproducto.ListIndex = 0
    Combo1.AddItem "Numero"
    Combo1.AddItem "Fecha"
    Combo1.AddItem "Vendedor"
    Combo1.ListIndex = 0
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub lso3232_Click()
    repprodu.Hide
    Unload repprodu

End Sub

Sub cabecera_cuentaxc()

    Dim buf   As String

    Dim I     As Integer

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
    
    buf = String(152, "_")
    found = formateaa(buf, 152, 2, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Responsable", 12, 0, 0)
    found = formateaa("Fechai", 11, 0, 0)
    found = formateaa("Fechaf", 11, 0, 0)
    found = formateaa("BMP", 4, 0, 0)
    found = formateaa("BPT", 4, 0, 0)
    found = formateaa("Observa", 20, 2, 0)
    
    buf = String(152, "_")
    found = formateaa(buf, 152, 2, 0)

End Sub

Sub cuerpo_programa_cuentaxc(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim tmp1  As String

    Dim sw    As Integer

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0

    Tmp = ""
    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do

        If Combo1 = "Fecha" Then
            tmp1 = "" & mytablex.Fields("Fecha")

        End If

        If Combo1 = "Numero" Then
            tmp1 = "" & mytablex.Fields("Numero")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("Vendedor")

        End If

        If sw = 0 Then
            If Combo1 = "Fecha" Then
                buf = "" & mytablex.Fields("Fecha")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("fecha")

            End If

            If Combo1 = "Numero" Then
                buf = "" & mytablex.Fields("Numero")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Numero")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("Vendedor")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Vendedor")

            End If
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        If Tmp <> tmp1 Then
            found = formateaa("SubGrupo ", 37, 0, 1)
            buf = Format(suma1, "00000")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
   
            If Combo1 = "Fecha" Then
                buf = "" & mytablex.Fields("Fecha")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("fecha")

            End If

            If Combo1 = "Numero" Then
                buf = "" & mytablex.Fields("Numero")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Numero")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("vendedor")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Fechai")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("bodegai")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("bodega")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("observa")
        found = formateaa(buf, 20, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas

        If veproducto = "S" Then  'visualiza los productos
            visualiza_productos mytablex

        End If
   
        suma1 = suma1 + 1
        ssuma1 = ssuma1 + 1
        mytablex.MoveNext
    Loop
    found = formateaa("Sub Grupo ", 37, 0, 1)
    buf = Format(suma1, "00000")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("Grupo ", 37, 0, 1)
    buf = Format(ssuma1, "00000")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
   
End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_cuentaxc

    End If

End Sub

Function busca_nombre(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("clientes")
    mytablex.Index = "codigo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_nombre = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_vendedor(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("vendedor")
    mytablex.Index = "codigo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_vendedor = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close
 
End Function

Function sql_cuentaxc(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    buf = "select * from cproducc where "
    buf = buf & "  fechai>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If IsNumeric(Numero) Then
        buf = buf & " and numero=" & Val(Numero) & ""

    End If

    buf = buf & " order by " & Combo1 & ""
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_cuentaxc = 1

End Function

Sub visualiza_productos(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim found    As Integer

    Dim sw       As Integer

    sw = 0
    suma2 = 0
    suma3 = 0
    mytablex.Open "select * from tarjetaproduccion where numero=" & Val("" & mytabley.Fields("numero")) & "", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If sw = 0 Then
            found = formateaa("", 12, 0, 0)
            buf = String(100, "*")
            found = formateaa(buf, 100, 2, 0)
            nlineas
            found = formateaa("", 12, 0, 0)
            buf = "Producto"
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "Descripcio"
            found = formateaa(buf, 60, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "Unid"
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "Fx"
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "Cant "
            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "Precio "
            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "Total "
            found = formateaa(buf, 9, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "Tarjeta "
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        found = formateaa("", 12, 0, 0)
        buf = "" & mytablex.Fields("Producto")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Descripcio")
        found = formateaa(buf, 60, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Unidad")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Factor")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Cantidad")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Costo")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("CostoTotal")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        buf = "" & mytablex.Fields("tarjeta")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        suma2 = suma2 + Val("" & mytablex.Fields("cantidad"))
        suma3 = suma3 + Val("" & mytablex.Fields("costoTotal"))
        sw = 1
        visualiza_secciones mytablex
        'If Len("" & mytablex.Fields("linea")) > 0 Then 'ahora las tallas
        '   imprime_lineas mytablex
        'End If
        mytablex.MoveNext
    Loop

    mytablex.Close
 
    If sw = 1 Then
        found = formateaa("", 12, 0, 1)
        found = formateaa("", 84, 0, 0)
        buf = Format(suma2, "0.00")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 8, 0, 0)
        buf = Format(suma3, "0.00")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas

        'buf = String(152, "")
        'found = formateaa(buf, 152, 2, 0)
        'nlineas
    End If

End Sub

Sub visualiza_secciones(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim found    As Integer

    Dim sw       As Integer

    sw = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    mytablex.Open "select * from seccionproduccion where tarjeta='" & Val("" & mytabley.Fields("tarjeta")) & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If sw = 0 Then
            found = formateaa("", 12, 0, 0)
            buf = String(100, "*")
            found = formateaa(buf, 100, 2, 0)
            nlineas
            found = formateaa("", 12, 0, 0)
            buf = "Seccion"
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "FechaI"
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "FechaF"
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "Respon"
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "Estado"
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "Material "
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "ManoObra "
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "Merma "
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        found = formateaa("", 12, 0, 0)
        buf = "" & mytablex.Fields("seccion")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fechai")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fechaf")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("vendedor")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Estado")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("CostoMaterial")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Costomanoobra")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Merma")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        suma2 = suma2 + Val("" & mytablex.Fields("costomaterial"))
        suma3 = suma3 + Val("" & mytablex.Fields("costomanoobra"))
        suma4 = suma4 + Val("" & mytablex.Fields("merma"))
        sw = 1
   
        mytablex.MoveNext
    Loop

    mytablex.Close
 
    If sw = 1 Then
        found = formateaa("", 12, 0, 1)
        found = formateaa("", 84, 0, 0)
        buf = Format(suma2, "0.00")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 8, 0, 0)
        buf = Format(suma3, "0.00")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas

        'buf = String(152, "")
        'found = formateaa(buf, 152, 2, 0)
        'nlineas
    End If

End Sub

Sub imprime_lineas(mytabley As Snapshot)

    Dim mytablex As Table

    Dim found    As Integer

    Dim buf      As String

    Set mytablex = mydbxglo.OpenTable("linea")
    mytablex.Index = "linea"
    mytablex.Seek "=", "" & mytabley.Fields("linea")

    If Not mytablex.NoMatch Then
        buf = "" & mytablex.Fields("t1")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t1")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("t2")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t2")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t3")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t3")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t4")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t4")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t5")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t5")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t6")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t6")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t7")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t7")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t8")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t8")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t9")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t9")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t10")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t10")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t11")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t11")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t12")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t12")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t13")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t13")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t14")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t14")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t15")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t15")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("t16")
        found = formateaa(buf, 3, 0, 1)
        found = formateaa(":", 1, 0, 0)
        buf = "" & mytabley.Fields("t16")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
   
    End If

    mytablex.Close

End Sub

