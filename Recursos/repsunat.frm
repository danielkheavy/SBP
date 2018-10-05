VERSION 5.00
Begin VB.Form repsunat 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Sunat - Inventario Permanente en Unidades Fisica"
   ClientHeight    =   6090
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox nrolineas 
      Height          =   495
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   19
      Text            =   "45"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   15
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ejecutar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   2055
   End
   Begin VB.ComboBox bodega 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox unidad 
      Height          =   495
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   11
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox descripcio 
      Height          =   495
      Left            =   2520
      MaxLength       =   60
      TabIndex        =   9
      Top             =   3720
      Width           =   5655
   End
   Begin VB.TextBox producto 
      Height          =   495
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox direccion 
      Height          =   495
      Left            =   2520
      MaxLength       =   60
      TabIndex        =   5
      Top             =   2760
      Width           =   5655
   End
   Begin VB.TextBox nombre 
      Height          =   495
      Left            =   2520
      MaxLength       =   60
      TabIndex        =   3
      Top             =   2280
      Width           =   5655
   End
   Begin VB.TextBox Ruc 
      Height          =   495
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo Unidad Medida"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo Existencia"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Direccion Establecimiento"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Razon Social"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ruc"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Menu fd44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repsunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim found    As Integer

    Dim mytablex As Snapshot

    Dim mytabley As Table

    Dim mytablez As Table

    Dim buf      As String

    If Not IsDate(fechai) Then
        MsgBox "Fecha Inicio no valido ", 48, "Aviso"
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        fechaf = ""
        fechaf.SetFocus
        MsgBox "Fecha Final no valido ", 48, "Aviso"
        Exit Sub

    End If

    Set mytabley = mydbxglo.OpenTable("detalle")
    Set mytablez = mydbxglo.OpenTable("saldoini")
    mytablez.Index = "saldoini"
    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        mytabley.Close
        mytablez.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_kardex
    cuerpo_programa_kardex mytablex, mytablez
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    mytabley.Close
    mytablez.Close
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub fd44_Click()
    repsunat.Hide
    Unload repsunat

End Sub

Private Sub Form_Load()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = "31/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    bodega.Clear
    Set mytablex = mydbxglo.OpenTable("bodega")
    bodega.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    bodega.ListIndex = 1
    mytablex.Close

End Sub

Function sql_producto(mytablex As Snapshot)

    Dim buf As String

    buf = "select * from producto where producto like '" & producto & "'"
    buf = buf & " order by familia,Subfamilia,descripcio"
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    sql_producto = 1

End Function

Sub cabecera_kardex()

    Dim buf      As String

    Dim I        As Integer

    Dim xperiodo As String

    Dim found    As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)

    Select Case Month(fechai)

        Case 1
            periodo = "ENERO"

        Case 2
            periodo = "FEBRERO"

        Case 3
            periodo = "MARZO"

        Case 4
            periodo = "ABRIL"

        Case 5
            periodo = "MAYO"

        Case 6
            periodo = "JUNIO"

        Case 7
            periodo = "JULIO"

        Case 8
            periodo = "AGOSTO"

        Case 9
            periodo = "SETIEMBRE"

        Case 10
            periodo = "OCTUBRE"

        Case 11
            periodo = "NOVIEMBRE"

        Case 12
            periodo = "DICIEMBRE"

    End Select
   
    found = formateaa("Periodo                   : " & periodo & " " & Year(fechai), 40, 2, 0)
    found = formateaa("Ruc                       :" & RUC, 40, 2, 0)
    found = formateaa("Denominacion Razon Social :" & nombre, 60, 2, 0)
    found = formateaa("Direccion                 :" & direccion, 60, 2, 0)
    found = formateaa("Codigo Existencia         :" & producto, 60, 2, 0)
    found = formateaa("Tipo                      :" & tipo, 15, 2, 0)
    found = formateaa("Descripcio                :" & descripcio, 60, 2, 0)
    found = formateaa("Codigo Unidad             :" & "UND", 60, 2, 0)
      
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Tipo", 10, 0, 0)
    found = formateaa("Serie", 6, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("T.Opera", 10, 0, 0)
    found = formateaa("Entradas ", 10, 0, 1)
    found = formateaa("Salidas ", 10, 0, 1)
    found = formateaa("SaldoFinal ", 10, 2, 1)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Sub cuerpo_programa_kardex(mytablex As Snapshot, mytablez As Table)

    Dim mytabley As Snapshot

    Dim sw       As Integer

    Dim temp     As String

    Dim buf      As String

    Dim sw1      As Integer

    Dim temp1    As String

    Dim buf1     As String

    Dim bufx     As String

    Dim saldoini As Double

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim xbuf     As String

    Dim xentrada As Double

    Dim xsalida  As Double

    Dim xsaldo   As Double

    Dim found    As Integer

    sdx1 = 0
    sdx2 = 0
    sw1 = 0
    Do

        If mytablex.EOF Then Exit Do
        'SALDO INICIAL
        saldoini = 0
        xentrada = 0
        xsalida = 0
        xsaldo = 0
        mytablez.Seek "=", local1, "" & mytablex.Fields("producto"), extra_loquesea(bodega), fechai

        If Not mytablez.NoMatch Then
            saldoini = Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))

        End If

        xentrada = saldoini
        xsalida = 0
        xsaldo = xsaldo + saldoini
        found = formateaa(fechai, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 9, 0, 0) 'tipo
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 5, 0, 0)  'serie
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 11, 0, 0)    'numero
        found = formateaa("", 1, 0, 0)
        found = formateaa("", 9, 0, 0)  'tipo operacion
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & xentrada, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & xsalida, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & xsaldo, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        '-------ahora las transacciones------------
        buf = "select * from detalle where "
        buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
        buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
        buf = buf & " and (acu='S' or acu='T' or acu='C' or acu='A' or acu='B' or acu='C' or acu='D'  or acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='E' or acu='N')"
        buf = buf & " and estado='2'"
        buf = buf & " order by fecha,hora"
        Set mytabley = mydbxglo.CreateSnapshot(buf)
        Do

            If mytabley.EOF Then Exit Do
            found = formateaa("" & mytabley.Fields("fecha"), 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa(busca_tipo("" & mytabley.Fields("tipo")), 9, 0, 0) 'tipo
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("serie"), 5, 0, 0)  'serie
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & mytabley.Fields("numero"), 11, 0, 0)    'numero
            found = formateaa("", 1, 0, 0)
            found = formateaa(busca_tipo("" & mytabley.Fields("tipo")), 9, 0, 0)  'tipo operacion
            found = formateaa("", 1, 0, 0)
            xentrada = 0
            xsalida = 0
   
            If "" & mytabley.Fields("acu") = "T" Or "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytabley.Fields("acu") = "D" Or "" & mytabley.Fields("acu") = "G" Or "" & mytabley.Fields("acu") = "N" Then
                xsalida = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                xsaldo = xsaldo - xsalida

            End If

            If "" & mytabley.Fields("acu") = "S" Or "" & mytabley.Fields("acu") = "J" Or "" & mytabley.Fields("acu") = "K" Or "" & mytabley.Fields("acu") = "L" Or "" & mytabley.Fields("acu") = "M" Or "" & mytabley.Fields("acu") = "P" Or "" & mytabley.Fields("acu") = "E" Then
                xentrada = Val("" & mytabley.Fields("cantidad")) * Val("" & mytabley.Fields("factor"))
                xsaldo = xsaldo + xentrada

            End If

            found = formateaa("" & xentrada, 9, 0, 1)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & xsalida, 9, 0, 1)
            found = formateaa("", 1, 0, 0)
            found = formateaa("" & xsaldo, 9, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas
            mytabley.MoveNext
        Loop
        mytablex.MoveNext
    Loop

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        If opcion2 = "1" Then
            cabecera_kardex

        End If

    End If

End Sub

Function busca_tipo(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf

    If mytablex.NoMatch Then
        busca_tipo = "09"

    End If

    If Not mytablex.NoMatch Then

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "C", "J", "L"
                busca_tipo = "03"

            Case "B", "D", "K", "M"
                busca_tipo = "03"

            Case "S", "T"
                busca_tipo = "09"

        End Select
   
    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_tipo1(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D"
                busca_tipo1 = "01"

            Case "J", "K", "L", "M"
                busca_tipo1 = "02"

        End Select
   
    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

