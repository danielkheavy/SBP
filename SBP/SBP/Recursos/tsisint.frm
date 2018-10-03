VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tsisint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formato Sistcont Prev"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   17115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Generar"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   6615
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   11668
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previo"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label nvoucher 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu dski44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsisint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mytabley As New ADODB.Recordset

Private Sub Command1_Click()
    generar
    carga_bd

End Sub

Private Sub Command2_Click()

    Dim buf   As String

    Dim found As Integer

    found = borra_nombre(globaldir & "\siscont\ventas.txt")
    Open globaldir & "\siscont\ventas.txt" For Append As #1
      
    Do

        If mytabley.EOF Then Exit Do
        buf = ""
        buf = buf & ("" & mytabley.Fields("origen"))
        buf = buf & ("" & mytabley.Fields("voucher"))
        buf = buf & ("" & mytabley.Fields("fecha"))
        buf = buf & ("" & mytabley.Fields("cuenta"))
        buf = buf & ("" & mytabley.Fields("monto"))
        buf = buf & ("" & mytabley.Fields("dh"))
        buf = buf & ("" & mytabley.Fields("moneda"))
        buf = buf & ("" & mytabley.Fields("paridad"))
        buf = buf & ("" & mytabley.Fields("tipodoc"))
        buf = buf & ("" & mytabley.Fields("numero"))
        buf = buf & ("" & mytabley.Fields("fechav"))
        buf = buf & poner_fijo("" & mytabley.Fields("codigo"), 11)
        buf = buf & poner_fijo("" & mytabley.Fields("ccosto"), 10)
        buf = buf & poner_fijo("" & mytabley.Fields("flujo"), 4)
        buf = buf & poner_fijo("" & mytabley.Fields("presupuesto"), 10)
        buf = buf & ("" & mytabley.Fields("tipolibro"))
        buf = buf & poner_fijo("" & mytabley.Fields("fechadoc"), 8)
        buf = buf & poner_fijo("" & mytabley.Fields("neto1"), 12)
        buf = buf & poner_fijo("" & mytabley.Fields("neto2"), 12)
        buf = buf & poner_fijo("" & mytabley.Fields("neto3"), 12)
        buf = buf & poner_fijo("" & mytabley.Fields("neto4"), 12)
        buf = buf & poner_fijo("" & mytabley.Fields("igv"), 12)
        buf = buf & poner_fijo("" & mytabley.Fields("ruc"), 11)
        buf = buf & ("" & mytabley.Fields("tipo"))
        buf = buf & poner_fijo("" & mytabley.Fields("razonsocial"), 40)
        buf = buf & poner_fijo("" & mytabley.Fields("glosa"), 30)
        buf = buf & ("" & mytabley.Fields("tipoidentidad"))
        buf = buf & poner_fijo("" & mytabley.Fields("mediopago"), 3)
        buf = buf & poner_fijo("" & mytabley.Fields("apellido1"), 20)
        buf = buf & poner_fijo("" & mytabley.Fields("apellido2"), 20)
        buf = buf & poner_fijo("" & mytabley.Fields("nombre"), 20)
        buf = buf & poner_fijo("" & mytabley.Fields("neto5"), 12)
        buf = buf & poner_fijo("" & mytabley.Fields("neto6"), 12)
        buf = buf & poner_fijo("" & mytabley.Fields("refnumero"), 20)
        buf = buf & poner_fijo("" & mytabley.Fields("reftipodoc"), 2)
        buf = buf & poner_fijo("" & mytabley.Fields("reffecha"), 8)
        buf = buf & poner_fijo("" & mytabley.Fields("detranumero"), 20)
        buf = buf & poner_fijo("" & mytabley.Fields("detrafecha"), 8)
        Print #1, buf
        mytabley.MoveNext
    Loop
    Close #1
    MsgBox "ACABE..."

End Sub

Private Sub dski44_Click()
    tsisint.Hide
    Unload tsisint

End Sub

Private Sub Form_Load()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_bd

End Sub

Sub carga_bd()

    If mytabley.State = 1 Then
        mytabley.Close

    End If

    Set mytabley = Nothing
    mytabley.Open "select * from  siscontin order by str(voucher)", cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytabley

End Sub

Sub generar()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim xvoucher As Integer

    Dim vr

    cn.Execute ("delete from siscontin")
    xvoucher = 1
    xfecha = fechai
    buf = "select * from factura where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
    buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D')"
    buf = buf & " order by fecha,tipo,serie,str(numero)"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()
        xvoucher = xvoucher + 1
        nvoucher = "" & xvoucher
        mytabley.AddNew
        mytabley.Fields("origen") = "02"
        mytabley.Fields("voucher") = Format(xvoucher, "0000")
        mytabley.Fields("fecha") = Format(xfecha, "dd/mm/yy")

        Select Case "" & mytablex.Fields("acu")

            Case "A", "C"
                mytabley.Fields("cuenta") = "12123"

            Case "B", "D"
                mytabley.Fields("cuenta") = "12121"

        End Select

        mytabley.Fields("monto") = Format(Val("" & mytablex.Fields("total")), "000000000.00")
        mytabley.Fields("dh") = "D"
        mytabley.Fields("moneda") = "S"
        mytabley.Fields("paridad") = Format(Val("" & mytablex.Fields("paridad")), "000000.000")

        Select Case "" & mytablex.Fields("acu")

            Case "A", "C"
                mytabley.Fields("tipodoc") = "03"

            Case "B", "D"
                mytabley.Fields("tipodoc") = "01"

        End Select

        mytabley.Fields("numero") = Trim("" & mytablex.Fields("serie")) & "-" & Trim("" & mytablex.Fields("numero")) 'ojo
        mytabley.Fields("fechav") = Format(Trim("" & mytablex.Fields("fecha")), "dd/mm/yy")
        mytabley.Fields("codigo") = Trim("" & mytablex.Fields("codigo"))
        mytabley.Fields("ccosto") = ""
        mytabley.Fields("flujo") = ""
        mytabley.Fields("presupuesto") = ""
        mytabley.Fields("tipolibro") = "V"
        mytabley.Fields("fecha") = Format(Trim("" & mytablex.Fields("fecha")), "dd/mm/yy")
        mytabley.Fields("neto1") = "" '& Format(Val("" & mytablex.Fields("subtotal")), "000000000.00")
        mytabley.Fields("neto2") = ""
        mytabley.Fields("neto3") = ""
        mytabley.Fields("neto4") = ""
        mytabley.Fields("igv") = "" 'Format(Val("" & mytablex.Fields("impuesto")), "000000000.00")
        mytabley.Fields("ruc") = poner_fijo(Trim("" & mytablex.Fields("codigo")), 11)
        mytabley.Fields("tipo") = "C"
        mytabley.Fields("razonsocial") = Mid$(Trim("" & mytablex.Fields("nombre")), 1, 40)
        mytabley.Fields("glosa") = Mid$(Trim("" & mytablex.Fields("observa")), 1, 30)
        mytabley.Fields("tipoidentidad") = "1"
        mytabley.Fields("mediopago") = "001"
        mytabley.Fields("apellido1") = ""
        mytabley.Fields("apellido2") = ""
        mytabley.Fields("nombre") = ""
        mytabley.Fields("neto5") = ""
        mytabley.Fields("neto6") = ""
        mytabley.Fields("refnumero") = ""
        mytabley.Fields("reftipodoc") = ""
        mytabley.Fields("reffecha") = ""
        mytabley.Fields("detranumero") = ""
        mytabley.Fields("detrafecha") = ""
        mytabley.Update

        mytabley.AddNew
        mytabley.Fields("origen") = "02"
        mytabley.Fields("voucher") = Format(xvoucher, "0000")
        mytabley.Fields("fecha") = Format(xfecha, "dd/mm/yy")

        Select Case "" & mytablex.Fields("acu")

            Case "A", "C"
                mytabley.Fields("cuenta") = "40111"

            Case "B", "D"
                mytabley.Fields("cuenta") = "40111"

        End Select

        mytabley.Fields("monto") = Format(Val("" & mytablex.Fields("impuesto")), "000000000.00")
        mytabley.Fields("dh") = "H"
        mytabley.Fields("moneda") = "S"
        mytabley.Fields("paridad") = Format(Val("" & mytablex.Fields("paridad")), "000000.000")

        Select Case "" & mytablex.Fields("acu")

            Case "A", "C"
                mytabley.Fields("tipodoc") = "03"

            Case "B", "D"
                mytabley.Fields("tipodoc") = "01"

        End Select

        mytabley.Fields("numero") = Trim("" & mytablex.Fields("serie")) & "-" & Trim("" & mytablex.Fields("numero")) 'ojo
        mytabley.Fields("fechav") = Format(Trim("" & mytablex.Fields("fecha")), "dd/mm/yy")
        mytabley.Fields("codigo") = Trim("" & mytablex.Fields("codigo"))
        mytabley.Fields("ccosto") = ""
        mytabley.Fields("flujo") = ""
        mytabley.Fields("presupuesto") = ""
        mytabley.Fields("tipolibro") = "V"
        mytabley.Fields("fecha") = Format(Trim("" & mytablex.Fields("fecha")), "dd/mm/yy")
        mytabley.Fields("neto1") = "" & Format(Val("" & mytablex.Fields("subtotal")), "000000000.00")
        mytabley.Fields("neto2") = ""
        mytabley.Fields("neto3") = ""
        mytabley.Fields("neto4") = ""
        mytabley.Fields("igv") = Format(Val("" & mytablex.Fields("impuesto")), "000000000.00")
        mytabley.Fields("ruc") = Trim("" & mytablex.Fields("codigo"))
        mytabley.Fields("tipo") = "C"
        mytabley.Fields("razonsocial") = Mid$(Trim("" & mytablex.Fields("nombre")), 1, 40)
        mytabley.Fields("glosa") = Mid$(Trim("" & mytablex.Fields("observa")), 1, 30)
        mytabley.Fields("tipoidentidad") = "1"
        mytabley.Fields("mediopago") = "001"
        mytabley.Fields("apellido1") = ""
        mytabley.Fields("apellido2") = ""
        mytabley.Fields("nombre") = ""
        mytabley.Fields("neto5") = ""
        mytabley.Fields("neto6") = ""
        mytabley.Fields("refnumero") = ""
        mytabley.Fields("reftipodoc") = ""
        mytabley.Fields("reffecha") = ""
        mytabley.Fields("detranumero") = ""
        mytabley.Fields("detrafecha") = ""
        mytabley.Update

        mytabley.AddNew
        mytabley.Fields("origen") = "02"
        mytabley.Fields("voucher") = Format(xvoucher, "0000")
        mytabley.Fields("fecha") = Format(xfecha, "dd/mm/yy")

        Select Case "" & mytablex.Fields("acu")

            Case "A", "C"
                mytabley.Fields("cuenta") = "70211"

            Case "B", "D"
                mytabley.Fields("cuenta") = "70211"

        End Select

        mytabley.Fields("monto") = Format(Val("" & mytablex.Fields("subtotal")), "000000000.00")
        mytabley.Fields("dh") = "H"
        mytabley.Fields("moneda") = "S"
        mytabley.Fields("paridad") = Format(Val("" & mytablex.Fields("paridad")), "000000.000")

        Select Case "" & mytablex.Fields("acu")

            Case "A", "C"
                mytabley.Fields("tipodoc") = "03"

            Case "B", "D"
                mytabley.Fields("tipodoc") = "01"

        End Select

        mytabley.Fields("numero") = Trim("" & mytablex.Fields("serie")) & "-" & Trim("" & mytablex.Fields("numero")) 'ojo
        mytabley.Fields("fechav") = Format(Trim("" & mytablex.Fields("fecha")), "dd/mm/yy")
        mytabley.Fields("codigo") = Trim("" & mytablex.Fields("codigo"))
        mytabley.Fields("ccosto") = ""
        mytabley.Fields("flujo") = ""
        mytabley.Fields("presupuesto") = ""
        mytabley.Fields("tipolibro") = "V"
        mytabley.Fields("fecha") = Format(Trim("" & mytablex.Fields("fecha")), "dd/mm/yy")
        mytabley.Fields("neto1") = "" '& Format(Val("" & mytablex.Fields("subtotal")), "000000000.00")
        mytabley.Fields("neto2") = ""
        mytabley.Fields("neto3") = ""
        mytabley.Fields("neto4") = ""
        mytabley.Fields("igv") = "" 'Format(Val("" & mytablex.Fields("impuesto")), "000000000.00")
        mytabley.Fields("ruc") = Trim("" & mytablex.Fields("codigo"))
        mytabley.Fields("tipo") = "C"
        mytabley.Fields("razonsocial") = Mid$(Trim("" & mytablex.Fields("nombre")), 1, 40)
        mytabley.Fields("glosa") = Mid$(Trim("" & mytablex.Fields("observa")), 1, 30)
        mytabley.Fields("tipoidentidad") = "1"
        mytabley.Fields("mediopago") = "001"
        mytabley.Fields("apellido1") = ""
        mytabley.Fields("apellido2") = ""
        mytabley.Fields("nombre") = ""
        mytabley.Fields("neto5") = ""
        mytabley.Fields("neto6") = ""
        mytabley.Fields("refnumero") = ""
        mytabley.Fields("reftipodoc") = ""
        mytabley.Fields("reffecha") = ""
        mytabley.Fields("detranumero") = ""
        mytabley.Fields("detrafecha") = ""
        mytabley.Update

        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Function poner_fijo(buf As String, n As Integer) As String

    Dim buf1 As String

    If Len(buf) > n Then
        buf = Mid$(buf, 1, n)
        poner_fijo = buf
        Exit Function

    End If

    buf1 = ""

    If Len(buf) < n Then

        For I = 1 To n - Len(buf)
            buf1 = buf1 & " "
        Next I

    End If

    buf = buf1 & buf
    poner_fijo = buf
    Exit Function

End Function
