VERSION 5.00
Begin VB.Form trepocli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Clientes Productos"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox tiporeporte 
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2400
      Width           =   2535
   End
   Begin VB.ComboBox ntipo 
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ComboBox nclasifica 
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Reporte"
      Height          =   420
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clasificacion"
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   420
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu lfo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trepocli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub

    If tiporeporte = "Resumen" Then
        buf = "SELECT     dbo.clientes.TIPOCLIE, dbo.clientes.CLASIFICA,dbo.clientes.codigo,   dbo.detalle.PRODUCTO, "
        buf = buf & "   sum(dbo.detalle.cantidad) as cant, sum( dbo.detalle.total) as tot "
        buf = buf & " FROM         dbo.detalle INNER JOIN "
        buf = buf & "  dbo.clientes ON dbo.detalle.CODIGO = dbo.clientes.codigo"
        buf = buf & " and dbo.detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and dbo.detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "'"

        If Len(Trim(nclasifica)) <> 0 Then
            buf = buf & " and dbo.clientes.clasifica='" & extra_loquesea1(nclasifica) & "'"

        End If

        If Len(Trim(ntipo)) <> 0 Then
            buf = buf & " and dbo.clientes.tipoclie='" & extra_loquesea1(ntipo) & "'"

        End If

        buf = buf & " and (dbo.detalle.acu='A' OR dbo.detalle.acu='B' OR dbo.detalle.acu='C' OR dbo.detalle.acu='D' OR dbo.detalle.acu='G')"
        buf = buf & " and dbo.detalle.estado='2' group by dbo.clientes.tipoclie,dbo.clientes.clasifica,dbo.clientes.codigo,dbo.detalle.producto order by dbo.clientes.tipoclie,dbo.clientes.clasifica"
        'MsgBox buf
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If

        reporte_excell mytablex

    End If

    If tiporeporte = "Detalle" Then
        buf = "SELECT     dbo.clientes.TIPOCLIE, dbo.clientes.CLASIFICA,dbo.clientes.codigo,   dbo.detalle.PRODUCTO, "
        buf = buf & "   dbo.detalle.cantidad,dbo.detalle.fecha , dbo.detalle.total "
        buf = buf & " FROM         dbo.detalle INNER JOIN "
        buf = buf & "  dbo.clientes ON dbo.detalle.CODIGO = dbo.clientes.codigo"
        buf = buf & " and dbo.detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and dbo.detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "'"

        If Len(Trim(nclasifica)) <> 0 Then
            buf = buf & " and dbo.clientes.clasifica='" & extra_loquesea1(nclasifica) & "'"

        End If

        If Len(Trim(ntipo)) <> 0 Then
            buf = buf & " and dbo.clientes.tipoclie='" & extra_loquesea1(ntipo) & "'"

        End If

        buf = buf & " and (dbo.detalle.acu='A' OR dbo.detalle.acu='B' OR dbo.detalle.acu='C' OR dbo.detalle.acu='D' OR dbo.detalle.acu='G')"
        buf = buf & " and dbo.detalle.estado='2'  order by dbo.clientes.tipoclie,dbo.clientes.clasifica,dbo.clientes.codigo"
        'MsgBox buf
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If

        reporte_excell1 mytablex

    End If

End Sub

Private Sub Form_Load()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_clasifica
    subclasifica

    tiporeporte.Clear
    tiporeporte.AddItem "Resumen"
    tiporeporte.AddItem "Detalle"
    tiporeporte.ListIndex = 0

End Sub

Private Sub lfo44_Click()
    trepocli.Hide
    Unload trepocli

End Sub

Sub reporte_excell(mytablex As ADODB.Recordset)

    Dim xhoy        As String

    Dim dias        As Integer

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim xtotal      As Double

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Dim Fecha1      As Date

    Dim Fecha2      As Date

    Dim meses       As Integer

    Dim mytabley    As New ADODB.Recordset

    Command1.Visible = True

    On Error GoTo cmd6561245_err
    
    Heading(1) = "Codigo"
    Heading(2) = "Nombre"
    Heading(3) = "Tipo"
    Heading(4) = "Clasifica"
    Heading(5) = "Descripcio"
    Heading(6) = "Cantidad"
    Heading(7) = "Total"

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    objExcel.ActiveSheet.Cells(1, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA HOY  " + Format(Now, "HH:MM:SS")
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO :" + Format(fechai, "DD/MM/YYYY") & " FECHA FINAL :" + Format(fechaf, "DD/MM/YYYY")

    v = 4
    h = 1
    sdx1 = 0
    sdx2 = 0
    
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("codigo")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_cliente(Trim("" & mytablex.Fields("codigo")))
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("tipoclie")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("clasifica")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & busca_producto(Trim("" & mytablex.Fields("producto")))
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("cant")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("tot")

        sdx1 = sdx1 + Val("" & mytablex.Fields("tot"))
        sdx2 = sdx2 + Val("" & mytablex.Fields("cant"))
        v = v + 1
        mytablex.MoveNext
    Loop
    objExcel.ActiveSheet.Cells(v, h + 5) = "" & sdx2
    objExcel.ActiveSheet.Cells(v, h + 6) = "" & sdx1

    Set objExcel = Nothing
    Exit Sub
cmd6561245_err:
    MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.bold = True
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 30
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 30
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 10
        .columns("i").ColumnWidth = 10
        .columns("j").ColumnWidth = 10
        .columns("k").ColumnWidth = 10
        .columns("l").ColumnWidth = 10
    
    End With

End Function

Sub carga_clasifica()

    Dim mytablex As New ADODB.Recordset

    nclasifica.Clear
    nclasifica.AddItem ""
    mytablex.Open "select * from clasifi", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        nclasifica.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("clasifica"))
        mytablex.MoveNext
    Loop
    nclasifica.ListIndex = 0

End Sub

Sub subclasifica()

    Dim mytablex As New ADODB.Recordset

    ntipo.Clear
    ntipo.AddItem ""
    mytablex.Open "select * from tipoclie ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        ntipo.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("tipoclie"))
        mytablex.MoveNext
    Loop
    ntipo.ListIndex = 0

End Sub

Function busca_cliente(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_cliente = Trim("" & mytablex.Fields("nombre"))

    End If

    mytablex.Close

End Function

Function busca_producto(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM producto where  producto='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        busca_producto = Trim("" & mytablex.Fields("descripcio"))

    End If

    mytablex.Close

End Function

Sub reporte_excell1(mytablex As ADODB.Recordset)

    Dim xhoy        As String

    Dim dias        As Integer

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim xtotal      As Double

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Dim Fecha1      As Date

    Dim Fecha2      As Date

    Dim meses       As Integer

    Dim mytabley    As New ADODB.Recordset

    Command1.Visible = True

    On Error GoTo cmd65561245_err
    
    Heading(1) = "Codigo"
    Heading(2) = "Nombre"
    Heading(3) = "Tipo"
    Heading(4) = "Clasifica"
    Heading(5) = "Descripcio"
    Heading(6) = "Cantidad"
    Heading(7) = "Total"
    Heading(8) = "Fecha"

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    objExcel.ActiveSheet.Cells(1, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA HOY  " + Format(Now, "HH:MM:SS")
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO :" + Format(fechai, "DD/MM/YYYY") & " FECHA FINAL :" + Format(fechaf, "DD/MM/YYYY")

    v = 4
    h = 1
    sdx1 = 0
    sdx2 = 0
    
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("codigo")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_cliente(Trim("" & mytablex.Fields("codigo")))
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("tipoclie")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("clasifica")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & busca_producto(Trim("" & mytablex.Fields("producto")))
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("cantidad")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("total")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("fecha")

        sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
        sdx2 = sdx2 + Val("" & mytablex.Fields("cantidad"))
        v = v + 1
        mytablex.MoveNext
    Loop
    objExcel.ActiveSheet.Cells(v, h + 5) = "" & sdx2
    objExcel.ActiveSheet.Cells(v, h + 6) = "" & sdx1

    Set objExcel = Nothing
    Exit Sub
cmd65561245_err:
    MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    Exit Sub

End Sub

