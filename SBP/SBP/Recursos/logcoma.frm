VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form logcoma 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos Comandadas Eliminados antes de Pagarse"
   ClientHeight    =   8040
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5430
      Left            =   255
      TabIndex        =   8
      Top             =   1860
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   9578
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.CommandButton cmdExportExcell 
      Caption         =   "Exportar a excell"
      Height          =   360
      Left            =   10470
      TabIndex        =   7
      Top             =   7485
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresca"
      Height          =   615
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   255
      Width           =   1455
   End
   Begin VB.ComboBox vendedor 
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
      TabIndex        =   5
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Autorizado Borrado"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "logcoma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim omytablex As New ADODB.Recordset

' -----------------------------------------------------------------------------------------
' \\ --  Descripción       : Exportar DataGrid a Excel
' \\ --  Controles         : Un Datagrid, un CommandButton y la referencia a ADO
' \\ --  Autor             : Luciano Lodola -- http://www.recursosvisualbasic.com.ar/
' -----------------------------------------------------------------------------------------
  
' -- Variables para la base de datos
'Dim cnn         As Connection
'Dim rs          As Recordset
' -- Variables para Excel
Dim Obj_Excel As Object

Dim Obj_Libro As Object

Dim Obj_Hoja  As Object
  
' -----------------------------------------------------------------------------------------
' \\ -- Sub para exportar
' -----------------------------------------------------------------------------------------
Private Sub exportar_Datagrid(Datagrid As Datagrid, n_Filas As Long)

    ''    'On Error GoTo Error_Handler
    ''    Dim i   As Integer
    ''    Dim j   As Integer
    ''    Dim iCol As Long
    ''    Dim vPATH As String
    ''    vPATH = "C:\Test.xls"
    ''    ' -- Colocar el cursor de espera mientras se exportan los datos
    ''    Me.MousePointer = vbHourglass
    ''
    ''    If n_Filas = 0 Then
    ''        MsgBox "No hay datos para exportar a excel. Se ha indicado 0 en el parámetro Filas ": Exit Sub
    ''    Else
    ''
    ''        ' -- Crear nueva instancia de Excel
    ''        Set Obj_Excel = CreateObject("Excel.Application")
    ''        ' -- Agregar nuevo libro
    ''        'Set Obj_Libro = Obj_Excel.Workbooks.Open(path)
    ''        Set Obj_Libro = Obj_Excel.Workbooks.Open(path)
    ''
    ''        ' -- Referencia a la Hoja activa ( la que añade por defecto Excel )
    ''        Set Obj_Hoja = Obj_Excel.ActiveSheet
    ''
    ''        iCol = 0
    ''        ' --  Recorrer el Datagrid ( Las columnas )
    ''        For i = 0 To Datagrid.columns.count - 1
    ''            If Datagrid.columns(i).Visible Then
    ''                ' -- Incrementar índice de columna
    ''                iCol = iCol + 1
    ''                ' -- Obtener el caption de la columna
    ''                Obj_Hoja.Cells(1, iCol) = Datagrid.columns(i).Caption
    ''                ' -- Recorrer las filas
    ''                For j = 0 To n_Filas - 1
    ''                'For j = 0 To Me.Datagrid1.ApproxCount - 1
    ''                    ' -- Asignar el valor a la celda del Excel
    ''                    Obj_Hoja.Cells(j + 2, iCol) = _
    ''                    Datagrid.columns(i).CellValue(Datagrid.GetBookmark(j))
    ''                Next
    ''            End If
    ''        Next
    ''
    ''        ' -- Hacer excel visible
    ''        Obj_Excel.Visible = True
    ''
    ''        ' -- Opcional : colocar en negrita y de color rojo los enbezados en la hoja
    ''        With Obj_Hoja
    ''            .rows(1).Font.bold = True
    ''            .rows(1).Font.color = vbRed
    ''            ' -- Autoajustar las cabeceras
    ''            .columns("A:Z").AutoFit
    ''        End With
    ''    End If
    ''
    ''    ' -- Eliminar las variables de objeto excel
    ''    Set Obj_Hoja = Nothing
    ''    Set Obj_Libro = Nothing
    ''    Set Obj_Excel = Nothing
    ''
    ''    ' -- Restaurar cursor
    ''    Me.MousePointer = vbDefault
    ''
    ''Exit Sub
    ''
    '''' -- Error
    '''Error_Handler:
    '''
    '''    MsgBox Err.Description, vbCritical
    '''    On Error Resume Next
    '''
    '''    Set Obj_Hoja = Nothing
    '''    Set Obj_Libro = Nothing
    '''    Set Obj_Excel = Nothing
    '''    Me.MousePointer = vbDefault
    ''
End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

    ''    On Error GoTo Error_Handler
    ''    ' -- Crear nueva conexión a la base de datos
    ''    Set cnn = New Connection
    ''    ' -- Abrir la base de datos.
    ''    cnn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\data\bd1.mdb"
    ''    ' -- Crear nuevo objeto Recordset
    ''    Set rs = New Recordset
    ''    ' -- Configurar recordset
    ''    With rs
    ''        .CursorLocation = adUseClient
    ''    End With
    ''    ' -- Cargar el recordset ( ESPECIFICAR LA CONSULTA SQL )
    'rs.Open "Select * From tabla1", cnn, adOpenStatic, adLockOptimistic
    ' -- Enlazar el datagrid con el recordset anterior
    ' Set DataGrid1.DataSource = rs
    'Command1.Caption = " Exportar datagrid a Excel "
    ' -- Errores
    'Exit Sub
    'Error_Handler:
    'MsgBox Err.Description, vbCritical, "Error en Form Load"
End Sub

' -------------------------------------------------------------------------------
' \\ -- Fin
' -------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    ' -- Cerrar y eliminar recordset
    If omytablex.State = adStateOpen Then omytablex.Close
    If Not omytablex Is Nothing Then Set omytablex = Nothing

    ' -- cerrar y Eliminar la conexión
    '    If cnn.State = adStateOpen Then cnn.Close
    '    Set cnn = Nothing
End Sub

Private Sub cmdExportExcell_Click()
    'Call exportar_Datagrid(DataGrid1, CLng(DataGrid1.ApproxCount))
    reporte_excell omytablex

End Sub

Sub reporte_excell(mytablex As ADODB.Recordset)

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Command1.Visible = True
    'On Error GoTo cmd6561245_err
    'omytablex.Open "SELECT administrador as Autorizo,observa1 as Motivo,FechaBorra,Salon,Mesa,Vendedor,HoraBorra,Producto,Descripcio,
    'Unidad as Und,Cantidad as Cant,Precio,Total,Caja, Turno FROM logcomanda   " & buf & "  order by fecha,hora", cn, adOpenStatic, adLockOptimistic

    Heading(1) = "Autorizo":    Heading(2) = "Motivo":    Heading(3) = "FechaBorra":    Heading(4) = "Salon":    Heading(5) = "Mesa"
    Heading(6) = "Vendedor":    Heading(7) = "HoraBorra":    Heading(8) = "CodProducto":    Heading(9) = "Producto":    Heading(10) = "Und"
    Heading(11) = "Cant":    Heading(12) = "Precio":    Heading(13) = "Total":    Heading(14) = "Caja":    Heading(15) = "Turno"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    objExcel.ActiveSheet.Cells(1, 1) = "REPORTE DE PRODUCTOS ANULADOS POR COMANDA-SALON Y MESA"
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")

    v = 4
    h = 1
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("Autorizo")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("Motivo")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("FechaBorra")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("Salon")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("Mesa")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("Vendedor")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("HoraBorra")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytablex.Fields("Producto")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytablex.Fields("descripcio") 'nombre de producto
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("Und") 'unidad
        objExcel.ActiveSheet.Cells(v, h + 10) = "" & mytablex.Fields("Cant") 'cantidad
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & mytablex.Fields("Precio")
        objExcel.ActiveSheet.Cells(v, h + 12) = "" & mytablex.Fields("Total")
        objExcel.ActiveSheet.Cells(v, h + 13) = "" & mytablex.Fields("Caja")
        objExcel.ActiveSheet.Cells(v, h + 14) = "" & mytablex.Fields("Turno")
        v = v + 1
        'imprime_recetaa mytablex, v, h
        mytablex.MoveNext
    Loop
    Set objExcel = Nothing

    'Exit Sub
    'cmd6561245_err:
    'MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    'Exit Sub
End Sub

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 15)).Font.bold = True
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 30 'motivo
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 15 ' cod-producto
        .columns("i").ColumnWidth = 30 'nombre producto
        .columns("j").ColumnWidth = 7
        .columns("k").ColumnWidth = 7
        .columns("l").ColumnWidth = 7

    End With

End Function

Private Sub Command1_Click()

    Dim buf As String

    buf = buf & "  where fechaborra>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fechaborra<='" & Format(fechaf, "YYYYMMDD") & "' "

    If Trim(vendedor) <> "%" Then
        buf = buf & " and vendedor='" & extra_loquesea(vendedor) & "'"

    End If

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "SELECT administrador as Autorizo,observa1 as Motivo,FechaBorra,Salon,Mesa,Vendedor,HoraBorra,Producto,Descripcio,Unidad as Und,Cantidad as Cant,Precio,Total,Caja, Turno FROM logcomanda   " & buf & "  order by fecha,hora", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Private Sub flo44_Click()
    logcoma.Hide
    Unload logcoma

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    vendedor.Clear
    vendedor.AddItem "%"
    mytablex.Open "SELECT * FROM vendedor  order by nombre", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & "" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    vendedor.ListIndex = 0
    mytablex.Close

End Sub

