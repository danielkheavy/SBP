VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReporteDelivery 
   Caption         =   "Form2"
   ClientHeight    =   10380
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10380
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   8535
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   15015
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   8055
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   14208
         _Version        =   393216
         HeadLines       =   3
         RowHeight       =   23
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CboVendedor 
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
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "%"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox Cboestado 
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
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox fechaf 
         Height          =   495
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   495
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn ChaGENERAR 
         Height          =   825
         Left            =   11040
         TabIndex        =   5
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1455
         BTYPE           =   4
         TX              =   "GENERAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteDelivery.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn cmdExcel 
         Height          =   825
         Left            =   11040
         TabIndex        =   6
         Top             =   1320
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1455
         BTYPE           =   4
         TX              =   "Exportar  Excel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteDelivery.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn7 
         Height          =   705
         Left            =   13320
         TabIndex        =   15
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
         BTYPE           =   4
         TX              =   "SALIR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteDelivery.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   17
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Items:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblOpción 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         TabIndex        =   9
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label lblFechaFinal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final:"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lblFechaInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio:"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1590
      End
   End
End
Attribute VB_Name = "FrmReporteDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rcconsultax As New ADODB.Recordset

Dim omytablex   As New ADODB.Recordset

Dim mytablex    As New ADODB.Recordset

' -- Variables para la base de datos
'Dim cnn         As Connection
'Dim rs          As Recordset
' -- Variables para Excel
Dim Obj_Excel   As Object

Dim Obj_Libro   As Object

Dim Obj_Hoja    As Object

Function sql_consultaX(sw As Integer)

    Dim buf        As String

    Dim queprecio  As String

    Dim mytablecom As New ADODB.Recordset

    'mytablecom.Open "SELECT * FROM mesa where salon='" & cmytablex.Fields("salon") & "' and mesa='" & cmytablex.Fields(1) & "'", cn, adOpenDynamic, adLockOptimistic
    'mytablecom.Open "SELECT * FROM mesa where salon='" & cmytablex.Fields("salon") & "' and mesa='" & cmytablex.Fields(1) & "'", cn, adOpenDynamic, adLockOptimistic
    Dim buf2       As String
   
    buf2 = "select count(tipo) as total from factura f  where f.local='01'  "
    buf2 = buf2 & " and estado='2' and f.servicio='D' "
    buf2 = buf2 & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' and fecha>='" & Format(fechai, "YYYYMMDD") & "' "
   
    If CboVendedor <> "%" Then
        buf2 = buf2 & " and vendedor like '" & extra_loquesea1(CboVendedor) & "'"

    End If

    If Cboestado <> "%" Then
        buf2 = buf2 & " and PLACA = '" & (Cboestado) & "'"

    End If
      
    Set rcconsultax = Nothing

    If rcconsultax.State = 1 Then
        rcconsultax.Close
        Set rcconsultax = Nothing

    End If
   
    mytablecom.Open buf2, cn, adOpenStatic, adLockOptimistic

    Label3 = mytablecom.Fields("total")

    If mytablecom.RecordCount = 0 Then
        mytablecom.Close

    End If

    mytablecom.Close
   
    buf = "select f.tipo,f.serie,f.Numero,f.Fecha,f.Nombre as Cliente,f.Codigo,f.Total,f.Placa as 'Estado',f.Vendedor AS Personal,f.Hora as 'Hora P.', horae as 'Hora Sal.',RENUMERO2 as 'Hora E.', f.Caja,f.Turno,d.telefono,d.direccion,d.referencia from factura f inner join deliveri d on d.codigo = f.codigo and f.local='01'  "
    buf = buf & " and estado='2' and f.servicio='D' "
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' and fecha>='" & Format(fechai, "YYYYMMDD") & "' "
   
    If CboVendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea1(CboVendedor) & "'"

    End If

    If Cboestado <> "%" Then
        buf = buf & " and PLACA = '" & (Cboestado) & "'"

    End If
   
    buf = buf & " order by f.HORA"
       
    Set rcconsultax = Nothing

    If rcconsultax.State = 1 Then
        rcconsultax.Close
        Set rcconsultax = Nothing

    End If
   
    rcconsultax.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DataGrid2.DataSource = rcconsultax
    DataGrid2.refresh
    
    DataGrid2.columns(0).Width = 450 'TIPO
    DataGrid2.columns(1).Width = 500 'SERIE
    DataGrid2.columns(2).Width = 900 'NUMERO
    DataGrid2.columns(3).Width = 1000 'fecha
    DataGrid2.columns(4).Width = 3200 ' cliente
    DataGrid2.columns(5).Width = 1200 'codigo
    DataGrid2.columns(6).Width = 900 ' total
    DataGrid2.columns(7).Width = 1100 ' estado
    DataGrid2.columns(8).Width = 1000 'personal
    DataGrid2.columns(9).Width = 1000 '
    DataGrid2.columns(10).Width = 1000 '
    DataGrid2.columns(11).Width = 1000
    DataGrid2.columns(12).Width = 500 '
    DataGrid2.columns(13).Width = 530 '
              
    DataGrid2.columns(13).Width = 900
    DataGrid2.columns(13).Width = 2000
    DataGrid2.columns(13).Width = 2000
               
    If rcconsultax.RecordCount = 0 Then
        Exit Function

    End If
              
    sql_consultaX = 1
    Exit Function
    'cmd8912_err:
    'MsgBox "Aviso en sql_consulta " & error$, 48, "Aviso"
    buffer = ""
    Exit Function

End Function

Function Formato_ordenDelivery(Num_Campos As Integer, _
                               Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(4, 1), .Cells(4, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, Num_Campos)).Font.bold = True
        .Range(.Cells(4, 1), .Cells(4, Num_Campos)).Interior.color = RGB(192, 192, 250)

        For I = 1 To Num_Campos Step 1
            .Cells(4, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        .columns("A").ColumnWidth = 4
        .columns("B").ColumnWidth = 6
        .columns("C").ColumnWidth = 9
        .columns("D").ColumnWidth = 11
        .columns("E").ColumnWidth = 20
        .columns("F").ColumnWidth = 12
        .columns("G").ColumnWidth = 8
        .columns("H").ColumnWidth = 11
        .columns("I").ColumnWidth = 10
        
        .columns("J").ColumnWidth = 9
        .columns("K").ColumnWidth = 9
        .columns("L").ColumnWidth = 8
        .columns("M").ColumnWidth = 5
        .columns("N").ColumnWidth = 6
        
        .columns("O").ColumnWidth = 12
        .columns("P").ColumnWidth = 20
        .columns("Q").ColumnWidth = 20
            
    End With

End Function

Sub reporte_excell(mytablex As ADODB.Recordset)

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim Heading(23) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Dim sumad       As Double

    Dim sumaa       As Double

    Dim sumat       As Double

    Heading(1) = "TIPO"
    Heading(2) = "SERIE"
    Heading(3) = "NÚMERO"
    Heading(4) = "FECHA"
    Heading(5) = "CLIENTE"
    Heading(6) = "CODIGO"
    Heading(7) = "TOTAL"
    Heading(8) = "ESTADO"
    Heading(9) = "PERSONAL"
    Heading(10) = "HORA P."
    Heading(11) = "HORA SAL."
    Heading(12) = "HORA E."
    Heading(13) = "CAJA"
    Heading(14) = "TURNO"
    
    Heading(15) = "TELÉFONO"
    Heading(16) = "DIRECCIÓN"
    Heading(17) = "REFERENCIA"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ordenDelivery(17, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    '    objExcel.ActiveSheet.Cells(1, 1) = "REPORTE >>>>>>>"
    '    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")

    objExcel.ActiveSheet.Cells(1, 6) = "      REPORTE DE DELIVERY"
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 5) = "FECHA FIN  " + fechaf
     
    v = 5
    h = 1
    sdx1 = 0
    Do

        If rcconsultax.EOF Then Exit Do
    
        objExcel.ActiveSheet.Cells(v, h) = "'" & rcconsultax.Fields("TIPO")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & rcconsultax.Fields("SERIE")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & rcconsultax.Fields("NUMERO")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & rcconsultax.Fields("FECHA")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & rcconsultax.Fields("CLIENTE")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & rcconsultax.Fields("CODIGO")
        objExcel.ActiveSheet.Cells(v, h + 6) = rcconsultax.Fields("TOTAL")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & rcconsultax.Fields("ESTADO")
        objExcel.ActiveSheet.Cells(v, h + 8) = "'" & rcconsultax.Fields("PERSONAL")
        objExcel.ActiveSheet.Cells(v, h + 9) = "'" & rcconsultax.Fields("HORA P.")
        objExcel.ActiveSheet.Cells(v, h + 10) = "'" & rcconsultax.Fields("HORA SAL.")
        objExcel.ActiveSheet.Cells(v, h + 11) = "'" & rcconsultax.Fields("HORA E.")
        objExcel.ActiveSheet.Cells(v, h + 12) = "'" & rcconsultax.Fields("CAJA")
        objExcel.ActiveSheet.Cells(v, h + 13) = "'" & rcconsultax.Fields("TURNO")
   
        objExcel.ActiveSheet.Cells(v, h + 14) = "'" & rcconsultax.Fields("TELEFONO")
        objExcel.ActiveSheet.Cells(v, h + 15) = "'" & rcconsultax.Fields("DIRECCION")
        objExcel.ActiveSheet.Cells(v, h + 16) = "'" & rcconsultax.Fields("REFERENCIA")
   
        sumad = sumad + Val("" & rcconsultax.Fields("TOTAL"))
        objExcel.ActiveSheet.Cells(v + 1, 7) = sumad
    
        v = v + 1
        rcconsultax.MoveNext
    Loop
 
    For I = 1 To 14
        objExcel.ActiveSheet.Cells(v, I).Font.bold = True
        objExcel.ActiveSheet.Cells(v, I).Interior.color = RGB(248, 243, 53)
    Next
 
    Set objExcel = Nothing

End Sub

Private Sub ChaGENERAR_Click()
 
    Dim found As Integer
 
    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub

    found = sql_consultaX(1)

End Sub

Private Sub ChameleonBtn7_Click()
    FrmReporteDelivery.Hide
    Unload FrmReporteDelivery

End Sub

Private Sub cmdExcel_Click()

    Dim found As Integer

    found = sql_consultaX(1)
    reporte_excell omytablex

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

    Cboestado.Clear
    Cboestado.AddItem "%"
    Cboestado.AddItem "PENDIENTE"
    Cboestado.AddItem "ENTREGADO"
    Cboestado.ListIndex = 0

    CboVendedor.Clear
    CboVendedor.AddItem "%"
    cad = "SELECT nombre,codigo FROM vendedor WHERE CARGO='MOTORIZADO' order by nombre "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        CboVendedor.AddItem "" & mytablex.Fields("NOMBRE") & "|" & mytablex.Fields("CODIGO")
        mytablex.MoveNext
    Loop
    CboVendedor.ListIndex = 0
    mytablex.Close
   
End Sub

