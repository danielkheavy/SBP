VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReporteGeneralTrans 
   BackColor       =   &H00808080&
   Caption         =   "Reporte General"
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
      Height          =   6975
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   15015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   8295
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   14631
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
      Top             =   0
      Width           =   15015
      Begin VB.TextBox producto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MaxLength       =   15
         TabIndex        =   16
         Text            =   "%"
         Top             =   1440
         Width           =   2655
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
         Left            =   840
         MaxLength       =   15
         TabIndex        =   12
         Text            =   "%"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox familia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox subfamilia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox CboOpcion 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
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
      Begin ChamaleonButton.ChameleonBtn cmdrefresca 
         Height          =   825
         Left            =   11040
         TabIndex        =   4
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
         MICON           =   "FrmReporteGeneralTrans.frx":0000
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1455
         BTYPE           =   4
         TX              =   "Exportar  Excel"
         ENAB            =   0   'False
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
         MICON           =   "FrmReporteGeneralTrans.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChaCERRAR 
         Height          =   705
         Left            =   13440
         TabIndex        =   9
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
         BTYPE           =   4
         TX              =   "CERRAR"
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
         BCOL            =   4210752
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteGeneralTrans.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label items 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "items"
         Height          =   195
         Left            =   14160
         TabIndex        =   18
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblProductoS 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblFamilia 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   14
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblSubfamilia 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Subfamilia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   13
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label lblOpción 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opción:"
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
         Top             =   1920
         Width           =   1005
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
         TabIndex        =   6
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label1 
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
         TabIndex        =   5
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   0
      X2              =   15000
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "FrmReporteGeneralTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim omytablex As New ADODB.Recordset

Dim Obj_Excel As Object

Dim Obj_Libro As Object

Dim Obj_Hoja  As Object

Public valor  As String

Private Sub ChaCERRAR_Click()
     
    FrmReporteGeneralTrans.Hide
    Unload FrmReporteGeneralTrans

End Sub

Private Sub cmdExcel_Click()

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        fechaf.SetFocus
        Exit Sub

    End If

    If Len(fechai) = 0 Then
        MsgBox "Fechai No valida", 48, "Aviso"
        Exit Sub

    End If

    If Len(fechai) <> 10 Then
        MsgBox "Fechai No valida", 48, "Aviso"
        Exit Sub

    End If
    
    If CboOpcion.ListIndex = 0 Then ReporteTotal
    If CboOpcion.ListIndex = 1 Then ReporteVentas
    If CboOpcion.ListIndex = 2 Then ReporteCompras
    If CboOpcion.ListIndex = 3 Then ReporteSalida
    If CboOpcion.ListIndex = 4 Then ReporteEntrada
    If CboOpcion.ListIndex = 5 Then ReporteStock

    items = 0
    items = DataGrid1.Row

    reporte_excell omytablex

End Sub

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 23)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, 23)).Interior.color = RGB(192, 192, 250)
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        .columns("A").ColumnWidth = 9
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 9
        .columns("D").ColumnWidth = 15
        .columns("E").ColumnWidth = 12
        .columns("F").ColumnWidth = 25
        .columns("G").ColumnWidth = 11
        .columns("H").ColumnWidth = 11
        .columns("I").ColumnWidth = 12 ' CONDICION DE PAGO
        .columns("J").ColumnWidth = 11
        .columns("K").ColumnWidth = 11
        .columns("L").ColumnWidth = 11
        .columns("M").ColumnWidth = 11
        .columns("N").ColumnWidth = 25
        .columns("O").ColumnWidth = 7
        .columns("P").ColumnWidth = 7
        .columns("Q").ColumnWidth = 8
        .columns("R").ColumnWidth = 10
        .columns("S").ColumnWidth = 8
        .columns("T").ColumnWidth = 5
        .columns("U").ColumnWidth = 8
        .columns("V").ColumnWidth = 8
        .columns("W").ColumnWidth = 8
        
    End With

End Function

Function Formato_ordenTotal() As Boolean

    'With objExcel.ActiveSheet
    '     .Range(.Cells(items + 4, 1), .Cells(items + 4, 23)).Font.bold = True
    '     .Range(.Cells(items + 4, 1), .Cells(items + 4, 23)).Interior.Color = RGB(248, 243, 53)
    '
    '
    '
    'End With

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

    Dim sumac       As Double

    Dim sumat       As Double

    Heading(1) = "TIP-DOC"
    Heading(2) = "ALM"
    Heading(3) = "ALM. DESC"
    Heading(4) = "DOCUMENTO"
    Heading(5) = "COD-CLI"
    Heading(6) = "RAZON SOCIAL"
    Heading(7) = "F-EMISION"
    Heading(8) = "F-VCTO"
    Heading(9) = "CON-PAGO"
    Heading(10) = "FAMILIA"
    Heading(11) = "SUBFAMILIA"
    Heading(12) = "CATEGORIA"
    Heading(13) = "CODIGO"
    Heading(14) = "ARTICULO"
    Heading(15) = "U.MED"
    Heading(16) = "CANT."
    Heading(17) = "PRECIO"
    Heading(18) = "TOTAL"
    Heading(19) = "USUARIO"
    Heading(20) = "CAJA"
    Heading(21) = "HORA"
    Heading(22) = "TURNO"
    Heading(23) = "COSTO"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(23, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    Call Formato_ordenTotal
    
    objExcel.ActiveSheet.Cells(1, 1) = "REPORTE >>>>>>>"
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")

    v = 4
    h = 1
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("TIP-DOC")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("ALM")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("ALM.DESC")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("DOCUMENTO")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("COD-CLI")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("RAZON SOCIAL")
    
        If mytablex.Fields("F-EMISION") = "01/01/1900" Then
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
        Else
            objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytablex.Fields("F-EMISION")

        End If
    
        If mytablex.Fields("F-VCTO") = "01/01/1900" Then
            objExcel.ActiveSheet.Cells(v, h + 7) = ""
        Else
            objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("F-VCTO")

        End If
    
        ' MUESTRA CONDICION DE PAGO SEGUN FPAGO
    
        If mytablex.Fields("CON-PAGO") = "3" Then  'CREDITO
            objExcel.ActiveSheet.Cells(v, h + 8) = "CREDITO"
        ElseIf mytablex.Fields("CON-PAGO") = "" Then
            objExcel.ActiveSheet.Cells(v, h + 8) = ""
        Else
            objExcel.ActiveSheet.Cells(v, h + 8) = "CONTADO"

        End If

        ' MUESTRA CONDICION DE PAGO SEGUN FPAGO
    
        objExcel.ActiveSheet.Cells(v, h + 9) = "'" & mytablex.Fields("FAMILIA")
        objExcel.ActiveSheet.Cells(v, h + 10) = "'" & mytablex.Fields("SUBFAMILIA")
        objExcel.ActiveSheet.Cells(v, h + 11) = "'" & mytablex.Fields("CATEGORIA")
        objExcel.ActiveSheet.Cells(v, h + 12) = "'" & mytablex.Fields("CODIGO")
        objExcel.ActiveSheet.Cells(v, h + 13) = "'" & mytablex.Fields("ARTICULO")
        objExcel.ActiveSheet.Cells(v, h + 14) = "'" & mytablex.Fields("U.MED.")
        objExcel.ActiveSheet.Cells(v, h + 15) = "" & mytablex.Fields("CANT.")
        objExcel.ActiveSheet.Cells(v, h + 16) = "" & mytablex.Fields("PRECIO")
        objExcel.ActiveSheet.Cells(v, h + 17) = "" & mytablex.Fields("TOTAL")
        objExcel.ActiveSheet.Cells(v, h + 18) = "'" & mytablex.Fields("USUARIO")
        objExcel.ActiveSheet.Cells(v, h + 19) = "'" & mytablex.Fields("CAJA") '
        objExcel.ActiveSheet.Cells(v, h + 20) = "" & mytablex.Fields("HORA") '
        objExcel.ActiveSheet.Cells(v, h + 21) = "" & mytablex.Fields("TURNO") '
        objExcel.ActiveSheet.Cells(v, h + 22) = "" & mytablex.Fields("COSTO")
    
        '    If mytablex.Fields("CON-PAGO") = "VENTAS" Then
        '     objExcel.ActiveSheet.Cells(v, h + 22) = "" & mytablex.Fields("COSTO") '
        '    Else
        '       objExcel.ActiveSheet.Cells(v, h + 22) = ""
        '    End If
  
        objExcel.ActiveSheet.Cells(v + 1, 15) = "NETO"
        
        sumac = sumac + Val("" & mytablex.Fields("CANT."))
        sumat = sumat + Val("" & mytablex.Fields("TOTAL"))
 
        objExcel.ActiveSheet.Cells(v + 1, 16) = sumac
        objExcel.ActiveSheet.Cells(v + 1, 18) = sumat

        v = v + 1
    
        mytablex.MoveNext
 
    Loop

    For I = 1 To 23
        objExcel.ActiveSheet.Cells(v, I).Font.bold = True
        objExcel.ActiveSheet.Cells(v, I).Interior.color = RGB(248, 243, 53)
    Next
  
    Set objExcel = Nothing

    'Exit Sub
    'cmd6561245_err:
    'MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    'Exit Sub
End Sub

Sub ReporteUpdateClientesVentas()

    Dim buf As String

    buf = buf & " IF not exists (SELECT * FROM clientes WHERE codigo ='123456') bEGIN  insert into clientes(CODIGO,NOMBRE,ESTADO,TIPO) "
    buf = buf & " values('123456','CLIENTES VARIOS','A','N')   END "
    buf = buf & " UPDATE DETALLE SET CODIGO='123456' WHERE CODIGO='' AND fecha<='" & Format(fechaf, "YYYYMMDD") & "' and fecha>='" & Format(fechai, "YYYYMMDD") & "' "
    buf = buf & " UPDATE FACTURA SET CODIGO='123456', NOMBRE='CLIENTES VARIOS' WHERE CODIGO='' AND fecha<='" & Format(fechaf, "YYYYMMDD") & "' and fecha>='" & Format(fechai, "YYYYMMDD") & "' "
    buf = buf & " UPDATE FPAGOV SET CODIGO='123456',NOMBRE='CLIENTES VARIOS'  WHERE CODIGO='' AND fecha<='" & Format(fechaf, "YYYYMMDD") & "' and fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & " ", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteVentas()

    ReporteUpdateClientesVentas

    cmdExcel.Enabled = True

    Dim buf As String

    'buf = buf & " SELECT 'VENTAS' AS 'TIP-DOC',D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO,   "
    'buf = buf & " D.CODIGO AS 'COD-CLI',CL.NOMBRE AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO',   fp.fpago as 'CON-PAGO', "
    'buf = buf & " D.FAMILIA,S.DESCRIPCIO AS SUBFAMILIA,D.CATEGORIA,D.PRODUCTO AS 'CODIGO',  "
    'buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.',D.PRECIO,D.TOTAL,   "
    'buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO,P.COSTOP AS 'COSTO'  "
    'buf = buf & " FROM subfamil S FULL join PRODUCTO P on S.subfamilia=p.subfamilia INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO   "
    'buf = buf & " INNER JOIN CLIENTES CL ON D.CODIGO=CL.CODIGO   "
    'buf = buf & " INNER JOIN FACTURA F ON (F.TIPO=D.TIPO AND F.SERIE=D.SERIE AND F.NUMERO=D.NUMERO)  inner join fpagov fp on (Fp.TIPO=D.TIPO AND Fp.SERIE=D.SERIE AND Fp.NUMERO=D.NUMERO) "
    'buf = buf & " AND (D.acu='1' or D.acu='A' or D.acu='B' or D.acu='C' or D.acu='D' or D.acu='G' or D.acu='E' or D.acu='F')   "
    'buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    buf = buf & " SELECT 'VENTAS' AS 'TIP-DOC',D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO,   "
    buf = buf & " D.CODIGO AS 'COD-CLI',CL.NOMBRE AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO',   f.c9 as 'CON-PAGO', "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS SUBFAMILIA,D.CATEGORIA,D.PRODUCTO AS 'CODIGO',  "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.',D.PRECIO,D.TOTAL,   "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO,P.COSTOP AS 'COSTO'  "
    buf = buf & " FROM PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO   "
    buf = buf & " INNER JOIN CLIENTES CL ON D.CODIGO=CL.CODIGO   "
    buf = buf & " INNER JOIN FACTURA F ON (F.TIPO=D.TIPO AND F.SERIE=D.SERIE AND F.NUMERO=D.NUMERO) "
    buf = buf & " AND (D.acu='1' or D.acu='A' or D.acu='B' or D.acu='C' or D.acu='D' or D.acu='G' or D.acu='E' or D.acu='F')   "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by d.fecha,d.hora", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteCompras()
    cmdExcel.Enabled = True

    Dim buf As String

    buf = buf & " SELECT 'COMPRAS' AS 'TIP-DOC',D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO,   "
    buf = buf & " D.CODIGO AS 'COD-CLI',PR.NOMBRE AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO', F.fpago as 'CON-PAGO',   "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS 'SUBFAMILIA',D.CATEGORIA,D.PRODUCTO AS 'CODIGO',  "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.',D.PRECIO,D.TOTAL,   "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO ,'' AS 'COSTO'   "
    buf = buf & " FROM PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO   "
    buf = buf & " INNER JOIN PROVEEDO PR  ON D.CODIGO=PR.CODIGO   "
    buf = buf & " INNER JOIN FACTURA F ON (F.TIPO=D.TIPO AND F.SERIE=D.SERIE AND F.NUMERO=D.NUMERO)   "
    buf = buf & " AND (D.acu='J' or D.acu='K' or D.acu='L' or D.acu='M' or D.acu='P' or D.acu='N' or D.acu='O')  "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    '
    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by d.fecha,d.hora", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteSalida()
    cmdExcel.Enabled = True

    Dim buf As String

    buf = buf & " SELECT 'SALIDA' AS 'TIP-DOC', D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO, "
    buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', '' AS 'F-VCTO', '' as 'CON-PAGO', "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS 'SUBFAMILIA',D.CATEGORIA,D.PRODUCTO AS 'CODIGO', "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.','' AS PRECIO,'' AS TOTAL, "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO ,'' AS 'COSTO'"
    buf = buf & " FROM PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO AND (D.acu='T' ) "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by d.fecha,d.hora", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteEntrada()
    cmdExcel.Enabled = True

    Dim buf As String

    buf = buf & " SELECT 'ENTRADA' AS 'TIP-DOC', D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO, "
    buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', '' AS 'F-VCTO','' as 'CON-PAGO',  "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS 'SUBFAMILIA',D.CATEGORIA,D.PRODUCTO AS 'CODIGO', "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.','' AS PRECIO,'' AS TOTAL, "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO ,'' AS 'COSTO' "
    buf = buf & " FROM  PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO AND (D.acu='S' ) "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by d.fecha,d.hora", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub GeneraReporteSaldo()

    Dim buf As String

    '''Crea tabla temporal Reporte
    buf = buf & " IF  exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='reportes') BEGIN  drop table reportes END "
    buf = buf & " CREATE TABLE reportes ( familia varCHAR(6), subfamilia varCHAR(6),categoria varCHAR(6), codigo varCHAR(15),"
    buf = buf & " articulo varCHAR(120), unidad   varCHAR(8),ingreso float,salida float, saldo float,costo float); "

    '''Registra datos a tabla temporal

    buf = buf & " INSERT INTO reportes (familia,subfamilia,categoria,codigo,articulo,"
    buf = buf & " unidad,ingreso,salida,saldo,costo)"
    buf = buf & " select FAMILIA,A.[SUB-FAMILIA] AS'SUBFAMILIA', ISNULL(A.CATEGORIA,'') AS 'CATEGORIA',A.CODIGO,A. ARTICULO ,"
    buf = buf & " A.UNIDAD  AS 'U.MED.',(a.INGRESO) as 'CANT.','','',a.costou AS 'COSTO'"
    buf = buf & " from (select D.producto AS 'CODIGO',P.DESCRIPCIO AS"
    buf = buf & " 'ARTICULO',P.FAMILIA AS 'FAMILIA', d.subfamilia AS 'SUB-FAMILIA',"
    buf = buf & " D.CATEGORIA, D.UNIDAD, SUM(CANTIDAD) as 'INGRESO',p.costou FROM DETALLE D"
    buf = buf & " INNER JOIN producto p on p.producto=d.PRODUCTO and d.ESTADO='2'  AND (D.acu='J' or D.acu='K'"
    buf = buf & " or D.acu='L' or D.acu='M' or D.acu='P' or D.acu='N' or D.acu='O' or D.acu='S')"
    buf = buf & " and d.fecha<='" & Format(fechai, "YYYYMMDD") & "' "

    buf = buf & " GROUP BY D.PRODUCTO,D.CATEGORIA, P.DESCRIPCIO,d.subfamilia,p.FAMILIA ,"
    buf = buf & " D.UNIDAD,p.costou ) as a WHERE A.CODIGO<>''  "

    If familia <> "%" Then
        buf = buf & " and a.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and a.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and a.codigo like '" & producto & "'"

    End If

    buf = buf & " order by A.ARTICULO"

    '''Actualiza Salidas

    buf = buf & " DECLARE @intFlag INT DECLARE @j INT SET @intFlag = 1 SET @j ="
    buf = buf & " (select count(a.row) from(SELECT ROW_NUMBER() OVER(ORDER BY d.producto DESC) AS 'Row',"
    buf = buf & " d.producto FROM DETALLE D where D.ESTADO='2' AND (D.acu='1' or D.acu='A' or"
    buf = buf & " D.acu='B' or D.acu='C' or   D.acu='D' or D.acu='G' or D.acu='E' or D.acu='F' or D.acu='T')"
    buf = buf & " and  d.fecha<='" & Format(fechai, "YYYYMMDD") & "' "
 
    If familia <> "%" Then
        buf = buf & " and D.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and D.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and D.PRODUCTO like '" & producto & "'"

    End If
 
    buf = buf & " group by d.producto) as a) WHILE (@intFlag <=@j )"
    buf = buf & " BEGIN"

    buf = buf & " update reportes set salida=(SELECT SUM(CANTIDAD) FROM DETALLE D where D.ESTADO='2' AND (D.acu='1' or D.acu='A' or"
    buf = buf & " D.acu='B' or D.acu='C' or   D.acu='D' or D.acu='G' or D.acu='E' or D.acu='F' or D.acu='T') "
    buf = buf & " and d.fecha<='" & Format(fechai, "YYYYMMDD") & "' "

    buf = buf & " and d.producto=(select a.PRODUCTO from"
    buf = buf & "(SELECT ROW_NUMBER() OVER(ORDER BY d.producto DESC) AS 'Row',d.producto"
    buf = buf & " FROM DETALLE D where D.ESTADO='2' AND (D.acu='1' or D.acu='A' or"
    buf = buf & " D.acu='B' or D.acu='C' or   D.acu='D' or D.acu='G' or D.acu='E' or D.acu='F' or D.acu='T')"
    buf = buf & " and d.fecha<='" & Format(fechai, "YYYYMMDD") & "' "
 
    If familia <> "%" Then
        buf = buf & " and D.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and D.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and D.PRODUCTO like '" & producto & "'"

    End If

    buf = buf & " group by d.producto) as a where a.Row=@intFlag))"

    buf = buf & " where codigo=(select a.PRODUCTO from (SELECT ROW_NUMBER() OVER(ORDER BY d.producto DESC) AS 'Row',d.producto"
    buf = buf & " FROM DETALLE D where D.ESTADO='2' AND (D.acu='1' or D.acu='A' or"
    buf = buf & " D.acu='B' or D.acu='C' or   D.acu='D' or D.acu='G'"
    buf = buf & " or D.acu='E' or D.acu='F' or D.acu='T')  "
    buf = buf & "  and d.fecha<='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and D.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and D.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and D.PRODUCTO like '" & producto & "'"

    End If

    buf = buf & " group by d.producto) as a"
    buf = buf & " where a.Row=@intFlag) SET @intFlag = @intFlag + 1 END"

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & " ", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteStock()
    cmdExcel.Enabled = True

    Dim buf As String

    'buf = buf & "select 'STOCK' AS 'TIP-DOC','%' as 'ALM','%' as 'ALM.DESC', '' AS DOCUMENTO,"
    'buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL','" & fechai & "'  AS 'F-EMISION',"
    '
    ' buf = buf & " '' AS 'F-VCTO','' as 'CON-PAGO',FAMILIA,A.[SUB-FAMILIA] AS 'SUBFAMILIA',"
    ' buf = buf & " ISNULL(A.CATEGORIA,'') AS 'CATEGORIA',A.CODIGO,A. ARTICULO ,"
    ' buf = buf & " A.UNIDAD  AS 'U.MED.',(a.INGRESO- b.egreso) as 'CANT.','' as PRECIO,"
    ' buf = buf & " '' as TOTAL, '' AS 'USUARIO','' as CAJA,'' as HORA,"
    'buf = buf & " '' as TURNO ,A.COSTOP AS 'COSTO' from (select D.producto AS 'CODIGO',P.DESCRIPCIO AS 'ARTICULO',P.FAMILIA AS 'FAMILIA', "
    'buf = buf & " S.DESCRIPCIO AS 'SUB-FAMILIA',D.CATEGORIA, D.UNIDAD,"
    'buf = buf & " SUM(CANTIDAD) as 'INGRESO',P.COSTOP"
    ' buf = buf & " FROM DETALLE D INNER JOIN producto p on p.producto=d.PRODUCTO"
    'buf = buf & " INNER JOIN SUBFAMIL S ON P.SUBFAMILIA=S.SUBFAMILIA"
    'buf = buf & " and d.ESTADO='2' "
    '
    'If familia <> "%" Then
    '   buf = buf & " and D.familia like '" & extra_loquesea1(familia) & "'"
    'End If
    '
    'If subfamilia <> "%" Then
    '   buf = buf & " and D.subfamilia like '" & extra_loquesea(subfamilia) & "'"
    'End If
    '
    '
    'If producto <> "%" Then
    '   buf = buf & " and D. producto like '" & producto & "'"
    'End If
    '
    '
    'buf = buf & " AND (D.acu='J' or D.acu='K' or D.acu='L' or D.acu='M' or D.acu='P'"
    'buf = buf & " or D.acu='N' or D.acu='O' or D.acu='S')"
    'buf = buf & " GROUP BY D.PRODUCTO,D.CATEGORIA,"
    'buf = buf & " P.DESCRIPCIO,P.FAMILIA ,S.DESCRIPCIO ,D.UNIDAD,P.COSTOP) as a,"
    'buf = buf & " (SELECT d.producto,SUM(CANTIDAD)  as 'EGRESO'  FROM DETALLE D INNER JOIN producto p on p.producto=d.PRODUCTO INNER JOIN SUBFAMIL S ON P.SUBFAMILIA=S.SUBFAMILIA AND  D.ESTADO='2'"
    ' buf = buf & " AND (D.acu='1' or D.acu='A' or D.acu='B' or D.acu='C' or   D.acu='D' or D.acu='G' or D.acu='E' or D.acu='F' or D.acu='T')"
    '
    ' If familia <> "%" Then
    '   buf = buf & " and D.familia like '" & extra_loquesea1(familia) & "'"
    'End If
    '
    'If subfamilia <> "%" Then
    '   buf = buf & " and D.subfamilia like '" & extra_loquesea(subfamilia) & "'"
    'End If
    '
    '
    'If producto <> "%" Then
    '   buf = buf & " and D. producto like '" & producto & "'"
    'End If
    '
    '
    '
    ' buf = buf & " and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' GROUP BY D.PRODUCTO) as b"
    '
    '
    '
    '
    'If omytablex.State = 1 Then
    'omytablex.Close
    'Set omytablex = Nothing
    'End If
    'omytablex.Open "   " & buf & "  order by A.ARTICULO", cn, adOpenStatic, adLockOptimistic
    'Set DataGrid1.DataSource = omytablex

    'buf = buf & "select 'STOCK' AS 'TIP-DOC','%' as 'ALM','%' as 'ALM.DESC', '' AS DOCUMENTO,"
    'buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL','" & fechai & "'  AS 'F-EMISION',"
    '
    ' buf = buf & " '' AS 'F-VCTO','' as 'CON-PAGO',FAMILIA,A.[SUB-FAMILIA] AS 'SUBFAMILIA',"
    ' buf = buf & " ISNULL(A.CATEGORIA,'') AS 'CATEGORIA',A.CODIGO,A. ARTICULO ,"
    ' buf = buf & " A.UNIDAD  AS 'U.MED.',(a.INGRESO) as 'CANT.','' as PRECIO,"
    ' buf = buf & " '' as TOTAL, '' AS 'USUARIO','' as CAJA,'' as HORA,"
    'buf = buf & " '' as TURNO ,A.COSTOP AS 'COSTO' from (select D.producto AS 'CODIGO',P.DESCRIPCIO AS 'ARTICULO',P.FAMILIA AS 'FAMILIA', "
    'buf = buf & " S.DESCRIPCIO AS 'SUB-FAMILIA',D.CATEGORIA, D.UNIDAD,"
    'buf = buf & " SUM(CANTIDAD) as 'INGRESO',P.COSTOP"
    ' buf = buf & " FROM DETALLE D INNER JOIN producto p on p.producto=d.PRODUCTO"
    'buf = buf & " FULL JOIN SUBFAMIL S ON P.SUBFAMILIA=S.SUBFAMILIA"
    'buf = buf & " and d.ESTADO='2' "
    '
    'If familia <> "%" Then
    '   buf = buf & " and D.familia like '" & extra_loquesea1(familia) & "'"
    'End If
    '
    'If subfamilia <> "%" Then
    '   buf = buf & " and D.subfamilia like '" & extra_loquesea(subfamilia) & "'"
    'End If
    '
    '
    'If producto <> "%" Then
    '   buf = buf & " and D. producto like '" & producto & "'"
    'End If
    '
    '
    'buf = buf & " AND (D.acu='J' or D.acu='K' or D.acu='L' or D.acu='M' or D.acu='P'"
    'buf = buf & " or D.acu='N' or D.acu='O' or D.acu='S')"
    'buf = buf & " GROUP BY D.PRODUCTO,D.CATEGORIA,"
    'buf = buf & " P.DESCRIPCIO,P.FAMILIA ,S.DESCRIPCIO ,D.UNIDAD,P.COSTOP) as a"
    '

    GeneraReporteSaldo
    buf = buf & "select 'STOCK' AS 'TIP-DOC','%' as 'ALM','%' as 'ALM.DESC', '' AS DOCUMENTO, "
    buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL','" & fechai & "'  AS 'F-EMISION', '' AS 'F-VCTO',"
    buf = buf & " '' as 'CON-PAGO',FAMILIA,a.SUBFAMILIA AS 'SUBFAMILIA', ISNULL(A.CATEGORIA,'') AS 'CATEGORIA',"
    buf = buf & " A.CODIGO,A. ARTICULO, A.UNIDAD  AS 'U.MED.',(a.INGRESO-a.salida) as 'CANT.','' as PRECIO, '' as TOTAL,"
    buf = buf & " '' AS  'USUARIO','' as CAJA,'' as HORA, '' as TURNO ,a.COSTO AS 'COSTO' from  reportes a  where codigo<>''"

    If familia <> "%" Then
        buf = buf & " and a.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and a.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and a. codigo like '" & producto & "'"

    End If

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by A.ARTICULO", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteTotal()

    ReporteUpdateClientesVentas

    cmdExcel.Enabled = True

    Dim buf As String

    'Ventas
    buf = buf & " SELECT 'VENTAS' AS 'TIP-DOC',D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO,   "
    buf = buf & " D.CODIGO AS 'COD-CLI',CL.NOMBRE AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO',   f.c9 as 'CON-PAGO', "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS 'SUBFAMILIA',D.CATEGORIA,D.PRODUCTO AS 'CODIGO',  "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.',D.PRECIO,D.TOTAL,   "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO,P.COSTOP AS 'COSTO'  "
    buf = buf & " FROM PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO   "
    buf = buf & " INNER JOIN CLIENTES CL ON D.CODIGO=CL.CODIGO   "
    buf = buf & " INNER JOIN FACTURA F ON (F.TIPO=D.TIPO AND F.SERIE=D.SERIE AND F.NUMERO=D.NUMERO)  "
    buf = buf & " AND (D.acu='1' or D.acu='A' or D.acu='B' or D.acu='C' or D.acu='D' or D.acu='G' or D.acu='E' or D.acu='F')   "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    buf = buf & " union all"

    'Compras
    buf = buf & " SELECT 'COMPRAS' AS 'TIP-DOC',D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO,   "
    buf = buf & " D.CODIGO AS 'COD-CLI',PR.NOMBRE AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO', F.fpago as 'CON-PAGO',   "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS 'SUBFAMILIA',D.CATEGORIA,D.PRODUCTO AS 'CODIGO',  "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.',D.PRECIO,D.TOTAL,   "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO,'' AS 'COSTO'    "
    buf = buf & " FROM PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO   "
    buf = buf & " INNER JOIN PROVEEDO PR  ON D.CODIGO=PR.CODIGO   "
    buf = buf & " INNER JOIN FACTURA F ON (F.TIPO=D.TIPO AND F.SERIE=D.SERIE AND F.NUMERO=D.NUMERO)   "
    buf = buf & " AND (D.acu='J' or D.acu='K' or D.acu='L' or D.acu='M' or D.acu='P' or D.acu='N' or D.acu='O')  "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    '
    buf = buf & " union all"

    'Salida

    buf = buf & " SELECT 'SALIDA' AS 'TIP-DOC', D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO, "
    buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', '' AS 'F-VCTO', '' as 'CON-PAGO', "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS 'SUBFAMILIA',D.CATEGORIA,D.PRODUCTO AS 'CODIGO', "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.','' AS PRECIO,'' AS TOTAL, "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO ,'' AS 'COSTO'"
    buf = buf & " FROM PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO AND (D.acu='T' ) "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    buf = buf & " union all"

    'entrada

    buf = buf & " SELECT 'ENTRADA' AS 'TIP-DOC', D.BODEGA AS ALM, B.NOMBRE AS 'ALM.DESC',(D.TIPO+'-'+D.SERIE+'-'+D.NUMERO) AS DOCUMENTO, "
    buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL',D.FECHACREA  AS 'F-EMISION', '' AS 'F-VCTO','' as 'CON-PAGO',  "
    buf = buf & " D.FAMILIA,D.SUBFAMILIA AS 'SUBFAMILIA',D.CATEGORIA,D.PRODUCTO AS 'CODIGO', "
    buf = buf & " D.DESCRIPCIO AS 'ARTICULO',D.UNIDAD AS 'U.MED.',D.CANTIDAD AS 'CANT.','' AS PRECIO,'' AS TOTAL, "
    buf = buf & " D.USUARIO AS 'USUARIO',D.CAJA,D.HORA,D.TURNO ,'' AS 'COSTO' "
    buf = buf & " FROM PRODUCTO P INNER JOIN DETALLE D  ON P.PRODUCTO=D.PRODUCTO INNER JOIN BODEGA B ON D.BODEGA=B.CODIGO AND (D.acu='S' ) "
    buf = buf & " and d.estado='2' and d.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and d.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and d.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and d.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and d.producto like '" & producto & "'"

    End If

    buf = buf & " union all "

    'STOCK

    GeneraReporteSaldo
    buf = buf & "select 'STOCK' AS 'TIP-DOC','%' as 'ALM','%' as 'ALM.DESC', '' AS DOCUMENTO, "
    buf = buf & " '' AS 'COD-CLI','' AS 'RAZON SOCIAL','" & fechai & "'  AS 'F-EMISION', '' AS 'F-VCTO',"
    buf = buf & " '' as 'CON-PAGO',FAMILIA,a.SUBFAMILIA AS 'SUBFAMILIA', ISNULL(A.CATEGORIA,'') AS 'CATEGORIA',"
    buf = buf & " A.CODIGO,A. ARTICULO, A.UNIDAD  AS 'U.MED.',(a.INGRESO-a.salida) as 'CANT.','' as PRECIO, '' as TOTAL,"
    buf = buf & " '' AS  'USUARIO','' as CAJA,'' as HORA, '' as TURNO ,a.COSTO AS 'COSTO' from  reportes a  where codigo<>''"

    If familia <> "%" Then
        buf = buf & " and a.familia like '" & extra_loquesea1(familia) & "'"

    End If

    If subfamilia <> "%" Then
        buf = buf & " and a.subfamilia like '" & extra_loquesea(subfamilia) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and a. codigo like '" & producto & "'"

    End If

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by 'TIP-DOC'", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Private Sub cmdrefresca_Click()

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        fechaf.SetFocus
        Exit Sub

    End If

    If Len(fechai) = 0 Then
        MsgBox "Fechai No valida", 48, "Aviso"
        Exit Sub

    End If

    If Len(fechai) <> 10 Then
        MsgBox "Fechai No valida", 48, "Aviso"
        Exit Sub

    End If

    If CboOpcion.ListIndex = 0 Then ReporteTotal
    If CboOpcion.ListIndex = 1 Then ReporteVentas
    If CboOpcion.ListIndex = 2 Then ReporteCompras
    If CboOpcion.ListIndex = 3 Then ReporteSalida
    If CboOpcion.ListIndex = 4 Then ReporteEntrada
    If CboOpcion.ListIndex = 5 Then ReporteStock

End Sub

Private Sub familia_Change()

    If extra_loquesea1(familia) <> "%" Then
        carga_subfamilia

    End If

End Sub

Sub carga_subfamilia()

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    subfamilia.Clear
    subfamilia.AddItem "%"
    cad = "SELECT * FROM subfamil where familia='" & extra_loquesea1(familia) & "' order by subfamilia "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subfamilia.AddItem "" & mytablex.Fields("subfamilia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    subfamilia.ListIndex = 0
    mytablex.Close

End Sub

Private Sub Form_Activate()

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    familia.Clear
    familia.AddItem "%"

    cad = "SELECT * FROM FAMILIA  order by descripcio "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        familia.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & mytablex.Fields("familia")
        mytablex.MoveNext
    Loop
    familia.ListIndex = 0
    mytablex.Close

    Set mytablex = Nothing

    subfamilia.Clear
    subfamilia.AddItem "%"

    cad = "SELECT * FROM subfamil  order by subfamilia "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        subfamilia.AddItem "" & mytablex.Fields("subfamilia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    subfamilia.ListIndex = 0
    mytablex.Close

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

    CboOpcion.Clear
    CboOpcion.AddItem "%"
    CboOpcion.AddItem "VENTAS"
    CboOpcion.AddItem "COMPRAS"
    CboOpcion.AddItem "SALIDA"
    CboOpcion.AddItem "ENTRADA"
    CboOpcion.AddItem "STOCK"
    CboOpcion.ListIndex = 0

End Sub

