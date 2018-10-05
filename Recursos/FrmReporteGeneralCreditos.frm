VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmReporteGeneralCreditos 
   Caption         =   "Reporte Creditos"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15015
      Begin VB.TextBox fechai 
         Height          =   495
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   495
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   5
         Top             =   960
         Width           =   1815
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
         TabIndex        =   4
         Top             =   1800
         Width           =   2895
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
         TabIndex        =   3
         Text            =   "%"
         Top             =   3240
         Width           =   1575
      End
      Begin ChamaleonButton.ChameleonBtn ChaGENERAR 
         Height          =   825
         Left            =   6360
         TabIndex        =   7
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
         MICON           =   "FrmReporteGeneralCreditos.frx":0000
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
         Left            =   6360
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
         MICON           =   "FrmReporteGeneralCreditos.frx":001C
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
         Left            =   13320
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
         MICON           =   "FrmReporteGeneralCreditos.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Items 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "items"
         Height          =   195
         Left            =   14040
         TabIndex        =   14
         Top             =   2160
         Width           =   855
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
         TabIndex        =   13
         Top             =   360
         Width           =   1590
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1545
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1005
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
   End
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   15015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
End
Attribute VB_Name = "FrmReporteGeneralCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim omytablex As New ADODB.Recordset

' -- Variables para la base de datos
'Dim cnn         As Connection
'Dim rs          As Recordset
' -- Variables para Excel
Dim Obj_Excel As Object

Dim Obj_Libro As Object

Dim Obj_Hoja  As Object

Private Sub ChaCERRAR_Click()
    FrmReporteGeneralCreditos.Hide
    Unload FrmReporteGeneralCreditos

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
    If CboOpcion.ListIndex = 1 Then ReporteCreditoPorCobrar
    If CboOpcion.ListIndex = 2 Then ReporteCreditoPorPagar

    items = 0
    items = DataGrid1.VisibleRows
    reporte_excell omytablex

End Sub

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 14)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, 14)).Interior.color = RGB(192, 192, 250)

        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        .columns("A").ColumnWidth = 18
        .columns("B").ColumnWidth = 14
        .columns("C").ColumnWidth = 25
        .columns("D").ColumnWidth = 12
        .columns("E").ColumnWidth = 12
        .columns("F").ColumnWidth = 8
        .columns("G").ColumnWidth = 11
        .columns("H").ColumnWidth = 12
        .columns("I").ColumnWidth = 12 ' CONDICION DE PAGO
        .columns("J").ColumnWidth = 12
        .columns("K").ColumnWidth = 11
        .columns("L").ColumnWidth = 11
        .columns("M").ColumnWidth = 11
        .columns("N").ColumnWidth = 20
            
    End With

End Function

Sub Formato_ordenTotal(I As Integer)

    ' objExcel.ActiveSheet.Cells(i, 7).Font.bold = True
    ' objExcel.ActiveSheet.Cells(i, 7).Interior.Color = RGB(248, 243, 53)
    '
End Sub

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

    Heading(1) = "TIP-DOC"
    Heading(2) = "DOCUMENTO"
    Heading(3) = "RAZON SOCIAL"
    Heading(4) = "F-EMISION"
    Heading(5) = "F-VCTO"
    Heading(6) = "DIAS"
    Heading(7) = "CON-PAGO"
    Heading(8) = "DEUDA TOTAL"
    Heading(9) = "ABONO"
    Heading(10) = "TOTAL SALDO"
    Heading(11) = "STATUS"
    Heading(12) = "VENDEDOR"
    Heading(13) = "EMPRESA"
    Heading(14) = "FECHA DE PROCESO"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(14, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    objExcel.ActiveSheet.Cells(1, 1) = "REPORTE >>>>>>>"
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA " + Format(Now, "HH:MM:SS")

    v = 4
    h = 1
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
    
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("TIP-DOC")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("DOCUMENTO")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("RAZON SOCIAL")

        If mytablex.Fields("F-EMISION") = "01/01/1900" Then
            objExcel.ActiveSheet.Cells(v, h + 3) = ""
        Else
            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("F-EMISION")

        End If
    
        If mytablex.Fields("F-VCTO") = "01/01/1900" Then
            objExcel.ActiveSheet.Cells(v, h + 4) = ""
        Else
            objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("F-VCTO")

        End If
      
        If mytablex.Fields("DIAS") <= 0 Then
            objExcel.ActiveSheet.Cells(v, h + 5) = "-"
        Else
            objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("DIAS")

        End If

        ' MUESTRA CONDICION DE PAGO SEGUN FPAGO
    
        If mytablex.Fields("CON-PAGO") = "3" Then  'CREDITO
            objExcel.ActiveSheet.Cells(v, h + 6) = "CREDITO"
        ElseIf mytablex.Fields("CON-PAGO") = "" Then
            objExcel.ActiveSheet.Cells(v, h + 6) = ""
        Else
            objExcel.ActiveSheet.Cells(v, h + 6) = "CONTADO"

        End If

        ' MUESTRA CONDICION DE PAGO SEGUN FPAGO

        objExcel.ActiveSheet.Cells(v, h + 7) = mytablex.Fields("DEUDA TOTAL")
        objExcel.ActiveSheet.Cells(v, h + 8) = mytablex.Fields("ABONO")
        objExcel.ActiveSheet.Cells(v, h + 9) = mytablex.Fields("TOTAL SALDO")
   
        If mytablex.Fields("TOTAL SALDO") = 0 Then
            objExcel.ActiveSheet.Cells(v, h + 10) = "CANCELADO"
   
        Else
            objExcel.ActiveSheet.Cells(v, h + 10) = "POR COBRAR"
   
        End If
   
        objExcel.ActiveSheet.Cells(v, h + 11) = "'" & mytablex.Fields("VENDEDOR")
        objExcel.ActiveSheet.Cells(v, h + 12) = "'" & mytablex.Fields("EMPRESA")
        objExcel.ActiveSheet.Cells(v, h + 13) = "'" & mytablex.Fields("FECHA DE PROCESO")
    
        objExcel.ActiveSheet.Cells(v + 1, 7) = "NETO"
        
        '  .Range(.Cells(items + 4, 1), .Cells(items + 4, 14)).Interior.Color = RGB(248, 243, 53)
        
        sumad = sumad + Val("" & mytablex.Fields("DEUDA TOTAL"))
        sumaa = sumaa + Val("" & mytablex.Fields("ABONO"))
        sumat = sumat + Val("" & mytablex.Fields("TOTAL SALDO"))
 
        objExcel.ActiveSheet.Cells(v + 1, 8) = sumad
        objExcel.ActiveSheet.Cells(v + 1, 9) = sumaa
        objExcel.ActiveSheet.Cells(v + 1, 10) = sumat
    
        v = v + 1
   
        'imprime_recetaa mytablex, v, h
        mytablex.MoveNext
    Loop
 
    For I = 1 To 14
        objExcel.ActiveSheet.Cells(v, I).Font.bold = True
        objExcel.ActiveSheet.Cells(v, I).Interior.color = RGB(248, 243, 53)
    Next
 
    Set objExcel = Nothing

    'Exit Sub
    'cmd6561245_err:
    'MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    'Exit Sub
End Sub

Sub ReporteCreditoPorCobrar()

    cmdExcel.Enabled = True

    Dim buf As String

    buf = buf & "  SELECT 'CRED. POR COBRAR' AS 'TIP-DOC',  (F.TIPO+'-'+F.SERIE+'-'+F.NUMERO) AS DOCUMENTO,"
    buf = buf & "  CL.NOMBRE AS 'RAZON SOCIAL',F.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO', DATEDIFF(day, fechae, getdate()) as DIAS ,"
    buf = buf & " fp.fpago as 'CON-PAGO', CC.TOTAL as  'DEUDA TOTAL' , CC.ABONO as  'ABONO' ,        CC.SALDO as  'TOTAL SALDO' ,"
    buf = buf & "  '-' as  'STATUS' ,  f.USUARIO AS 'VENDEDOR',  '01' AS EMPRESA,  CONVERT(VARCHAR(10), GETDATE(), 103) AS 'FECHA DE PROCESO'"
    buf = buf & "  FROM cuentac CC INNER JOIN factura f ON (CC.TIPO=F.TIPO AND CC.SERIE=F.SERIE AND       CC.NUMERO=f.NUMERO)"
    buf = buf & "   INNER JOIN BODEGA B ON F.BODEGA=B.CODIGO      INNER JOIN CLIENTES CL"
    buf = buf & "  ON f.CODIGO=CL.CODIGO inner join fpagov fp on (Fp.TIPO=F.TIPO AND Fp.SERIE=F.SERIE AND"
    buf = buf & "  Fp.NUMERO=f.NUMERO)  AND (f.acu='1' or F.acu='A' or F.acu='B' or F.acu='C'"
    buf = buf & " or F.acu='D' or F.acu='G' or F.acu='E' or F.acu='F')"
    buf = buf & " and f.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and f.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by f.fecha,f.hora", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteCreditoPorPagar()

    cmdExcel.Enabled = True

    Dim buf As String

    buf = buf & "  SELECT 'DEUDA POR PAGAR' AS 'TIP-DOC',  (F.TIPO+'-'+F.SERIE+'-'+F.NUMERO) AS DOCUMENTO,"
    buf = buf & "  PR.NOMBRE AS 'RAZON SOCIAL',F.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO', DATEDIFF(day, fechae, getdate()) as DIAS ,"
    buf = buf & " fp.fpago as 'CON-PAGO', CC.TOTAL as  'DEUDA TOTAL' , CC.ABONO as  'ABONO' ,        CC.SALDO as  'TOTAL SALDO' ,"
    buf = buf & "  '-' as  'STATUS' ,  f.USUARIO AS 'VENDEDOR',  '01' AS EMPRESA,  CONVERT(VARCHAR(10), GETDATE(), 103) AS 'FECHA DE PROCESO'"
    buf = buf & "  FROM cuentap CC INNER JOIN factura f ON (CC.TIPO=F.TIPO AND CC.SERIE=F.SERIE AND       CC.NUMERO=f.NUMERO)"
    buf = buf & "   INNER JOIN BODEGA B ON F.BODEGA=B.CODIGO  INNER JOIN PROVEEDO PR"
    buf = buf & "  ON f.CODIGO=PR.CODIGO inner join fpagov fp on (Fp.TIPO=F.TIPO AND Fp.SERIE=F.SERIE AND FP.NUMERO=F.NUMERO) AND "
    buf = buf & "  (F.acu='J' or F.acu='K' or F.acu='L' or F.acu='M' or F.acu='P' or F.acu='N' or F.acu='O')  "
    buf = buf & " and f.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and f.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by f.fecha,f.hora", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Sub ReporteTotal()

    cmdExcel.Enabled = True

    Dim buf As String

    'Credito por Cobrar

    buf = buf & "  SELECT 'CRED. POR COBRAR' AS 'TIP-DOC',  (F.TIPO+'-'+F.SERIE+'-'+F.NUMERO) AS DOCUMENTO,"
    buf = buf & "  CL.NOMBRE AS 'RAZON SOCIAL',F.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO', DATEDIFF(day, fechae, getdate()) as DIAS ,"
    buf = buf & " fp.fpago as 'CON-PAGO', CC.TOTAL as  'DEUDA TOTAL' , CC.ABONO as  'ABONO' ,        CC.SALDO as  'TOTAL SALDO' ,"
    buf = buf & "  '-' as  'STATUS' ,  f.USUARIO AS 'VENDEDOR',  '01' AS EMPRESA,  CONVERT(VARCHAR(10), GETDATE(), 103) AS 'FECHA DE PROCESO'"
    buf = buf & "  FROM cuentac CC INNER JOIN factura f ON (CC.TIPO=F.TIPO AND CC.SERIE=F.SERIE AND       CC.NUMERO=f.NUMERO)"
    buf = buf & "   INNER JOIN BODEGA B ON F.BODEGA=B.CODIGO      INNER JOIN CLIENTES CL"
    buf = buf & "  ON f.CODIGO=CL.CODIGO inner join fpagov fp on (Fp.TIPO=F.TIPO AND Fp.SERIE=F.SERIE AND"
    buf = buf & "  Fp.NUMERO=f.NUMERO)  AND (f.acu='1' or F.acu='A' or F.acu='B' or F.acu='C'"
    buf = buf & " or F.acu='D' or F.acu='G' or F.acu='E' or F.acu='F')"
    buf = buf & " and f.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and f.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    buf = buf & " union all"

    'Por Pagar

    buf = buf & "  SELECT 'DEUDA POR PAGAR' AS 'TIP-DOC',  (F.TIPO+'-'+F.SERIE+'-'+F.NUMERO) AS DOCUMENTO,"
    buf = buf & "  PR.NOMBRE AS 'RAZON SOCIAL',F.FECHACREA  AS 'F-EMISION', F.FECHAE AS 'F-VCTO', DATEDIFF(day, fechae, getdate()) as DIAS ,"
    buf = buf & " fp.fpago as 'CON-PAGO', CC.TOTAL as  'DEUDA TOTAL' , CC.ABONO as  'ABONO' ,        CC.SALDO as  'TOTAL SALDO' ,"
    buf = buf & "  '-' as  'STATUS' ,  f.USUARIO AS 'VENDEDOR',  '01' AS EMPRESA,  CONVERT(VARCHAR(10), GETDATE(), 103) AS 'FECHA DE PROCESO'"
    buf = buf & "  FROM cuentap CC INNER JOIN factura f ON (CC.TIPO=F.TIPO AND CC.SERIE=F.SERIE AND       CC.NUMERO=f.NUMERO)"
    buf = buf & "   INNER JOIN BODEGA B ON F.BODEGA=B.CODIGO      INNER JOIN PROVEEDO PR"
    buf = buf & "  ON f.CODIGO=PR.CODIGO inner join fpagov fp on (Fp.TIPO=F.TIPO AND Fp.SERIE=F.SERIE AND FP.NUMERO=F.NUMERO) AND "
    buf = buf & "  (F.acu='J' or F.acu='K' or F.acu='L' or F.acu='M' or F.acu='P' or F.acu='N' or F.acu='O')  "
    buf = buf & " and f.fecha<='" & Format(fechaf, "YYYYMMDD") & "' and f.fecha>='" & Format(fechai, "YYYYMMDD") & "' "

    If omytablex.State = 1 Then
        omytablex.Close
        Set omytablex = Nothing

    End If

    omytablex.Open "   " & buf & "  order by 'TIP-DOC'", cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = omytablex

End Sub

Private Sub ChaGENERAR_Click()

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
    If CboOpcion.ListIndex = 1 Then ReporteCreditoPorCobrar
    If CboOpcion.ListIndex = 2 Then ReporteCreditoPorPagar

    items = 0
    items = DataGrid1.VisibleRows

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

    CboOpcion.Clear
    CboOpcion.AddItem "%"
    CboOpcion.AddItem "CREDITO POR COBRAR"
    CboOpcion.AddItem "CREDITO POR PAGAR"
    CboOpcion.ListIndex = 0

End Sub

