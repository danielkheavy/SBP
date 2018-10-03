VERSION 5.00
Begin VB.Form Ejemplo 
   Caption         =   "Ejemplo de paso de datos de  Bd Access a Excel"
   ClientHeight    =   3210
   ClientLeft      =   1590
   ClientTop       =   1575
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   7200
   Begin VB.CommandButton ALaCalle 
      Caption         =   "&Salir"
      Height          =   465
      Left            =   2835
      TabIndex        =   1
      Top             =   2475
      Width           =   1635
   End
   Begin VB.CommandButton Crear 
      Caption         =   "&CrearHoja"
      Height          =   690
      Left            =   2160
      TabIndex        =   0
      Top             =   1620
      Width           =   2985
   End
   Begin VB.Data Dtickets 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4995
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label2 
      Caption         =   "Creo que es la forma más sencilla de trabajar."
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   1170
      Width           =   6990
   End
   Begin VB.Label Label1 
      Caption         =   $"BdAExcel.frx":0000
      Height          =   735
      Left            =   180
      TabIndex        =   2
      Top             =   315
      Width           =   6990
   End
End
Attribute VB_Name = "Ejemplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ApExcel As Excel.Application
Dim Up As String

Private Sub ALaCalle_Click()
Unload Me
End
End Sub

Private Sub Crear_Click()
 HojaStock
End Sub

Private Sub Form_Load()
On Error Resume Next
    MkDir "C:\PRUEBAS"
    Up = "C:\PRUEBAS\"
    ChDir "C:\PRUEBAS"
    Dtickets.DatabaseName = Up + "DATOS.MDB"
    Dtickets.RecordSource = "SELECT * FROM TBARTICU ORDER BY CODIGO;"
    Dtickets.Refresh
   
End Sub
Sub HojaStock()
Dim Lin As Long
Dim OldLin As Long
Dim x As Integer
Dim i As Integer
Dim E As Integer

    E = 0
    On Error GoTo ErOpen
    Set ApExcel = New Excel.Application
    If E <> 0 Then End
    E = 0
    ApExcel.Workbooks.Open (Up + "HOJASTOCK_ORIGINAL.XLS")
    On Error GoTo 0
    If E <> 0 Then
        ApExcel.Application.Workbooks.Add
        ApExcel.Visible = True
        ApExcel.Sheets(1).Select
        With ApExcel
            .Application.Calculation = xlManual
'           .Application.ActivePrinter = Printer.DeviceName & " en LPT1:"
            With .ActiveSheet.PageSetup
                .PrintTitleRows = "$1:$3"
                .PrintTitleColumns = "$A:$H"
                .PrintArea = ""
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = "Página &P"
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0)
                .RightMargin = Application.InchesToPoints(0)
                .TopMargin = Application.InchesToPoints(0)
                .BottomMargin = Application.InchesToPoints(0)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = 100
            End With
            .Cells.Select
            With .Selection.Font
                .Name = "Arial"
                .Size = 8
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
            End With
            .Range("A3").FormulaR1C1 = "Referencia"
            .Range("B3").FormulaR1C1 = "Descripción"
            .Range("C3").FormulaR1C1 = "P.Costo"
            .Range("D3").FormulaR1C1 = "P.V.P."
            .Range("E3").FormulaR1C1 = "Dto."
            .Range("F3").FormulaR1C1 = "Stock"
            .Range("G3").FormulaR1C1 = "ValorCosto"
            .Range("H3").FormulaR1C1 = "Colores"
            
            .Range("A3:H3").Select
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Negrita"
                .Size = 10
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
            End With
            .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With .Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
        ApExcel.ActiveWorkbook.SaveAs FileName:=Up + "HOJASTOCK_ORIGINAL.XLS", _
            FileFormat:=xlNormal, Password:="", _
            WriteResPassword:="", ReadOnlyRecommended:=False, _
            CreateBackup:=False

     End If
    
    ApExcel.Visible = True
    ApExcel.Sheets(1).Select
    With ApExcel
    .Application.Calculation = xlManual
    .Application.ScreenUpdating = False
    .Columns("C:D").Select
    .Selection.NumberFormat = "#,##0.00"
    .Columns("G:G").Select
    .Selection.NumberFormat = "#,##0.00"

'On Error GoTo 0
    Lin = 4
    If Dtickets.Recordset.EOF = False Then
        Dtickets.Recordset.MoveLast
        Dtickets.Recordset.MoveFirst
    End If
    If Dtickets.Recordset.BOF = False Then Dtickets.Recordset.MoveFirst
    For i = 1 To Dtickets.Recordset.RecordCount
        If Dtickets.Recordset.Fields("Stock") <> 0 Then
            If (i Mod 14) = 0 Then .Range("A" + Mid$(Str$(i), 2)).Select
            .Cells(Lin, 1) = Dtickets.Recordset.Fields("Codigo")
            .Cells(Lin, 2) = Dtickets.Recordset.Fields("Nombre")
            .Cells(Lin, 3) = Dtickets.Recordset.Fields("PrecioCosto") / 100
          '  .Cells(Lin, 3).NumberFormat = "#,##0.00"
            .Cells(Lin, 4) = Dtickets.Recordset.Fields("Pvp1") / 100
'            .Cells(Lin, 4).NumberFormat = "#,##0.00"
            .Cells(Lin, 5) = Dtickets.Recordset.Fields("Descuento")
            .Cells(Lin, 6) = Dtickets.Recordset.Fields("Stock")
            .Cells(Lin, 7).FormulaR1C1 = "=RC[-1]*RC[-4]"
 '           .Cells(Lin, 7).NumberFormat = "#,##0.00"
            .Cells(Lin, 8) = Dtickets.Recordset.Fields("Colores")
'            .Cells(Lin, 9) = DTickets.Recordset.Fields("Observaciones")
            Lin = Lin + 1
        End If
        Dtickets.Recordset.MoveNext
    Next i
    .Application.Calculation = xlAutomatic
    .Application.ScreenUpdating = True

    OldLin = Lin
    .Cells(Lin, 7).NumberFormat = "#,##0.00"
    .Cells(Lin, 7).FormulaR1C1 = "=RC[-1]*RC[-4]"
    .Cells(Lin, 7).NumberFormat = "#,##0.00"
      
    
    .Columns("A:A").EntireColumn.AutoFit
    .Columns("B:B").EntireColumn.AutoFit
    .Columns("C:C").EntireColumn.AutoFit
    .Columns("D:D").EntireColumn.AutoFit
    .Columns("E:E").EntireColumn.AutoFit
    .Columns("F:F").EntireColumn.AutoFit
    .Columns("G:G").EntireColumn.AutoFit
    .Columns("H:H").EntireColumn.AutoFit
    
      
    .Range("A3:H" + Mid$(Str$(Lin), 2)).Select
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    
    
    .Range("F" + Mid$(Str$(Lin), 2)).FormulaR1C1 = "=SUM(R[-" + Mid$(Str$(Lin - 4), 2) + "]C:R[-1]C)"
    .Range("G" + Mid$(Str$(Lin), 2)).FormulaR1C1 = "=SUM(R[-" + Mid$(Str$(Lin - 4), 2) + "]C:R[-1]C)"
    .Range("A" + Mid$(Str$(Lin), 2) + ":H" + Mid$(Str$(Lin), 2)).Select
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    
    .Range("B" + Mid$(Str$(Lin), 2)).FormulaR1C1 = "TOTALES......: "
    .Range("B" + Mid$(Str$(Lin), 2) + ":G" + Mid$(Str$(Lin + 1), 2)).Font.Bold = True
    With .Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    .Range("B" + Mid$(Str$(Lin + 1), 2)).FormulaR1C1 = " Fecha:  " + Mid$(Date$, 4, 3) + Left$(Date$, 3) + Right$(Date$, 4)
        
        .Range("A1").FormulaR1C1 = "jagrane@yahoo.es"
        .Range("E1").FormulaR1C1 = "RELACION DE EXISTENCIAS"
        .Rows("1:1").Select
        .Selection.Font.Bold = False
        .Selection.Font.Bold = True
        With .Selection.Font
            .Name = "Arial"
            .Size = 11
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
    End With
    Open Up + "STOCKREL.xls" For Random As #1
    Close #1
    Kill Up + "STOCKREL.xls"
    ApExcel.ActiveWorkbook.SaveAs FileName:=Up + "STOCKREL.xls", FileFormat:= _
            xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
            , CreateBackup:=False
             
    ApExcel.ActiveWindow.Close
    ApExcel.Quit
    Set ApExcel = Nothing
      Exit Sub
ErOpen:
    E = 1
    Resume Next



End Sub
