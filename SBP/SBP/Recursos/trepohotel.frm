VERSION 5.00
Begin VB.Form trepohotel 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Hotel"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox estado 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox habitacion 
      Height          =   375
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   11
      Text            =   "%"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   2040
      MaxLength       =   11
      TabIndex        =   8
      Text            =   "%"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox nombre 
      Height          =   375
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "%"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox categoria 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GenerarReporte"
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   210
      Width           =   1695
   End
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Habitacion"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Categoria"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trepohotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    buf = "select * from hotelcheckin where "
    buf = buf & "  arribofecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and arribofecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If categoria <> "%" Then
        buf = buf & " and categoria='" & categoria.Text & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo.Text & "%'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre.Text & "%'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado.Text & "'"

    End If

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    reporte_excell mytablex

End Sub

Private Sub flo44_Click()
    trepohotel.Hide
    Unload trepohotel

End Sub

Private Sub Form_Load()
    categoria.Clear
    categoria.AddItem "%"
    categoria.AddItem "NOCHES"
    categoria.AddItem "HORAS"
    categoria.ListIndex = 0
    fechai = Format(Now, "dd/mm/yyyy") '"01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

    estado.Clear
    estado.AddItem "%"
    estado.AddItem "ENTRADA"
    estado.AddItem "RESERVA"
    estado.AddItem "CERRADO"

    estado.ListIndex = 0

End Sub

Sub reporte_excell(mytablex As ADODB.Recordset)

    Dim xhoy        As String

    Dim dias        As Integer

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim xtotal      As Double

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Dim Fecha1      As Date

    Dim Fecha2      As Date

    Dim meses       As Integer

    Dim mytabley    As New ADODB.Recordset

    Dim xconsumo    As Double

    Dim xabono      As Double

    Dim xsaldo      As Double

    Dim xneto       As Double

    Dim xxtotal     As Double

    Dim xxconsumo   As Double

    Dim xxabono     As Double

    Dim xxsaldo     As Double

    Dim xxneto      As Double

    xxtotal = 0
    xxconsumo = 0
    xxabono = 0
    xxsaldo = 0
    xxneto = 0

    Command1.Visible = True

    On Error GoTo cmd6561245_err
    
    Heading(1) = "Habitacion"
    Heading(2) = "FechaEnt."
    Heading(3) = "FechaSal."
    Heading(4) = "H.Ingreso"
    Heading(5) = "H.Salida"
    Heading(6) = "Apellidos y Nombres"
    Heading(7) = "Doc.Ident"
    Heading(8) = "Categoria"
    Heading(9) = "NroDias"
    Heading(10) = "Hospedaje"
    Heading(11) = "Consumo"
    Heading(12) = "Total"
    Heading(13) = "Abono"
    Heading(14) = "Saldo"
    Heading(15) = "Estado"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook

    objExcel.ActiveSheet.Cells(1, 1) = "FECHA HOY  " + Format(Now, "dd/mm/yyyy") & "- HORA HOY  " + Format(Now, "HH:MM:SS")
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO :" + Format(fechai, "DD/MM/YYYY") & " FECHA FINAL :" + Format(fechaf, "DD/MM/YYYY")

    v = 4
    h = 1
    sdx1 = 0
    
    Do

        If mytablex.EOF Then Exit Do
        'objExcel.ActiveSheet.Cells(v, H) = "'" & mytablex.Fields("categoria")
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("habitacion")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("arribofecha")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("arribofechaf")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("arribohora")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("arribohoraf")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("hnombre")
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytablex.Fields("huesped")
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("categoria")

        'dias ocupadas

        If "" & mytablex.Fields("estado") = "ENTRADA" Then
            dias = 1

            If IsDate("" & mytablex.Fields("arribofecha")) Then
                xhoy = Format(Now, "dd/mm/yyyy")
                dias = DateDiff("d", Trim("" & mytablex.Fields("arribofecha")), xhoy)

                If dias = 0 Then
                    dias = 1

                End If

            End If

            objExcel.ActiveSheet.Cells(v, h + 8) = dias
            sdx = Val("" & mytablex.Fields("precio")) * dias
            xtotal = Val(Format(sdx, "0.00"))
            objExcel.ActiveSheet.Cells(v, h + 9) = xtotal
            xxtotal = xxtotal + xtotal
   
        End If
 
        If "" & mytablex.Fields("estado") = "RESERVA" Then
            dias = 1

            If IsDate("" & mytablex.Fields("arribofecha")) Then
                xhoy = Format(Now, "dd/mm/yyyy")
                dias = DateDiff("d", Trim("" & mytablex.Fields("arribofecha")), xhoy)

                If dias = 0 Then
                    dias = 1

                End If

            End If

            objExcel.ActiveSheet.Cells(v, h + 8) = dias
            sdx = 0 * dias
            xtotal = Val(Format(sdx, "0.00"))
            objExcel.ActiveSheet.Cells(v, h + 9) = xtotal
            xxtotal = xxtotal + xtotal
   
        End If

        If "" & mytablex.Fields("estado") = "CERRADO" Then
            dias = 1

            If IsDate("" & mytablex.Fields("arribofecha")) Then
                xhoy = Format("" & mytablex.Fields("arribofechaf"), "dd/mm/yyyy")
                dias = DateDiff("d", Trim("" & mytablex.Fields("arribofecha")), xhoy)

                If dias = 0 Then
                    dias = 1

                End If

            End If

            objExcel.ActiveSheet.Cells(v, h + 8) = dias
            sdx = Val("" & mytablex.Fields("precio")) * dias
            xtotal = Val(Format(sdx, "0.00"))
            objExcel.ActiveSheet.Cells(v, h + 9) = xtotal
            xxtotal = xxtotal + xtotal
   
        End If
 
        'consumos

        sdx = 0
        mytabley.Open "select * from hotelconsumo where idecheckin=" & Val("" & mytablex.Fields("checkin")), cn, adOpenStatic, adLockOptimistic
        Do

            If mytabley.EOF Then Exit Do
            sdx = sdx + Val("" & mytabley.Fields("total"))
            mytabley.MoveNext
        Loop
        xconsumo = Val(Format(sdx, "0.00"))
        objExcel.ActiveSheet.Cells(v, h + 10) = xconsumo
        mytabley.Close
        xxconsumo = xxconsumo + xconsumo

        xneto = xconsumo + xtotal
        objExcel.ActiveSheet.Cells(v, h + 11) = xneto
        xxneto = xxneto + xneto
        'abonos

        mytabley.Open "select * from hotelfactura where idcheckin=" & Val("" & mytablex.Fields("checkin")), cn, adOpenStatic, adLockOptimistic
        Do

            If mytabley.EOF Then Exit Do
            sdx = sdx + Val("" & mytabley.Fields("total"))
            mytabley.MoveNext
        Loop

        xabono = Val(Format(sdx, "0.00"))
        xxabono = xxabono + xabono
        mytabley.Close
        objExcel.ActiveSheet.Cells(v, h + 12) = xabono
        xsaldo = xneto - xabono
        objExcel.ActiveSheet.Cells(v, h + 13) = xsaldo
        objExcel.ActiveSheet.Cells(v, h + 14) = Trim("" & mytablex.Fields("estado"))
        xxsaldo = xxsaldo + xsaldo

        v = v + 1
        mytablex.MoveNext
    Loop

    objExcel.ActiveSheet.Cells(v, h + 9) = xxtotal
    objExcel.ActiveSheet.Cells(v, h + 10) = xxconsumo
    objExcel.ActiveSheet.Cells(v, h + 11) = xxneto
    objExcel.ActiveSheet.Cells(v, h + 12) = xxabono
    objExcel.ActiveSheet.Cells(v, h + 13) = xxsaldo

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
       
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 10
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        .columns("F").ColumnWidth = 30
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 10
        .columns("i").ColumnWidth = 10
        .columns("j").ColumnWidth = 10
        .columns("k").ColumnWidth = 10
        .columns("l").ColumnWidth = 10
    
    End With

End Function

