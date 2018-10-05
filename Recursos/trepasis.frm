VERSION 5.00
Begin VB.Form trepasis 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Asistencia"
   ClientHeight    =   3240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Nombre 
      Height          =   495
      Left            =   1680
      MaxLength       =   11
      TabIndex        =   6
      Text            =   "%"
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox codigo 
      Height          =   495
      Left            =   1680
      MaxLength       =   11
      TabIndex        =   1
      Text            =   "%"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenado por"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu dj7822 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu dflo8922 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trepasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dflo8922_Click()
    trepasis.Hide
    Unload trepasis

End Sub

Private Sub dj7822_Click()
    exporta_excel

End Sub

Function calcula_hora(mare2 As String, mare1 As String) As String

    Dim horax As String

    On Error GoTo cmd341_err

    horax = ""

    If Len(mare2) = 0 Then Exit Function
    If Len(mare1) = 0 Then Exit Function
    If Val(Mid$(mare2, 1, 2)) >= 0 And Val(Mid$(mare2, 1, 2)) <= 24 Then
        If Val(Mid$(mare2, 4, 2)) >= 0 And Val(Mid$(mare2, 4, 2)) <= 59 Then
            If Val(Mid$(mare1, 1, 2)) >= 0 And Val(Mid$(mare1, 1, 2)) <= 24 Then
                If Val(Mid$(mare1, 4, 2)) >= 0 And Val(Mid$(mare1, 4, 2)) <= 59 Then
                    horax = Format(TimeValue(mare2) - TimeValue(mare1), "hh:mm")

                End If

            End If

        End If

    End If

    calcula_hora = horax
    Exit Function
cmd341_err:
    Exit Function

End Function

Function busca_tipo(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Fecha"
    Combo1.ListIndex = 0
    fechai = "01/" + Format(Month(Now), "00") + "/" + Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Sub exporta_excel()

    Dim v           As Long

    Dim h           As Long

    Dim I           As Long

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim mytablex    As New ADODB.Recordset

    Dim strSQL      As String

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim buf         As String

    strSQL = "SELECT a.codigo,b.nombre,a.fecha,MIN(a.TimeIn) as TimeIn,MAX(a.TimeOut) as TimeOut FROM ingper a "
    strSQL = strSQL & "INNER JOIN vendedor b ON a.codigo = b.codigo  WHERE "
    strSQL = strSQL & "  a.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    strSQL = strSQL & " and a.fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If codigo <> "%" Then
        strSQL = strSQL & " and a.codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        strSQL = strSQL & " and b.nombre like '" & nombre & "'"

    End If

    strSQL = strSQL & " GROUP BY a.fecha,a.codigo,b.nombre  "

    If Combo1 = "Codigo" Then
        strSQL = strSQL & " ORDER BY a.codigo,a.fecha"

    End If

    If Combo1 = "Fecha" Then
        strSQL = strSQL & " ORDER BY a.fecha,a.codigo"

    End If

    mytablex.Open strSQL, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existen Datos", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    Heading(1) = "Codigo"
    Heading(2) = "Nombre"
    Heading(3) = "Fecha"
    Heading(4) = "HoraInt"
    Heading(5) = "HoraSal"
    Heading(6) = "NroHora"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    With objExcel.ActiveSheet
        
        For I = 1 To 14 Step 1
            .Cells(1, I) = Heading(I)
        Next I
       
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 30
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        
    End With
    
    v = 2
    h = 1
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("codigo")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("nombre")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("timein")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("timeout")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & calcula_hora("" & mytablex.Fields("timeout"), "" & mytablex.Fields("timein"))
        v = v + 1
        mytablex.MoveNext
    Loop
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 
End Sub

