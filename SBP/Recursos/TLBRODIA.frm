VERSION 5.00
Begin VB.Form tlbrodia 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Diario"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ccosto 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4560
      Width           =   6735
   End
   Begin VB.TextBox ruc 
      Height          =   495
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   20
      Text            =   "%"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ComboBox grupo 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5040
      Width           =   6735
   End
   Begin VB.ComboBox libro 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1680
      Width           =   6735
   End
   Begin VB.TextBox documento 
      Height          =   495
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   13
      Text            =   "%"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox tipo 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2160
      Width           =   6735
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   5
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox digitos 
      Height          =   495
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "2"
      Top             =   3600
      Width           =   495
   End
   Begin VB.ComboBox fuente 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2640
      Width           =   6735
   End
   Begin VB.ComboBox cuenta 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3120
      Width           =   6735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CCosto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ruc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Libro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Digitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fuente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label tiporeporte 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -120
      TabIndex        =   0
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Menu dk999 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu fdlo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tlbrodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dk999_Click()
Select Case tiporeporte
       Case "NORMAL"
            libro_diario
       Case "SUNAT"
            sunat_reporte
       Case "MAYOR"
            mayor_reporte
       Case "BALANCEPRUEBA"
            balance_prueba
End Select
End Sub

Private Sub fdlo33_Click()
tlbrodia.Hide
Unload tlbrodia
End Sub

Private Sub Form_Activate()
If tiporeporte = "MAYOR" Then
   Grupo.ListIndex = 4
End If
End Sub

Private Sub Form_Load()
Dim mytablex As New ADODB.Recordset
fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = "30/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fuente.Clear
fuente.AddItem "%"
    mytablex.Open "select * from fuente ", cn, adOpenStatic, adLockOptimistic
    Do
    If mytablex.EOF Then Exit Do
    fuente.AddItem Trim("" & mytablex.Fields("fuente")) & "|" & mytablex.Fields("descripcio")
    mytablex.MoveNext
    Loop
    mytablex.Close
    fuente.ListIndex = 0

cuenta.Clear
cuenta.AddItem "%"
    mytablex.Open "select * from cuentas order by codcta", cn, adOpenStatic, adLockOptimistic
    Do
    If mytablex.EOF Then Exit Do
    cuenta.AddItem Trim("" & mytablex.Fields("codcta")) & "|" & mytablex.Fields("descripcio")
    mytablex.MoveNext
    Loop
    mytablex.Close
    cuenta.ListIndex = 0

libro.Clear
libro.AddItem "%"
    mytablex.Open "select * from libroauxiliar order by libroauxiliar", cn, adOpenStatic, adLockOptimistic
    Do
    If mytablex.EOF Then Exit Do
    libro.AddItem Trim("" & mytablex.Fields("libroauxiliar")) & "|" & mytablex.Fields("descripcio")
    mytablex.MoveNext
    Loop
    mytablex.Close
    libro.ListIndex = 0


ccosto.Clear
ccosto.AddItem "%"
    mytablex.Open "select * from ccosto order by ccosto", cn, adOpenStatic, adLockOptimistic
    Do
    If mytablex.EOF Then Exit Do
    ccosto.AddItem Trim("" & mytablex.Fields("ccosto")) & "|" & mytablex.Fields("descripcio")
    mytablex.MoveNext
    Loop
    mytablex.Close
    ccosto.ListIndex = 0

tipo.Clear
tipo.AddItem "%"
    mytablex.Open "select * from docta order by docta", cn, adOpenStatic, adLockOptimistic
    Do
    If mytablex.EOF Then Exit Do
    tipo.AddItem Trim("" & mytablex.Fields("docta")) & "|" & mytablex.Fields("descripcio")
    mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0

Grupo.Clear
Grupo.AddItem "ASIENTO"
Grupo.AddItem "DOCUMENTO"
Grupo.AddItem "LIBRO"
Grupo.AddItem "FUENTE"
Grupo.AddItem "CUENTA"
Grupo.AddItem "RUC"
Grupo.AddItem "CCOSTO"
Grupo.AddItem "FECHA"
Grupo.ListIndex = 0








End Sub
Sub sunat_reporte()
 Dim found As Integer
 Dim i As Integer
 Dim v As Long
 Dim R As Long
 Dim ih As Integer
 Dim H As Integer
 Dim cad As String
 Dim tmp1 As String
 Dim buf As String
 Dim sdx As Double
 Dim xdebito As Double
 Dim xcredito As Double
 Dim xxdebito As Double
 Dim xxcredito As Double
 Dim TTMP As String
 Dim mytabley As New ADODB.Recordset
 Dim mytablex As New ADODB.Recordset
 Dim sw As Integer
 Dim Tmp As String
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd4445612_err
If Val(digitos) <= 0 Then
   digitos = "2"
End If
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
'MsgBox "abc"


cad = "SELECT left(Cuenta," & digitos & ") as Cuenta,id,cod_asien,fecha_asi,debito,credito,descripcio,motivo,cod_libro,corre_libro,comproba,fuente,nro_ruc,ccosto  from asientos   where  "
'MsgBox cad
cad = cad & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
cad = cad & " and fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "' "
If Trim(documento) <> "%" Then
   cad = cad & " and comproba='" & documento & "'"
End If
If Trim(libro) <> "%" Then
   cad = cad & " and cod_libro='" & Trim(extra_loquesea(libro)) & "'"
End If
If Trim(tipo) <> "%" Then
   cad = cad & " and tipo='" & Trim(extra_loquesea(tipo)) & "'"
End If
If Trim(RUC) <> "%" Then
   cad = cad & " and nro_ruc='" & RUC & "'"
End If
If Trim(ccosto) <> "%" Then
   cad = cad & " and ccosto='" & Trim(extra_loquesea(ccosto)) & "'"
End If

If Trim(fuente) <> "%" Then
   cad = cad & " and fuente='" & Trim(extra_loquesea(fuente)) & "'"
End If
If Trim(cuenta) <> "%" Then
   cad = cad & " and cuenta='" & Mid$(Trim(extra_loquesea(cuenta)), 1, digitos) & "'"
End If
TTMP = ""
            If UCase$(Grupo) = "ASIENTO" Then
             TTMP = "cod_asien"
          End If
                    
          If UCase$(Grupo) = "DOCUMENTO" Then
             TTMP = "comproba"
          End If
          If UCase$(Grupo) = "LIBRO" Then
             TTMP = "cod_libro"
          End If
          If UCase$(Grupo) = "FUENTE" Then
             TTMP = "fuente"
          End If
          If UCase$(Grupo) = "CUENTA" Then
             TTMP = "cuenta"
          End If
          If UCase$(Grupo) = "RUC" Then
             TTMP = "cod_ruc"
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             TTMP = "ccosto"
          End If
          If UCase$(Grupo) = "FECHA" Then
             TTMP = "fecha_asi"
          End If

If TTMP = "cod_asien" Then
cad = cad & " order by " & TTMP & ",id,fecha_asi"
GoTo sigue2
End If
If TTMP = "fecha_asi" Then
cad = cad & " order by " & TTMP & ",cod_asien,id"
GoTo sigue2
End If
cad = cad & " order by " & TTMP & ",cod_asien,id,fecha_asi"


sigue2:
mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    If mytablex.RecordCount = 0 Then
       MsgBox "No existen datos", 48, "Aviso"
       mytablex.Close
       Exit Sub
    End If
    mytablex.MoveFirst
If MsgBox("Desea Generar Reporte", 1, "Aviso") <> 1 Then
   mytablex.Close
   Exit Sub
End If
   If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 10)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 15
        .columns("C").ColumnWidth = 30
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 10
        .columns("f").ColumnWidth = 15
        .columns("g").ColumnWidth = 15
        .columns("h").ColumnWidth = 40
End With
    'cabecera
    mytabley.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic
    If mytabley.RecordCount > 0 Then
    objExcel.ActiveSheet.Cells(2, 1) = "'" & mytabley.Fields("nombre")
    End If
    mytabley.Close
    objExcel.ActiveSheet.Cells(2, 5) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 1) = "'Reporte de Libro Diario:" + "Desde:" + fechai + "   ----Hasta:" + fechaf + " ---en Soles(S/.)"
    '------------------------------------------------
    objExcel.ActiveSheet.Cells(4, 1) = "'Cod.Asiento"
    objExcel.ActiveSheet.Cells(4, 2) = "'Fecha"
    objExcel.ActiveSheet.Cells(4, 3) = "'Comentario (Glosa)"
    objExcel.ActiveSheet.Cells(4, 4) = "'Libro"
    objExcel.ActiveSheet.Cells(4, 5) = "'Correlativo"
    objExcel.ActiveSheet.Cells(4, 6) = "'Nro.Documento"
    objExcel.ActiveSheet.Cells(4, 7) = "'Cod.Cta"
    objExcel.ActiveSheet.Cells(4, 8) = "'Descripcio.Cta"
    objExcel.ActiveSheet.Cells(4, 9) = "'Debitos"
    objExcel.ActiveSheet.Cells(4, 10) = "'Creditos"
    
    
    '------------------------------------------------
v = 5
H = 1
    xdebito = 0
    xcredito = 0
    xxdebito = 0
    xxcredito = 0
    sw = 0
    
    Do
         If mytablex.EOF Then Exit Do
          If UCase$(Grupo) = "ASIENTO" Then
             tmp1 = "" & mytablex.Fields("cod_asien")
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             tmp1 = "" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             tmp1 = "" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             tmp1 = "" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             tmp1 = "" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             tmp1 = "" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             tmp1 = "" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             tmp1 = "" & mytablex.Fields("fecha_asi")
          End If
          
          
         If sw = 0 Then
            sw = 1
            If UCase$(Grupo) = "ASIENTO" Then
             Tmp = "" & mytablex.Fields("cod_asien")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
             
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             Tmp = "" & mytablex.Fields("comproba")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             Tmp = "" & mytablex.Fields("cod_libro")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             Tmp = "" & mytablex.Fields("fuente")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             Tmp = "" & mytablex.Fields("cuenta")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             Tmp = "" & mytablex.Fields("cod_ruc")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             Tmp = "" & mytablex.Fields("ccosto")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             Tmp = "" & mytablex.Fields("fecha_asi")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fecha_asi")
          End If
            
            
            v = v + 1
         End If
         If Tmp <> tmp1 Then
            objExcel.ActiveSheet.Cells(v, 3) = "Subtotal"
            objExcel.ActiveSheet.Cells(v, 9) = xdebito
            objExcel.ActiveSheet.Cells(v, 10) = xcredito
            xdebito = 0
            xcredito = 0
            v = v + 1
          If UCase$(Grupo) = "ASIENTO" Then
             Tmp = "" & mytablex.Fields("cod_asien")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             Tmp = "" & mytablex.Fields("comproba")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             Tmp = "" & mytablex.Fields("cod_libro")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             Tmp = "" & mytablex.Fields("fuente")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             Tmp = "" & mytablex.Fields("cuenta")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             Tmp = "" & mytablex.Fields("cod_ruc")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             Tmp = "" & mytablex.Fields("ccosto")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             Tmp = "" & mytablex.Fields("fecha_asi")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fecha_asi")
          End If
          v = v + 1
         End If
         
      
         objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
         objExcel.ActiveSheet.Cells(v, 2) = "'" & mytablex.Fields("fecha_asi")
         objExcel.ActiveSheet.Cells(v, 3) = "'" & mytablex.Fields("motivo")
         objExcel.ActiveSheet.Cells(v, 4) = "'" & mytablex.Fields("cod_libro")
         objExcel.ActiveSheet.Cells(v, 5) = "'" & mytablex.Fields("corre_libro")
         objExcel.ActiveSheet.Cells(v, 6) = "'" & mytablex.Fields("comproba")
         
         objExcel.ActiveSheet.Cells(v, 7) = "'" & mytablex.Fields("cuenta")
         objExcel.ActiveSheet.Cells(v, 8) = "'" & mytablex.Fields("descripcio")
         objExcel.ActiveSheet.Cells(v, 9) = Val("" & mytablex.Fields("debito"))
         objExcel.ActiveSheet.Cells(v, 10) = Val("" & mytablex.Fields("credito"))
         
         xdebito = xdebito + Val("" & mytablex.Fields("debito"))
         xcredito = xcredito + Val("" & mytablex.Fields("credito"))
         
         xxdebito = xxdebito + Val("" & mytablex.Fields("debito"))
         xxcredito = xxcredito + Val("" & mytablex.Fields("credito"))
         v = v + 1
mytablex.MoveNext
Loop
'mytablex.Close
objExcel.ActiveSheet.Cells(v, 3) = "Subtotal"
objExcel.ActiveSheet.Cells(v, 9) = xdebito
objExcel.ActiveSheet.Cells(v, 10) = xcredito
v = v + 1
objExcel.ActiveSheet.Cells(v, 3) = "Total Libro Diario"
objExcel.ActiveSheet.Cells(v, 9) = xxdebito
objExcel.ActiveSheet.Cells(v, 10) = xxcredito
Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
Exit Sub
cmd4445612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub


End Sub
Sub libro_diario()
 Dim found As Integer
 Dim i As Integer
 Dim v As Long
 Dim R As Long
 Dim ih As Integer
 Dim H As Integer
 Dim cad As String
 Dim tmp1 As String
 Dim buf As String
 Dim sdx As Double
 Dim xdebito As Double
 Dim xcredito As Double
 Dim xxdebito As Double
 Dim xxcredito As Double
 Dim TTMP As String
 Dim mytabley As New ADODB.Recordset
 Dim mytablex As New ADODB.Recordset
 Dim sw As Integer
 Dim Tmp As String
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd445612_err
If Val(digitos) <= 0 Then
   digitos = "2"
End If
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
'MsgBox "abc"


cad = "SELECT left(Cuenta," & digitos & ") as Cuenta,ID,cod_asien,fecha_asi,debito,credito,descripcio,motivo,cod_libro,comproba,fuente,nro_ruc,ccosto  from asientos   where  "
'MsgBox cad
cad = cad & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
cad = cad & " and fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "' "
If Trim(documento) <> "%" Then
   cad = cad & " and comproba='" & documento & "'"
End If
If Trim(libro) <> "%" Then
   cad = cad & " and cod_libro='" & Trim(extra_loquesea(libro)) & "'"
End If
If Trim(tipo) <> "%" Then
   cad = cad & " and tipo='" & Trim(extra_loquesea(tipo)) & "'"
End If
If Trim(RUC) <> "%" Then
   cad = cad & " and nro_ruc='" & RUC & "'"
End If
If Trim(ccosto) <> "%" Then
   cad = cad & " and ccosto='" & Trim(extra_loquesea(ccosto)) & "'"
End If

If Trim(fuente) <> "%" Then
   cad = cad & " and fuente='" & Trim(extra_loquesea(fuente)) & "'"
End If
If Trim(cuenta) <> "%" Then
   cad = cad & " and cuenta='" & Mid$(Trim(extra_loquesea(cuenta)), 1, digitos) & "'"
End If
TTMP = ""
            If UCase$(Grupo) = "ASIENTO" Then
             TTMP = "cod_asien"
          End If
                    
          If UCase$(Grupo) = "DOCUMENTO" Then
             TTMP = "comproba"
          End If
          If UCase$(Grupo) = "LIBRO" Then
             TTMP = "cod_libro"
          End If
          If UCase$(Grupo) = "FUENTE" Then
             TTMP = "fuente"
          End If
          If UCase$(Grupo) = "CUENTA" Then
             TTMP = "cuenta"
          End If
          If UCase$(Grupo) = "RUC" Then
             TTMP = "cod_ruc"
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             TTMP = "ccosto"
          End If
          If UCase$(Grupo) = "FECHA" Then
             TTMP = "fecha_asi"
          End If

If TTMP = "cod_asien" Then
cad = cad & " order by " & TTMP & ",ID,fecha_asi"
GoTo sigue
End If
If TTMP = "fecha_asi" Then
cad = cad & " order by " & TTMP & ",cod_asien,ID"
GoTo sigue
End If
cad = cad & " order by " & TTMP & ",cod_asien,ID,fecha_asi"


sigue:
mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    If mytablex.RecordCount = 0 Then
       MsgBox "No existen datos", 48, "Aviso"
       mytablex.Close
       Exit Sub
    End If
    mytablex.MoveFirst
If MsgBox("Desea Generar Reporte", 1, "Aviso") <> 1 Then
   mytablex.Close
   Exit Sub
End If
   If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 7)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 15
        .columns("C").ColumnWidth = 40
        .columns("D").ColumnWidth = 15
        .columns("E").ColumnWidth = 15
        .columns("f").ColumnWidth = 15
        .columns("g").ColumnWidth = 40
    
End With
    'cabecera
    mytabley.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic
    If mytabley.RecordCount > 0 Then
    objExcel.ActiveSheet.Cells(2, 1) = "'" & mytabley.Fields("nombre")
    End If
    mytabley.Close
    objExcel.ActiveSheet.Cells(2, 5) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 1) = "'Reporte de Libro Diario:" + "Desde:" + fechai + "   ----Hasta:" + fechaf + " ---en Soles(S/.)"
    '------------------------------------------------
    objExcel.ActiveSheet.Cells(4, 1) = "'Cod.Asiento"
    objExcel.ActiveSheet.Cells(4, 2) = "'Cod.Cta"
    objExcel.ActiveSheet.Cells(4, 3) = "'Descripcio.Cta"
    objExcel.ActiveSheet.Cells(4, 4) = "'Fecha"
    objExcel.ActiveSheet.Cells(4, 5) = "'Debitos"
    objExcel.ActiveSheet.Cells(4, 6) = "'Creditos"
    objExcel.ActiveSheet.Cells(4, 7) = "'Comentario (Glosa Movimiento)"
    
    '------------------------------------------------
v = 5
H = 1
    xdebito = 0
    xcredito = 0
    xxdebito = 0
    xxcredito = 0
    sw = 0
    
    Do
         If mytablex.EOF Then Exit Do
          If UCase$(Grupo) = "ASIENTO" Then
             tmp1 = "" & mytablex.Fields("cod_asien")
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             tmp1 = "" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             tmp1 = "" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             tmp1 = "" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             tmp1 = "" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             tmp1 = "" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             tmp1 = "" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             tmp1 = "" & mytablex.Fields("fecha_asi")
          End If
          
          
         If sw = 0 Then
            sw = 1
            If UCase$(Grupo) = "ASIENTO" Then
             Tmp = "" & mytablex.Fields("cod_asien")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
             
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             Tmp = "" & mytablex.Fields("comproba")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             Tmp = "" & mytablex.Fields("cod_libro")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             Tmp = "" & mytablex.Fields("fuente")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             Tmp = "" & mytablex.Fields("cuenta")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             Tmp = "" & mytablex.Fields("cod_ruc")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             Tmp = "" & mytablex.Fields("ccosto")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             Tmp = "" & mytablex.Fields("fecha_asi")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fecha_asi")
          End If
            
            
            v = v + 1
         End If
         If Tmp <> tmp1 Then
            objExcel.ActiveSheet.Cells(v, 3) = "Subtotal"
            objExcel.ActiveSheet.Cells(v, 5) = xdebito
            objExcel.ActiveSheet.Cells(v, 6) = xcredito
            xdebito = 0
            xcredito = 0
            v = v + 1
          If UCase$(Grupo) = "ASIENTO" Then
             Tmp = "" & mytablex.Fields("cod_asien")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             Tmp = "" & mytablex.Fields("comproba")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             Tmp = "" & mytablex.Fields("cod_libro")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             Tmp = "" & mytablex.Fields("fuente")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             Tmp = "" & mytablex.Fields("cuenta")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             Tmp = "" & mytablex.Fields("cod_ruc")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             Tmp = "" & mytablex.Fields("ccosto")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             Tmp = "" & mytablex.Fields("fecha_asi")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fecha_asi")
          End If
          v = v + 1
         End If
         objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
         objExcel.ActiveSheet.Cells(v, 2) = "'" & mytablex.Fields("cuenta")
         objExcel.ActiveSheet.Cells(v, 3) = "'" & mytablex.Fields("descripcio")
         objExcel.ActiveSheet.Cells(v, 4) = "'" & mytablex.Fields("fecha_asi")
         objExcel.ActiveSheet.Cells(v, 5) = Val("" & mytablex.Fields("debito"))
         objExcel.ActiveSheet.Cells(v, 6) = Val("" & mytablex.Fields("credito"))
         objExcel.ActiveSheet.Cells(v, 7) = "'" & mytablex.Fields("motivo")
         xdebito = xdebito + Val("" & mytablex.Fields("debito"))
         xcredito = xcredito + Val("" & mytablex.Fields("credito"))
         
         xxdebito = xxdebito + Val("" & mytablex.Fields("debito"))
         xxcredito = xxcredito + Val("" & mytablex.Fields("credito"))
         v = v + 1
mytablex.MoveNext
Loop
'mytablex.Close
objExcel.ActiveSheet.Cells(v, 3) = "Subtotal"
objExcel.ActiveSheet.Cells(v, 5) = xdebito
objExcel.ActiveSheet.Cells(v, 6) = xcredito
v = v + 1
objExcel.ActiveSheet.Cells(v, 3) = "Total Libro Diario"
objExcel.ActiveSheet.Cells(v, 5) = xxdebito
objExcel.ActiveSheet.Cells(v, 6) = xxcredito
Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
Exit Sub
cmd445612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub

End Sub
Sub mayor_reporte()
Dim found As Integer
 Dim i As Integer
 Dim v As Long
 Dim R As Long
 Dim ih As Integer
 Dim H As Integer
 Dim cad As String
 Dim tmp1 As String
 Dim buf As String
 Dim sdx As Double
 Dim xdebito As Double
 Dim xcredito As Double
 Dim xxdebito As Double
 Dim xxcredito As Double
 Dim TTMP As String
 Dim mytabley As New ADODB.Recordset
 Dim mytablex As New ADODB.Recordset
 Dim sw As Integer
 Dim Tmp As String
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd9445612_err
If Val(digitos) <= 0 Then
   digitos = "2"
End If
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
'MsgBox "abc"


cad = "SELECT left(Cuenta," & digitos & ") as Cuenta,ID,cod_asien,fecha_asi,debito,credito,descripcio,motivo,cod_libro,comproba,fuente,nro_ruc,ccosto  from asientos   where  "
'MsgBox cad
cad = cad & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
cad = cad & " and fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "' "
If Trim(documento) <> "%" Then
   cad = cad & " and comproba='" & documento & "'"
End If
If Trim(libro) <> "%" Then
   cad = cad & " and cod_libro='" & Trim(extra_loquesea(libro)) & "'"
End If
If Trim(tipo) <> "%" Then
   cad = cad & " and tipo='" & Trim(extra_loquesea(tipo)) & "'"
End If
If Trim(RUC) <> "%" Then
   cad = cad & " and nro_ruc='" & RUC & "'"
End If
If Trim(ccosto) <> "%" Then
   cad = cad & " and ccosto='" & Trim(extra_loquesea(ccosto)) & "'"
End If

If Trim(fuente) <> "%" Then
   cad = cad & " and fuente='" & Trim(extra_loquesea(fuente)) & "'"
End If
If Trim(cuenta) <> "%" Then
   cad = cad & " and cuenta='" & Mid$(Trim(extra_loquesea(cuenta)), 1, digitos) & "'"
End If
TTMP = ""
            If UCase$(Grupo) = "ASIENTO" Then
             TTMP = "cod_asien"
          End If
                    
          If UCase$(Grupo) = "DOCUMENTO" Then
             TTMP = "comproba"
          End If
          If UCase$(Grupo) = "LIBRO" Then
             TTMP = "cod_libro"
          End If
          If UCase$(Grupo) = "FUENTE" Then
             TTMP = "fuente"
          End If
          If UCase$(Grupo) = "CUENTA" Then
             TTMP = "cuenta"
          End If
          If UCase$(Grupo) = "RUC" Then
             TTMP = "cod_ruc"
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             TTMP = "ccosto"
          End If
          If UCase$(Grupo) = "FECHA" Then
             TTMP = "fecha_asi"
          End If

If TTMP = "cod_asien" Then
cad = cad & " order by " & TTMP & ",id,fecha_asi"
GoTo sigue9
End If
If TTMP = "fecha_asi" Then
cad = cad & " order by " & TTMP & ",cod_asien,id"
GoTo sigue9
End If
cad = cad & " order by " & TTMP & ",cod_asien,id,fecha_asi"


sigue9:
mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    If mytablex.RecordCount = 0 Then
       MsgBox "No existen datos", 48, "Aviso"
       mytablex.Close
       Exit Sub
    End If
    mytablex.MoveFirst
If MsgBox("Desea Generar Reporte", 1, "Aviso") <> 1 Then
   mytablex.Close
   Exit Sub
End If
   If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 7)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 15
        .columns("C").ColumnWidth = 40
        .columns("D").ColumnWidth = 15
        .columns("E").ColumnWidth = 15
        .columns("f").ColumnWidth = 15
        .columns("g").ColumnWidth = 40
    
End With
    'cabecera
    mytabley.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic
    If mytabley.RecordCount > 0 Then
    objExcel.ActiveSheet.Cells(2, 1) = "'" & mytabley.Fields("nombre")
    End If
    mytabley.Close
    objExcel.ActiveSheet.Cells(2, 5) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 1) = "'Reporte de Libro Diario:" + "Desde:" + fechai + "   ----Hasta:" + fechaf + " ---en Soles(S/.)"
    '------------------------------------------------
    objExcel.ActiveSheet.Cells(4, 1) = "'Fecha"
    objExcel.ActiveSheet.Cells(4, 2) = "'Cod.Asiento"
    objExcel.ActiveSheet.Cells(4, 3) = "'Comentario (Glosa)"
    objExcel.ActiveSheet.Cells(4, 4) = "'Deudor"
    objExcel.ActiveSheet.Cells(4, 5) = "'Acreedor"
    
    '------------------------------------------------
v = 5
H = 1
    xdebito = 0
    xcredito = 0
    xxdebito = 0
    xxcredito = 0
    sw = 0
    
    Do
         If mytablex.EOF Then Exit Do
          If UCase$(Grupo) = "ASIENTO" Then
             tmp1 = "" & mytablex.Fields("cod_asien")
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             tmp1 = "" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             tmp1 = "" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             tmp1 = "" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             tmp1 = "" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             tmp1 = "" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             tmp1 = "" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             tmp1 = "" & mytablex.Fields("fecha_asi")
          End If
          
          
         If sw = 0 Then
            sw = 1
            If UCase$(Grupo) = "ASIENTO" Then
             Tmp = "" & mytablex.Fields("cod_asien")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
             
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             Tmp = "" & mytablex.Fields("comproba")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             Tmp = "" & mytablex.Fields("cod_libro")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             Tmp = "" & mytablex.Fields("fuente")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             Tmp = "" & mytablex.Fields("cuenta")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             Tmp = "" & mytablex.Fields("cod_ruc")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             Tmp = "" & mytablex.Fields("ccosto")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             Tmp = "" & mytablex.Fields("fecha_asi")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fecha_asi")
          End If
            
            
            v = v + 1
         End If
         If Tmp <> tmp1 Then
            objExcel.ActiveSheet.Cells(v, 3) = "Subtotal"
            objExcel.ActiveSheet.Cells(v, 4) = xdebito
            objExcel.ActiveSheet.Cells(v, 5) = xcredito
            xdebito = 0
            xcredito = 0
            v = v + 1
          If UCase$(Grupo) = "ASIENTO" Then
             Tmp = "" & mytablex.Fields("cod_asien")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_asien")
          End If
          If UCase$(Grupo) = "DOCUMENTO" Then
             Tmp = "" & mytablex.Fields("comproba")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("comproba")
          End If
          If UCase$(Grupo) = "LIBRO" Then
             Tmp = "" & mytablex.Fields("cod_libro")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_libro")
          End If
          If UCase$(Grupo) = "FUENTE" Then
             Tmp = "" & mytablex.Fields("fuente")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fuente")
          End If
          If UCase$(Grupo) = "CUENTA" Then
             Tmp = "" & mytablex.Fields("cuenta")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
          End If
          If UCase$(Grupo) = "RUC" Then
             Tmp = "" & mytablex.Fields("cod_ruc")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cod_ruc")
          End If
          If UCase$(Grupo) = "CCOSTO" Then
             Tmp = "" & mytablex.Fields("ccosto")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("ccosto")
          End If
          If UCase$(Grupo) = "FECHA" Then
             Tmp = "" & mytablex.Fields("fecha_asi")
             objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fecha_asi")
          End If
          v = v + 1
         End If
         
         
         
         objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("fecha_asi")
         objExcel.ActiveSheet.Cells(v, 2) = "'" & mytablex.Fields("cod_asien")
         objExcel.ActiveSheet.Cells(v, 3) = "'" & mytablex.Fields("motivo")
         objExcel.ActiveSheet.Cells(v, 4) = Val("" & mytablex.Fields("debito"))
         objExcel.ActiveSheet.Cells(v, 5) = Val("" & mytablex.Fields("credito"))
         
         xdebito = xdebito + Val("" & mytablex.Fields("debito"))
         xcredito = xcredito + Val("" & mytablex.Fields("credito"))
         
         xxdebito = xxdebito + Val("" & mytablex.Fields("debito"))
         xxcredito = xxcredito + Val("" & mytablex.Fields("credito"))
         v = v + 1
mytablex.MoveNext
Loop
'mytablex.Close
objExcel.ActiveSheet.Cells(v, 3) = "Subtotal"
objExcel.ActiveSheet.Cells(v, 4) = xdebito
objExcel.ActiveSheet.Cells(v, 5) = xcredito
v = v + 1
objExcel.ActiveSheet.Cells(v, 3) = "Total "
objExcel.ActiveSheet.Cells(v, 4) = xxdebito
objExcel.ActiveSheet.Cells(v, 5) = xxcredito
Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
Exit Sub
cmd9445612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub

End Sub

Sub balance_prueba()
Dim found As Integer
 Dim i As Integer
 Dim v As Long
 Dim R As Long
 Dim ih As Integer
 Dim H As Integer
 Dim cad As String
 Dim tmp1 As String
 Dim buf As String
 Dim sdx As Double
 Dim xdebito As Double
 Dim xcredito As Double
 Dim xxdebito As Double
 Dim xxcredito As Double
 Dim mytabley As New ADODB.Recordset
 Dim mytablex As New ADODB.Recordset
 Dim sw As Integer
 Dim Tmp As String
    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd0445612_err
If Val(digitos) <= 0 Then
   digitos = "2"
End If
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub


cad = "SELECT left(Cuenta," & digitos & ") as Cuenta,Fecha_asi,Fuente,Comproba,tipo_cta,Cantidad,cod_asien,referencia  from asientos   where  "
cad = cad & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
cad = cad & " and fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "' "
If Trim(fuente) <> "%" Then
   cad = cad & " and fuente='" & Trim(extra_loquesea(fuente)) & "'"
End If
If Trim(cuenta) <> "%" Then
   cad = cad & " and cuenta='" & Trim(extra_loquesea(cuenta)) & "'"
End If

cad = cad & " order by cuenta,fecha_asi,cod_asien"
mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
'Set dbGrid1.DataSource = mytablex
    If mytablex.RecordCount = 0 Then
       MsgBox "No existen datos", 48, "Aviso"
       mytablex.Close
       Exit Sub
    End If
    mytablex.MoveFirst
If MsgBox("Desea Generar Reporte", 1, "Aviso") <> 1 Then
   mytablex.Close
   Exit Sub
End If
   If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 5)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 30
        .columns("C").ColumnWidth = 15
        .columns("D").ColumnWidth = 15
        .columns("E").ColumnWidth = 15
        .columns("f").ColumnWidth = 15
    
End With
    'cabecera
    mytabley.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic
    If mytabley.RecordCount > 0 Then
    objExcel.ActiveSheet.Cells(2, 1) = "'" & mytabley.Fields("nombre")
    End If
    mytabley.Close
    objExcel.ActiveSheet.Cells(2, 4) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 1) = "' Balance Prueba:" + "Desde:" + fechai + "   ----Hasta:" + fechaf + " ---en Soles(S/.)"
    '------------------------------------------------
    objExcel.ActiveSheet.Cells(4, 1) = "'Cuenta"
    objExcel.ActiveSheet.Cells(4, 2) = "'NOmbreCuenta"
    objExcel.ActiveSheet.Cells(4, 3) = "'Debito"
    objExcel.ActiveSheet.Cells(4, 4) = "'Credito"
    
    '------------------------------------------------
v = 5
H = 1
    xdebito = 0
    xcredito = 0
    xxdebito = 0
    xxcredito = 0
    sw = 0
    
    Do
         If mytablex.EOF Then Exit Do
          tmp1 = "" & mytablex.Fields("cuenta")
         If sw = 0 Then
            sw = 1
            Tmp = "" & mytablex.Fields("cuenta")
            objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
            buf = ""
            mytabley.Open "select * from cuentas where codcta='" & "" & mytablex.Fields("cuenta") & "'", cn, adOpenStatic, adLockOptimistic
            If mytabley.RecordCount > 0 Then
               buf = "" & mytabley.Fields("descripcio")
            End If
            mytabley.Close
            objExcel.ActiveSheet.Cells(v, 2) = "'" & buf
            'v = v + 1
         End If
         If Tmp <> tmp1 Then
            objExcel.ActiveSheet.Cells(v, 3) = xdebito
            objExcel.ActiveSheet.Cells(v, 4) = xcredito
            v = v + 1
            xdebito = 0
            xcredito = 0
            Tmp = "" & mytablex.Fields("cuenta")
            objExcel.ActiveSheet.Cells(v, 1) = "'" & mytablex.Fields("cuenta")
            buf = ""
            mytabley.Open "select * from cuentas where codcta='" & "" & mytablex.Fields("cuenta") & "'", cn, adOpenStatic, adLockOptimistic
            If mytabley.RecordCount > 0 Then
               buf = "" & mytabley.Fields("descripcio")
            End If
            mytabley.Close
            objExcel.ActiveSheet.Cells(v, 2) = "'" & buf
         End If
         
   If "" & mytablex.Fields("tipo_cta") = "D" Then
      'objExcel.ActiveSheet.Cells(v, 5) = "" & mytablex.Fields("cantidad")
      xdebito = xdebito + Val("" & mytablex.Fields("cantidad"))
      xxdebito = xxdebito + Val("" & mytablex.Fields("cantidad"))
   End If
   If "" & mytablex.Fields("tipo_cta") = "H" Then
      'objExcel.ActiveSheet.Cells(v, 6) = "" & mytablex.Fields("cantidad")
      xcredito = xcredito + Val("" & mytablex.Fields("cantidad"))
      xxcredito = xxcredito + Val("" & mytablex.Fields("cantidad"))
   End If
   'v = v + 1
mytablex.MoveNext
Loop
'mytablex.Close
objExcel.ActiveSheet.Cells(v, 3) = xdebito
objExcel.ActiveSheet.Cells(v, 4) = xcredito
v = v + 1
'objExcel.ActiveSheet.Cells(v, 4) = "Total Libro Diario"
objExcel.ActiveSheet.Cells(v, 3) = xxdebito
objExcel.ActiveSheet.Cells(v, 4) = xxcredito
Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
Exit Sub
cmd0445612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub

End Sub


