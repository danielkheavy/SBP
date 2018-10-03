VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcajacie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrega de Dinero"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Cuadre de Efectivo"
      Height          =   8655
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   14775
      Begin VB.CommandButton Command3 
         Caption         =   "BorraLinea"
         Height          =   495
         Left            =   5520
         TabIndex        =   23
         Top             =   7800
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "LimpiaPantalla"
         Height          =   495
         Left            =   3840
         TabIndex        =   21
         Top             =   7800
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   7215
         Left            =   3840
         TabIndex        =   20
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   12726
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   7
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   8
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   9
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   10
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   11
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dolares"
         Height          =   375
         Left            =   7680
         TabIndex        =   26
         Top             =   8040
         Width           =   1215
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Soles"
         Height          =   375
         Left            =   7680
         TabIndex        =   25
         Top             =   7680
         Width           =   1215
      End
      Begin VB.Label xtotald 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8880
         TabIndex        =   24
         Top             =   8040
         Width           =   2055
      End
      Begin VB.Label xtotals 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8880
         TabIndex        =   22
         Top             =   7680
         Width           =   2055
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1440
         Picture         =   "tcajacie.frx":0000
         Stretch         =   -1  'True
         Top             =   7800
         Width           =   1200
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   240
         Picture         =   "tcajacie.frx":1FA6
         Stretch         =   -1  'True
         Top             =   7800
         Width           =   1200
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Filtro"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox turno 
      Height          =   495
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox caja 
      Height          =   495
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox fecha 
      Height          =   495
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu flo933 
      Caption         =   "&Excell"
   End
   Begin VB.Menu flo444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcajacie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim buffer(50)      As String

Dim jindx           As Integer

Dim mmesacod(15000) As String

Dim wmesacod(15000) As String

Dim wwmesacod(50)   As String

Dim mmesapag        As Integer

Dim mmesatop        As Integer

Dim msalcod(100)    As String

Dim msalpag         As Integer

Dim msaltop         As Integer

Option Explicit

Private Sub Command1_Click()

    If Not IsDate(fecha) Then
        fecha = ""
        fecha.SetFocus
        Exit Sub

    End If

    If Len(Trim(caja)) = 0 Then
        caja.SetFocus
        Exit Sub

    End If

    If Len(Trim(turno)) = 0 Then
        turno.SetFocus
        Exit Sub

    End If

    Frame1.Visible = True
    carga_fpago

End Sub

Private Sub Command2_Click()

    Dim buf As String

    buf = "delete  from cajaciega where fecha='" & fecha & "' and caja='" & caja & "' and turno='" & turno & "'"
    cn.Execute (buf)
    carga_fpago

End Sub

Private Sub Command3_Click()

    On Error GoTo cmd_9012

    Dim buf As String

    buf = "delete  from cajaciega where fecha='" & fecha & "' and caja='" & caja & "' and turno='" & turno & "' and id=" & Val(dbGrid1.columns(0))
    cn.Execute (buf)
    carga_fpago
    Exit Sub
cmd_9012:
    Exit Sub

End Sub

Private Sub dbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Select Case ColIndex

        Case 7, 8
            'if val()
            'MsgBox OldValue
            'dbgrid1.Columns(9)=val(dbgrid1.Columns(7))-val(dbgrid1.Columns(7))
     
        Case Else
            Cancel = True

    End Select

End Sub

Private Sub flo444_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tcajacie.Hide
    Unload tcajacie

End Sub

Private Sub flo933_Click()

    Dim found      As Integer

    Dim I          As Integer

    Dim v          As Long

    Dim R          As Long

    Dim ih         As Integer

    Dim h          As Integer

    Dim cad        As String

    Dim Tmp        As String

    Dim sw         As Integer

    Dim sdx        As Double

    Dim buf        As String
 
    Dim sdx1       As Double
 
    Dim sdx2       As Double

    Dim sdx3       As Double
 
    Dim mytablex   As New ADODB.Recordset

    Dim mytabley   As New ADODB.Recordset

    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd45612_err

    If Frame1.Visible = False Then
        Exit Sub

    End If

    If MsgBox("Desea Exportar excel", 1, "Aviso") <> 1 Then Exit Sub
    Heading(1) = "FormaPago"
    Heading(2) = "Moneda"
    Heading(3) = "Entrega"
    Heading(4) = "Sistema"
    
    buf = "select * from cajaciega where fecha='" & fecha & "' and caja='" & caja & "' and turno='" & turno & "'"
    mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        Exit Sub

    End If

    sdx = 0
    sdx1 = 0
    sdx3 = 0
    sdx2 = 0
    
    If Inicio_Excel() = 0 Then
        mytabley.Close
        Exit Sub

    End If

    'Llamamos a la funcion que abre el workbook en excel
    
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 20)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 50
        .columns("B").ColumnWidth = 15
        .columns("C").ColumnWidth = 15
        .columns("D").ColumnWidth = 15
    
    End With

    'cabecera
    mytablex.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        objExcel.ActiveSheet.Cells(2, 1) = "'" & mytablex.Fields("nombre")

    End If

    mytablex.Close
    objExcel.ActiveSheet.Cells(2, 5) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 2) = "'Caja Ciega"
    
    '------------------------------------------------
    objExcel.ActiveSheet.Cells(4, 1) = "'FormaPago"
    objExcel.ActiveSheet.Cells(4, 2) = "'Moneda"
    objExcel.ActiveSheet.Cells(4, 3) = "'Entrega"
    objExcel.ActiveSheet.Cells(4, 4) = "'Caja"
    '------------------------------------------------
    v = 5
    h = 1
    
    Do

        If mytabley.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, 1) = "'" & mytabley.Fields("descripcio")
        objExcel.ActiveSheet.Cells(v, 2) = "'" & mytabley.Fields("moneda")
        objExcel.ActiveSheet.Cells(v, 3) = "" & mytabley.Fields("entrega")
        objExcel.ActiveSheet.Cells(v, 4) = suma_fpago("" & mytabley.Fields("fpago"))

        If "" & mytabley.Fields("moneda") = "S" Then
            sdx = sdx + Val("" & mytabley.Fields("entrega"))
            sdx2 = sdx2 + Val("" & mytabley.Fields("encaja"))

        End If

        If "" & mytabley.Fields("moneda") = "D" Then
            sdx1 = sdx1 + Val("" & mytabley.Fields("entrega"))
            sdx3 = sdx3 + Val("" & mytabley.Fields("encaja"))

        End If

        v = v + 1
        mytabley.MoveNext
    Loop
    mytabley.Close
    objExcel.ActiveSheet.Cells(v, 2) = "Soles"
    objExcel.ActiveSheet.Cells(v, 3) = "" & sdx
    objExcel.ActiveSheet.Cells(v, 4) = "" & sdx2
    v = v + 1
    objExcel.ActiveSheet.Cells(v, 2) = "Dolar"
    objExcel.ActiveSheet.Cells(v, 3) = "" & sdx1
    objExcel.ActiveSheet.Cells(v, 4) = "" & sdx3
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd45612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Load()
    fecha = Format(Now, "dd/mm/yyyy")
    menu_carga_mesa "TODOS"
    menu_mesa "INI"

End Sub

Sub menu_carga_mesa(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    For I = 0 To 11
        wwmesacod(I) = ""
    Next I

    For I = 0 To 14999
        mmesacod(I) = ""
        wmesacod(I) = ""
    Next I

    I = -1

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM fpago", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        I = I + 1
        mmesacod(I) = "" & mytablex.Fields("fpago")
        wmesacod(I) = "" & mytablex.Fields("Descripcio")
  
        mytablex.MoveNext
    Loop

    mytablex.Close
    mmesatop = I
    mmesapag = 0

End Sub

Sub menu_mesa(buf As String)

    Dim I As Integer

    Dim j As Integer

    Select Case buf

        Case "INI"
            mmesapag = 0

        Case "SIG"
            mmesapag = mmesapag + 11

            If mmesapag > 102 Then
                mmesapag = 0

            End If

        Case "ANT"
            mmesapag = mmesapag - 11

            If mmesapag < 0 Then
                mmesapag = 0

            End If

    End Select

    j = -1

    For I = mmesapag To 11 + mmesapag
        j = j + 1
        groupmesa(j).Caption = wmesacod(I) 'mmesacod(i)
    Next I

End Sub

Private Sub groupmesa_Click(Index As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim buf      As String

    'MsgBox Trim("" & mmesacod(Index))
    mytabley.Open "select * from fpago where fpago='" & Trim("" & mmesacod(Index)) & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        buf = "select * from cajaciega where fecha='" & fecha & "' and caja='" & caja & "' and turno='" & turno & "'"
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
        mytablex.AddNew
        mytablex.Fields("caja") = Trim(caja)
        mytablex.Fields("turno") = Trim(turno)
        mytablex.Fields("fecha") = Trim(fecha)
        mytablex.Fields("moneda") = Trim("" & mytabley.Fields("moneda"))
        mytablex.Fields("fpago") = Trim("" & mytabley.Fields("fpago"))
        mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))
        mytablex.Fields("encaja") = 0
        mytablex.Update
   
        mytablex.Close

    End If

    mytabley.Close
    carga_fpago

End Sub

Private Sub Image2_Click()

    Dim I As Integer

    For I = 0 To 11
        'groupmesa(i).BackColor = &H80FF80
        'mesa = ""
    Next I

    menu_mesa "SIG"

End Sub

Private Sub Image3_Click()

    Dim I As Integer

    For I = 0 To 11
        'groupmesa(i).BackColor = &H80FF80
        'mesa = ""
    Next I

    menu_mesa "ANT" ', salon

End Sub

Sub carga_fpago()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    Dim sdx1     As Double

    Set mytablex = Nothing
    buf = "select * from cajaciega where fecha='" & fecha & "' and caja='" & caja & "' and turno='" & turno & "'"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex
    dbGrid1.columns(0).Width = 800
    dbGrid1.columns(1).Width = 1000
    dbGrid1.columns(2).Width = 800
    dbGrid1.columns(6).Width = 500
    dbGrid1.columns(7).Width = 1000
    dbGrid1.columns(8).Width = 1000
    '     dbgrid1.columns(9).Width = 1000
    sdx = 0
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("moneda") = "S" Then
            sdx = sdx + Val("" & mytablex.Fields("entrega"))

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            sdx1 = sdx1 + Val("" & mytablex.Fields("entrega"))

        End If

        mytablex.MoveNext
    Loop
    xtotals = Format(sdx, "0.00")
    xtotald = Format(sdx1, "0.00")

End Sub

Function suma_fpago(xbuf As String) As Double

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    buf = "select * from fpagov where fecha='" & fecha & "' and caja='" & caja & "' and turno='" & turno & "' and fpago='" & xbuf & "'"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("recibe"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    suma_fpago = sdx

End Function
