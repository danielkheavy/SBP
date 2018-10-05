VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tmoentre 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despacho Monitor"
   ClientHeight    =   8955
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Validacion de Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   390
      TabIndex        =   14
      Top             =   135
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox xbarra 
         Height          =   615
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   21
         Top             =   5520
         Width           =   3255
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   4215
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   10200
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acepta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         TabIndex        =   22
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Barras"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label xnumero 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label xserie 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label xtipo 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label xlocal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   375
         Width           =   1095
      End
   End
   Begin VB.TextBox FECHAI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1080
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   5895
      Left            =   4320
      TabIndex        =   10
      Top             =   1800
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
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
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PROCESAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   12975
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   495
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TERMINADOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EN PROCESO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EN ESPERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA INICIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label pespera 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   4320
      Top             =   7800
      Width           =   10935
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   4320
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   9
      Top             =   7920
      Width           =   4335
   End
   Begin VB.Label ptermino 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label pproceso 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENTREGADOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EN PROCESO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EN ESPERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Menu fl9343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tmoentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tkitchen  As New ADODB.Recordset

Dim tkitchen1 As New ADODB.Recordset

Dim swestado  As String

Private Sub Command1_Click()
    swestado = "0"
    sql_cabeza "ASC"
    Command4.Caption = "PROCESAR"
    Command1.BackColor = &H80FF80
    Command2.BackColor = &H8000000F

End Sub

Private Sub Command2_Click()
    swestado = "1"
    Command4.Caption = "TERMINAR"
    sql_cabeza "ASC"
    Command1.BackColor = &H8000000F
    Command2.BackColor = &HFF8080

End Sub

Private Sub Command3_Click()
    Command4.Caption = ""
    swestado = "2"
    sql_cabeza "DESC"

End Sub

Private Sub Command4_Click()

    Select Case Command4.Caption

        Case "PROCESAR"
            tkitchen.Fields("e") = "1"
            tkitchen.Update

        Case "TERMINAR"
            tkitchen.Fields("e") = "2"
            tkitchen.Update

    End Select

    sql_cabeza "ASC"

End Sub

Private Sub dbgrid1_Click()

    'escribe_comentario
End Sub

Private Sub dbgrid1_DblClick()

    Dim buf As String

    On Error GoTo cmd90888_err

    xlocal = Trim("" & tkitchen.Fields("local"))
    xtipo = Trim("" & tkitchen.Fields("tipo"))
    xserie = Trim("" & tkitchen.Fields("serie"))
    xnumero = Trim("" & tkitchen.Fields("numero"))
    valida_barras
    Exit Sub
cmd90888_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    'escribe_comentario
End Sub

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    escribe_comentario

End Sub

Private Sub fl9343_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tmoentre.Hide
    Unload tmoentre

End Sub

Sub sql_cabeza(buf1 As String)

    Dim buf As String
   
    buf = "Select Local,Tipo,Serie,Numero,Codigo,Nombre,Total,Estado as E,yausado as e from factura where yausado='" & swestado & "' order by local,tipo,serie,numero"

    If tkitchen.State = 1 Then tkitchen.Close
    Set tkitchen = Nothing
    tkitchen.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = tkitchen
    dbGrid1.columns(0).Width = 500 'local
    dbGrid1.columns(1).Width = 500 'tipo
    dbGrid1.columns(2).Width = 500  'serie
    dbGrid1.columns(3).Width = 1000  'numero
    dbGrid1.columns(4).Width = 1000  'codigo
    dbGrid1.columns(5).Width = 4500 'nombre
    dbGrid1.columns(6).Width = 1500  'total
    dbGrid1.columns(7).Width = 500  'estado
   
End Sub

Private Sub Form_Activate()

    sql_cabeza "ASC"

End Sub

Private Sub Form_Load()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    swestado = "0"

End Sub

Private Sub Label10_Click()

End Sub

Private Sub seccion_Click()

    If seccion <> "%" Then
        sql_cabeza "ASC"

    End If

End Sub

Sub escribe_comentario()

    On Error GoTo cmd9090_error

    Label10 = "" & tkitchen.Fields("Observa1")
    Label10 = Label10 & "" & tkitchen.Fields("Observa2")
    Label10 = Label10 & "" & tkitchen.Fields("Observa3")
    Label10 = Label10 & "" & tkitchen.Fields("Observa4")
    Exit Sub
cmd9090_error:
    Exit Sub

End Sub

Sub valida_barras()

    Dim buf As String

    buf = "Select Dua as Ex,Producto,Descripcio,Unidad,Factor,Cantidad,Precio,Total from detalle where local='" & Trim("" & tkitchen.Fields("local")) & "' and tipo='" & Trim("" & tkitchen.Fields("tipo")) & "' and  serie='" & Trim("" & tkitchen.Fields("serie")) & "'  and numero='" & Trim("" & tkitchen.Fields("numero")) & "'"

    If tkitchen1.State = 1 Then tkitchen1.Close
    Set tkitchen1 = Nothing
    tkitchen1.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = tkitchen1
    DBGrid2.columns(0).Width = 500 'local
    DBGrid2.columns(1).Width = 1000 'local
    DBGrid2.columns(2).Width = 4500 'tipo
    DBGrid2.columns(3).Width = 800  'serie
    DBGrid2.columns(4).Width = 800  'numero
    DBGrid2.columns(5).Width = 1000  'codigo
    DBGrid2.columns(6).Width = 1000 'nombre
    Frame1.Visible = True

End Sub

Function valida_producto()

    Dim mytablex As New ADODB.Recordset

End Function

Private Sub Label6_Click()

    Dim buf      As String

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    buf1 = Trim(xbarra)

    If Len(Trim(buf1)) = 0 Then Exit Sub
    mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        found = busca_equiva(buf1) 'busca en la table codigo barras

        If found = 0 Then
            Exit Sub

        End If

        mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If

    End If

    mytablex.Close
    cn.Execute ("update detalle set dua='1' where local='" & Trim("" & tkitchen.Fields("local")) & "' and tipo='" & Trim("" & tkitchen.Fields("tipo")) & "' and  serie='" & Trim("" & tkitchen.Fields("serie")) & "'  and numero='" & Trim("" & tkitchen.Fields("numero")) & "' and producto='" & buf1 & "'")
    MsgBox "Procesado ", 48, "Aviso"
    xbarra = ""
    dbgrid1_DblClick
 
End Sub

Private Sub Label7_Click()
    fl9343_Click

End Sub

Function busca_equiva(buf As String) As Integer

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Integer

    Dim I        As Integer

    buf1 = ""

    If flag_denisse = "1" Then
        sdx = 18 - Len(buf)

        For I = 1 To sdx
            buf1 = buf1 & "0"
        Next I

    End If

    buf1 = buf1 & buf

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM productb where barras='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
   
    mytablex.Open "SELECT * FROM producto where barras='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1

    End If

    mytablex.Close

End Function

