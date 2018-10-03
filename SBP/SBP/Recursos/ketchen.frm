VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form kitchen 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cocina Monitor"
   ClientHeight    =   8955
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox seccion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10920
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   7920
      Width           =   4215
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   5895
      Left            =   4320
      TabIndex        =   15
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
      BackColor       =   &H00FF8080&
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
      Height          =   1095
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   2175
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SECCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   16
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   4320
      Top             =   7800
      Width           =   10935
   End
   Begin VB.Shape Shape2 
      Height          =   8655
      Left            =   120
      Top             =   120
      Width           =   4095
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
      TabIndex        =   14
      Top             =   7920
      Width           =   4335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   3975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comentarios del Mesero"
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
      TabIndex        =   12
      Top             =   4320
      Width           =   3975
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
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
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label mocupado 
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
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PEDIDOS TERMINADOS"
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
      TabIndex        =   7
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PEDIDOS EN PROCESO"
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
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PEDIDOS EN ESPERA"
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
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MESAS OCUPADAS"
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
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Menu fl9343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "kitchen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tkitchen As New ADODB.Recordset

Dim swestado As String

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
    escribe_comentario

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    'escribe_comentario
End Sub

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    escribe_comentario

End Sub

Private Sub fl9343_Click()
    kitchen.Hide
    Unload kitchen

End Sub

Sub sql_cabeza(buf1 As String)

    Dim buf As String
   
    buf = "Select Servicio as S,Salon,Mesa,Comanda,Producto,Descripcio,Cantidad,Vendedor as Mesero,Hora,Fecha,Estado as E,Zona,Observa1,Observa2,Observa3,observa4 from dcomanda where estado='" & swestado & "'"

    If seccion <> "%" Then
        buf = buf & " and zona='" & Trim(extra_loquesea(seccion)) & "'"

    End If

    buf = buf & " order by hora " & buf1

    If tkitchen.State = 1 Then tkitchen.Close
    Set tkitchen = Nothing
    tkitchen.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = tkitchen
    dbGrid1.columns(0).Width = 250
    dbGrid1.columns(1).Width = 500
    dbGrid1.columns(2).Width = 500
    dbGrid1.columns(3).Width = 800
    dbGrid1.columns(4).Width = 800
    dbGrid1.columns(5).Width = 4000
    dbGrid1.columns(6).Width = 700
    dbGrid1.columns(7).Width = 900
    dbGrid1.columns(8).Width = 800
    dbGrid1.columns(9).Width = 1000
    dbGrid1.columns(10).Width = 200
   
End Sub

Private Sub Form_Activate()
    sql_cabeza "ASC"

End Sub

Private Sub Form_Load()
    carga_seccion
    swestado = "0"

End Sub

Sub carga_seccion()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from seccion", cn, adOpenStatic, adLockOptimistic
    seccion.Clear
    seccion.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        seccion.AddItem "" & mytablex.Fields("Seccion") & "|" & mytablex.Fields("Descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    seccion.ListIndex = 0

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
