VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tdiagnos 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Diagnosticos "
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   7455
      Left            =   480
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid2 
         Height          =   6015
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   10610
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
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
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tdiagnos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox diagnostico 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   22
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox observa4 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   21
         Top             =   3240
         Width           =   6255
      End
      Begin VB.TextBox observa3 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   20
         Top             =   2760
         Width           =   6255
      End
      Begin VB.TextBox observa2 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   19
         Top             =   2280
         Width           =   6255
      End
      Begin VB.TextBox observa1 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   17
         Top             =   1800
         Width           =   6255
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tdiagnos.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tdiagnos.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox enfermedad 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label empresa 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3720
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label cliente 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dconsulta 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dsede 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comentario"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label nenfermedad 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   3240
         TabIndex        =   11
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enfermedad"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   11910
      TabIndex        =   3
      Top             =   0
      Width           =   11970
      Begin VB.TextBox sede 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox consulta 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   4440
         MaxLength       =   11
         TabIndex        =   0
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdiagnos.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdiagnos.frx":291C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Borrar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdiagnos.frx":3B2E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdiagnos.frx":4D40
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAddEntry 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Picture         =   "tdiagnos.frx":5F52
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label xempresa 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6720
         TabIndex        =   38
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label xnombree 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8040
         TabIndex        =   37
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label xnombre 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8040
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xcliente 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6720
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empresa"
         Height          =   375
         Left            =   6000
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
         Height          =   375
         Left            =   6000
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sede"
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Consulta"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
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
   Begin VB.Menu ahyy1 
      Caption         =   "&Add"
   End
   Begin VB.Menu dmi22 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dfj8221 
      Caption         =   "&Borra"
   End
   Begin VB.Menu dk281 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu fdo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tdiagnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsdiag As New ADODB.Recordset

Private Sub SQL()

    On Error GoTo cmd5_err

    Dim cad As String

    cad = "SELECT Diagnostico.Sede,Diagnostico.Consulta,Diagnostico.Diagnostico ,Diagnostico.Enfermedad,Enfermedad.Nombre,Diagnostico.observa1,Diagnostico.observa2,Diagnostico.observa3,Diagnostico.observa4,diagnostico.cliente,diagnostico.empresa FROM diagnostico,Enfermedad where diagnostico.enfermedad=enfermedad.enfermedad  "
    cad = cad & " and diagnostico.sede='" & sede & "'"
    cad = cad & " and diagnostico.consulta='" & consulta & "'"

    If rsdiag.State = 1 Then rsdiag.Close
    rsdiag.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rsdiag
    dbGrid1.columns(0).Width = 500
    dbGrid1.columns(1).Width = 1000
    dbGrid1.columns(2).Width = 1000
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 3000

    Exit Sub
cmd5_err:
    MsgBox "Aviso en sql " + error, 48, "Aviso"
    Exit Sub

End Sub

Private Sub ahyy1_Click()

    On Error GoTo cmd45_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Len(sede) = 0 Then
        MsgBox "Sede no existe ", 48, "Aviso"
        sede.SetFocus
        Exit Sub

    End If

    If Len(xcliente) = 0 Then
        MsgBox "No existe Cliente"
        consulta.SetFocus
        Exit Sub

    End If

    If Len(consulta) = 0 Then
        MsgBox "No existe Consulta"
        consulta.SetFocus
        Exit Sub

    End If

    dsede = sede
    dconsulta = consulta
    cliente = xcliente
    empresa = xempresa
    Frame1.Visible = True
    Frame1.Caption = "NUEVO"
    diagnostico = ""
    inicializa
    enfermedad.SetFocus
    Exit Sub
cmd45_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub inicializa()
    observa1 = ""
    observa2 = ""
    observa3 = ""
    observa4 = ""
    nenfermedad = ""

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        If opcion1 = 1 Then
            Frame2.Visible = False
            enfermedad.SetFocus
            Exit Sub

        End If

        If opcion1 = 2 Then
            Frame2.Visible = False
            sede.SetFocus
            Exit Sub

        End If

        If opcion1 = 3 Then
            Frame2.Visible = False
            consulta.SetFocus
            Exit Sub

        End If

        If opcion1 = 20 Then
            Frame2.Visible = False
            'xcliente.SetFocus
            Exit Sub

        End If

    End If

End Sub

Private Sub cmdAddEntry_Click()
    ahyy1_Click

End Sub

Private Sub cmdDelete_Click()
    dfj8221_Click

End Sub

Private Sub cmdExit_Click()
    fdo33_Click

End Sub

Private Sub cmdHelp_Click()
    dmi22_Click

End Sub

Private Sub cmdPrint_Click()
    dk281_Click

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa1.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_enfermedad

    End If

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Private Sub Command2_Click()
    SQL
    dbGrid1.SetFocus

End Sub

Private Sub Command3_Click()

    Dim found    As Integer

    Dim rs1      As New ADODB.Recordset

    Dim rsexiste As New ADODB.Recordset

    Dim cad      As String

    Dim sdx      As Double

    On Error GoTo cmd2_err

    If Len(enfermedad) = 0 Then
        enfermedad.SetFocus
        Exit Sub

    End If

    If Len(observa1) = 0 Then
        observa1.SetFocus
        Exit Sub

    End If

    If Frame1.Caption = "NUEVO" Then
        If rs1.State = 1 Then rs1.Close
        rs1.Open "SELECT Numerod FROM parame where codigo='01'", cn, adOpenKeyset, adLockOptimistic

        If rs1.RecordCount = 0 Then  'si existe
            MsgBox "No hay Consultas  ", 48, "Aviso"
            enfermedad.SetFocus
            Exit Sub

        End If

        sdx = Val("" & rs1.Fields("numerod").Value) + 1
siguen:
        diagnostico = "" & sdx

        If rsexiste.State = 1 Then rsexiste.Close
        rsexiste.Open "SELECT diagnostico,Consulta,Sede FROM diagnostico where consulta='" & Trim(dconsulta) & "' and sede='" & dsede & "' and diagnostico='" & diagnostico & "'", cn, adOpenKeyset, adLockOptimistic

        If rsexiste.RecordCount > 0 Then  'si existe
            sdx = sdx + 1
            GoTo siguen
            Exit Sub

        End If

        cad = "update parame set numerod='" & diagnostico & "' where codigo='01'"
        cn.Execute (cad)
        cad = "INSERT INTO diagnostico VALUES('" & Trim(diagnostico) & "','" & Trim(dsede) & "','" & Trim(dconsulta) & "','" & Trim(enfermedad) & "','" & Trim(observa1) & "','" & Trim(observa2) & "','" & Trim(observa3) & "','" & Trim(observa4) & "','" & Trim(cliente) & "','" & Trim(empresa) & "')"
        cn.Execute (cad)
        SQL
        dbGrid1.SetFocus
        fdo33_Click
        Exit Sub

    End If

    If Frame1.Caption = "MODIFICA" Then
        cad = "UPDATE diagnostico SET enfermedad = '" & Trim(enfermedad) & "', observa1= '" & Trim(observa1) & "', observa2= '" & Trim(observa2) & "', observa3= '" & Trim(observa3) & "', observa4= '" & Trim(observa4) & "' WHERE diagnostico = '" & Trim(diagnostico) & "' and sede='" & dsede & "' and consulta='" & dconsulta & "'"
        cn.Execute (cad)
        SQL
        dbGrid1.SetFocus
        fdo33_Click

    End If

    Exit Sub
cmd2_err:
    MsgBox "Aviso en command3 " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command4_Click()
    fdo33_Click

End Sub

Sub ejecuta(sw As Integer)

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    If opcion1 = 1 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Enfermedad FROM Enfermedad  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Enfermedad FROM Enfermedad where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 2 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Nombre,Codigo FROM Tlocal  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Codigo FROM Tlocal where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 3 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Clientes.nombre,Consulta.consulta,consulta.cliente,consulta.empresa FROM Consulta,clientes where consulta.cliente=clientes.codigo and consulta.sede='" & sede & "'"

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Clientes.nombre,Consulta.consulta,consulta.cliente,consulta.empresa FROM Consulta,clientes where consulta.cliente=clientes.codigo and  consulta.sede='" & sede & "' and " & Combo1 & " like '" & buffer & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

    If opcion1 = 20 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Clientes.Nombre,consulta.Consulta,consulta.fecha FROM Clientes,consulta  where clientes.codigo=consulta.cliente order by fecha"

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Clientes.Nombre,consulta.Consulta,consulta.fecha FROM Clientes,consulta where  clientes.codigo=consulta.cliente and " & Combo1 & " like '" & buffer & "%' order by fecha"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            buffer.SetFocus
            Exit Sub

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 5000
        DBGrid2.columns(1).Width = 1000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

    End If

End Sub

Private Sub consulta_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If Len(sede) > 0 Then
            consulta_consulta

        End If

    End If

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = 1 Then
            enfermedad = Trim(DBGrid2.columns(1))
            nenfermedad = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            observa1.SetFocus
            Exit Sub

        End If

        If opcion1 = 2 Then
            sede = Trim(DBGrid2.columns(1))
            Frame2.Visible = False
            Frame2.Enabled = False
            consulta.SetFocus
            Exit Sub

        End If

        If opcion1 = 3 Then
            consulta = Trim(DBGrid2.columns(1))
            xcliente = Trim(DBGrid2.columns(2))
            xempresa = Trim(DBGrid2.columns(3))
            xnombre = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            xnombree = existe_empresa()
            Command2_Click
            Exit Sub

        End If

        If opcion1 = 20 Then
            consulta = Trim(DBGrid2.columns(1))
      
            Frame2.Visible = False
            Frame2.Enabled = False
            Command2_Click
            Exit Sub

        End If

    End If

End Sub

Private Sub dbgrid2_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(buffer) > 0 Then
                buf = Mid$(buffer, 1, Len(buffer) - 1)
                buffer = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyAscii <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0

    End If

End Sub

Private Sub dfj8221_Click()

    Dim buf  As String

    Dim buf1 As String

    On Error GoTo cmd4_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    buf = Trim(dbGrid1.columns(0))

    If MsgBox("Desea Borrar " + dbGrid1.columns(1), 1, "Aviso") = 1 Then
        buf1 = "DELETE   FROM diagnostico WHERE diagnostico ='" & Trim(dbGrid1.columns(2)) & "' and sede='" & Trim(dbGrid1.columns(0)) & "' and consulta='" & Trim(dbGrid1.columns(1)) & "'"
        cn.Execute (buf1)
        rsdiag.Requery
        SQL
        dbGrid1.SetFocus

    End If

    dbGrid1.SetFocus
    Exit Sub
cmd4_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    dbGrid1.SetFocus
    Exit Sub

End Sub

Private Sub dk281_Click()
    'If Frame2.Visible = True Then Exit Sub
    'If Frame1.Visible = True Then Exit Sub
    'If rt.State = 1 Then rt.Close
    'rt.Open "SELECT * FROM diagnostico ", cn, adOpenKeyset, adLockOptimistic
    'Set trepcli1.DataSource = rt
    'trepcli1.Show 1

End Sub

Private Sub dmi22_Click()

    On Error GoTo cmd3_err

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    dsede = Trim(dbGrid1.columns(0))
    dconsulta = Trim(dbGrid1.columns(1))
    diagnostico = Trim(dbGrid1.columns(2))
    enfermedad = Trim(dbGrid1.columns(3))
    nenfermedad = Trim(dbGrid1.columns(4))
    observa1 = Trim(dbGrid1.columns(5))
    observa2 = Trim(dbGrid1.columns(6))
    observa3 = Trim(dbGrid1.columns(7))
    observa4 = Trim(dbGrid1.columns(8))
    cliente = Trim(dbGrid1.columns(9))
    empresa = Trim(dbGrid1.columns(10))
    Frame1.Visible = True
    Frame1.Caption = "MODIFICA"
    enfermedad.SetFocus
    Exit Sub
cmd3_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub enfermedad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa1.SetFocus

End Sub

Private Sub enfermedad_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_enfermedad

    End If

End Sub

Private Sub fdo33_Click()

    If Frame2.Visible = True Then
        buffer_KeyPress 27
        Exit Sub

    End If

    If Frame1.Visible = True Then
        If Frame1.Caption = "NUEVO" Then
            Frame1.Visible = False
            dbGrid1.SetFocus

        End If

        If Frame1.Caption = "MODIFICA" Then
            Frame1.Visible = False
            dbGrid1.SetFocus

        End If

        Exit Sub

    End If

    tdiagnos.Hide
    Unload tdiagnos

End Sub

Private Sub Form_Load()
    sede = glocal
    SQL

End Sub

Sub consulta_enfermedad()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM enfermedad  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Enfermedad"
    Combo1.ListIndex = 0
    opcion1 = 1
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_sede()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM sede  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Sede"
    Combo1.ListIndex = 0
    opcion1 = 2
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_consulta()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM consulta  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Consulta"
    Combo1.ListIndex = 0
    opcion1 = 3
    buffer.SetFocus
    Command1_Click

End Sub

Sub consulta_xcliente()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM clientes  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "codigo"
    Combo1.ListIndex = 0
    opcion1 = 20
    buffer.SetFocus
    Command1_Click

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub observa1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa2.SetFocus

End Sub

Private Sub observa2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa3.SetFocus

End Sub

Private Sub observa3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa4.SetFocus

End Sub

Private Sub observa4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub sede_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    consulta.SetFocus

End Sub

Private Sub sede_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_sede

    End If

End Sub

Private Sub tr633_Click()

End Sub

Function existe_empresa() As String

    Dim rs1 As New ADODB.Recordset
   
    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT nombre FROM clientes where codigo='" & xempresa & "'", cn, adOpenDynamic, adLockReadOnly

    If Not rs1.EOF Then
        existe_empresa = "" & rs1.Fields("nombre")

    End If

    rs1.Close
    Set rs1 = Nothing

End Function

