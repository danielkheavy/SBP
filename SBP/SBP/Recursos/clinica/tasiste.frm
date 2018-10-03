VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tasiste 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Asistencias "
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   7215
      Left            =   15
      TabIndex        =   24
      Top             =   45
      Visible         =   0   'False
      Width           =   11895
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid2 
         Height          =   5775
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   10186
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
      Height          =   615
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tasiste.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4575
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox fecha 
         Height          =   495
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   33
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox observa 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   31
         Top             =   3240
         Width           =   6135
      End
      Begin VB.TextBox horas 
         Height          =   495
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox asistencia 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox horai 
         Height          =   495
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2280
         Width           =   1575
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
         Picture         =   "tasiste.frx":07AE
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
         Picture         =   "tasiste.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox terapista 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label empresa 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6240
         TabIndex        =   48
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label diagnostico 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4800
         TabIndex        =   45
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label consulta 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3360
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label cliente 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Atencion"
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora Salida"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label dtratamiento 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dsede 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora Ingreso"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label nterapista 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   3240
         TabIndex        =   11
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label hy611 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atendido Por"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   11910
      TabIndex        =   3
      Top             =   0
      Width           =   11970
      Begin VB.TextBox sede 
         BackColor       =   &H00FFFF00&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox tratamiento 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4920
         MaxLength       =   11
         TabIndex        =   0
         Top             =   600
         Width           =   1215
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
         Picture         =   "tasiste.frx":170A
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
         Picture         =   "tasiste.frx":291C
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
         Picture         =   "tasiste.frx":3B2E
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
         Picture         =   "tasiste.frx":4D40
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
         Picture         =   "tasiste.frx":5F52
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label xempresa 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7200
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
      Begin VB.Label xnombree 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8160
         TabIndex        =   46
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diagnostico"
         Height          =   375
         Left            =   8160
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Label xdiagnostico 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9120
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Consulta"
         Height          =   375
         Left            =   6240
         TabIndex        =   40
         Top             =   960
         Width           =   975
      End
      Begin VB.Label xconsulta 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7200
         TabIndex        =   39
         Top             =   960
         Width           =   975
      End
      Begin VB.Label xnombre 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8160
         TabIndex        =   38
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label xcliente 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7200
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empresa"
         Height          =   375
         Left            =   6240
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
         Height          =   375
         Left            =   6240
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sede"
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tratamiento"
         Height          =   375
         Left            =   3960
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   6375
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11245
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
Attribute VB_Name = "tasiste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rt     As New ADODB.Recordset

Public rsdiag As New ADODB.Recordset

Private Sub SQL()

    On Error GoTo cmd5_err

    Dim cad As String

    cad = "SELECT Sede,Asistencia,Tratamiento,Terapista,Fecha,Horai,Horas,Observa,cliente,consulta,diagnostico,empresa FROM  asistencia where  "
    cad = cad & " sede='" & sede & "'"
    cad = cad & " and tratamiento='" & tratamiento & "' order by fecha"

    If rsdiag.State = 1 Then rsdiag.Close
    rsdiag.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rsdiag
    dbGrid1.columns(0).Width = 500
    dbGrid1.columns(1).Width = 900
    dbGrid1.columns(2).Width = 900
    dbGrid1.columns(3).Width = 900

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

    If Len(tratamiento) = 0 Then
        MsgBox "Tratamiento no existe ", 48, "Aviso"
        tratamiento.SetFocus
        Exit Sub

    End If

    If Len(xcliente) = 0 Then
        MsgBox "Cliente no existe ", 48, "Aviso"
        tratamiento.SetFocus
        Exit Sub

    End If

    dsede = sede
    empresa = xempresa
    dtratamiento = tratamiento
    cliente = xcliente
    consulta = xconsulta
    diagnostico = xdiagnostico
    Frame1.Visible = True
    Frame1.Caption = "NUEVO"
    asistencia = ""
    inicializa
    fecha = Format(Now, "dd/mm/yyyy")
    horai = Format(Now, "hh:mm:ss")
    horas = Format(Now, "hh:mm:ss")
    terapista.SetFocus
    Exit Sub
cmd45_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub inicializa()
    terapista = ""
    nterapista = ""
    fecha = ""
    horai = ""
    horas = ""
    observa = ""

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        If opcion1 = 1 Then
            Frame2.Visible = False
            terapista.SetFocus
            Exit Sub

        End If

        If opcion1 = 2 Then
            Frame2.Visible = False
            sede.SetFocus
            Exit Sub

        End If

        If opcion1 = 3 Then
            Frame2.Visible = False
            tratamiento.SetFocus
            Exit Sub

        End If

        If opcion1 = 4 Then
            Frame2.Visible = False
      
            Exit Sub

        End If

    End If

End Sub

Private Sub cantidad_Change()

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    horai.SetFocus

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

    If Len(terapista) = 0 Then
        terapista.SetFocus
        Exit Sub

    End If

    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Sub

    End If

    If Frame1.Caption = "NUEVO" Then
        If rs1.State = 1 Then rs1.Close
        rs1.Open "SELECT Numeroa FROM parame where codigo='01'", cn, adOpenKeyset, adLockOptimistic

        If rs1.RecordCount = 0 Then  'si existe
            MsgBox "No hay sedes  ", 48, "Aviso"
            terapista.SetFocus
            Exit Sub

        End If

        sdx = Val("" & rs1.Fields("numeroa").Value) + 1
siguen:
        asistencia = "" & sdx

        If rsexiste.State = 1 Then rsexiste.Close
        rsexiste.Open "SELECT asistencia,tratamiento,Sede FROM asistencia where asistencia='" & Trim(asistencia) & "' and sede='" & dsede & "' and tratamiento='" & dtratamiento & "'", cn, adOpenKeyset, adLockOptimistic

        If rsexiste.RecordCount > 0 Then  'si existe
            sdx = sdx + 1
            GoTo siguen
            Exit Sub

        End If

        cad = "update parame set numeroa='" & asistencia & "' where codigo='01'"
        cn.Execute (cad)
        cad = "INSERT INTO asistencia VALUES('" & Trim(asistencia) & "','" & Trim(dsede) & "','" & Trim(dtratamiento) & "','" & Trim(terapista) & "','" & Trim(fecha) & "','" & Trim(horai) & "','" & Trim(horas) & "','" & Trim(observa) & "','" & Trim(cliente) & "','" & Trim(consulta) & "','" & Trim(diagnostico) & "','" & Trim(empresa) & "')"
        cn.Execute (cad)
        SQL
        fdo33_Click
        Exit Sub

    End If

    If Frame1.Caption = "MODIFICA" Then
        cad = "UPDATE asistencia SET terapista = '" & Trim(terapista) & "', fecha= '" & Trim(fecha) & "', horai= '" & Trim(horai) & "', horas= '" & Trim(horas) & "', observa= '" & Trim(observa) & "' WHERE tratamiento = '" & Trim(dtratamiento) & "' and sede='" & dsede & "' and asistencia='" & asistencia & "'"
        cn.Execute (cad)
        SQL
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
            cad = "SELECT Nombre,Codigo FROM Vendedor  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Codigo FROM vendedor where  " & Combo1 & " like '" & buffer & "%'"

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
            cad = "SELECT Nombre,Sede FROM Sede  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Nombre,Sede FROM Sede where  " & Combo1 & " like '" & buffer & "%'"

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
            cad = "SELECT clientes.Nombre,tratamiento.Tratamiento,tratamiento.diagnostico,tratamiento.consulta,tratamiento.cliente,tratamiento.empresa FROM tratamiento,clientes where tratamiento.cliente=clientes.codigo and  tratamiento.sede='" & sede & "' "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT clientes.Nombre,tratamiento.Tratamiento,tratamiento.diagnostico,tratamiento.consulta,tratamiento.cliente,tratamiento.empresa FROM tratamiento,clientes where tratamiento.cliente=clientes.codigo and  tratamiento.sede='" & sede & "' and " & Combo1 & " like '" & buffer & "%' "

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

    If opcion1 = 4 Then  'clientes
        If Len(buffer) = 0 Then
            cad = "SELECT Clientes.nombre,tratamiento.Tratamiento,Diagnostico.Diagnostico,Consulta.Consulta,Consulta.fecha FROM clientes,diagnostico,consulta,tratamiento where clientes.codigo=consulta.cliente and diagnostico.consulta=consulta.consulta and tratamiento.diagnostico=diagnostico.diagnostico  and consulta.sede='" & sede & "' order by consulta.fecha"

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT Clientes.nombre,tratamiento.Tratamiento,Diagnostico.Diagnostico,Consulta.Consulta,Consulta.fecha FROM clientes,diagnostico,consulta,tratamiento where clientes.codigo=consulta.cliente and diagnostico.consulta=consulta.consulta and tratamiento.diagnostico=diagnostico.diagnostico  and consulta.sede='" & sede & "' and " & Combo1 & " like '" & buffer & "%' order by consulta.fecha"

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

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = 1 Then
            terapista = Trim(DBGrid2.columns(1))
            nterapista = Trim(DBGrid2.columns(0))
            Frame2.Visible = False
            Frame2.Enabled = False
            fecha.SetFocus
            Exit Sub

        End If

        If opcion1 = 2 Then
            sede = Trim(DBGrid2.columns(1))
            Frame2.Visible = False
            Frame2.Enabled = False
            tratamiento.SetFocus
            Exit Sub

        End If

        If opcion1 = 3 Then
            tratamiento = Trim(DBGrid2.columns(1))
            xnombre = Trim(DBGrid2.columns(0))
            xdiagnostico = Trim(DBGrid2.columns(2))
            xconsulta = Trim(DBGrid2.columns(3))
            xcliente = Trim(DBGrid2.columns(4))
            Frame2.Visible = False
            Frame2.Enabled = False
            Command2_Click
            Exit Sub

        End If

        If opcion1 = 4 Then
            tratamiento = Trim(DBGrid2.columns(1))
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
        buf1 = "DELETE   FROM asistencia WHERE asistencia ='" & Trim(dbGrid1.columns(1)) & "' and sede='" & Trim(dbGrid1.columns(0)) & "' and tratamiento='" & Trim(dbGrid1.columns(2)) & "'"
        cn.Execute (buf1)
        rsdiag.Requery
        SQL

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
    asistencia = Trim(dbGrid1.columns(1))
    dtratamiento = Trim(dbGrid1.columns(2))
    terapista = Trim(dbGrid1.columns(3))
    fecha = Trim(dbGrid1.columns(4))
    horai = Trim(dbGrid1.columns(5))
    horas = Trim(dbGrid1.columns(6))
    observa = Trim(dbGrid1.columns(7))
    cliente = Trim(dbGrid1.columns(8))
    consulta = Trim(dbGrid1.columns(9))
    diagnostico = Trim(dbGrid1.columns(10))
    empresa = Trim(dbGrid1.columns(11))
    Frame1.Visible = True
    Frame1.Caption = "MODIFICA"
    terapista.SetFocus
    Exit Sub
cmd3_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

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

    tasiste.Hide
    Unload tasiste

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    horai.SetFocus

End Sub

Private Sub Form_Load()
    sede = glocal
    SQL

End Sub

Sub consulta_terapista()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM vendedor  "

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
    Combo1.AddItem "personal"
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

Sub consulta_tratamiento()

    Dim cad       As String

    Dim rconsulta As New ADODB.Recordset

    cad = "SELECT * FROM tratamiento  "

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True Or rconsulta.BOF = True Then
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Enabled = True
    buffer = ""
    Combo1.Clear
    Combo1.AddItem "diagnostico"
    Combo1.AddItem "tratamiento"
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
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    opcion1 = 4
    buffer.SetFocus
    Command1_Click

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub observa1_Change()

End Sub

Private Sub horai_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    horas.SetFocus

End Sub

Private Sub horas_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa.SetFocus

End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub sede_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    tratamiento.SetFocus

End Sub

Private Sub sede_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_sede

    End If

End Sub

Private Sub tr633_Click()

End Sub

Private Sub terapista_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fecha.SetFocus

End Sub

Private Sub terapista_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_terapista

    End If

End Sub

Private Sub tratamiento_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If Len(sede) > 0 Then
            consulta_tratamiento

        End If

    End If

End Sub

