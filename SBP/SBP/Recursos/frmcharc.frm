VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmCharc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graficos Estadisticos"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   14505
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Condicion Busqueda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   15
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   11610
      Begin VB.ComboBox turno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3960
         Width           =   2655
      End
      Begin VB.ComboBox caja 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   3600
         Width           =   2655
      End
      Begin VB.ComboBox cajero 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox acu 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   33
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2880
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   5640
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmcharc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmcharc.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox producto 
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   18
         Text            =   "*"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox familia 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox Vendedor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2520
         Width           =   2655
      End
      Begin VB.ComboBox moneda 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2160
         Width           =   2655
      End
      Begin VB.ComboBox tipo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   13
         Text            =   "%"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "C.ompras V.entas"
         Height          =   375
         Left            =   3480
         TabIndex        =   41
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acu"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servicio"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoAnalisis"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Familia"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   14055
      Begin VB.TextBox fechaii 
         Height          =   375
         Left            =   11640
         MaxLength       =   10
         TabIndex        =   44
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox fechafi 
         Height          =   375
         Left            =   11640
         MaxLength       =   10
         TabIndex        =   43
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton CmdSaveas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grabar Como..."
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin MSChart20Lib.MSChart msChart1 
         Height          =   7335
         Left            =   120
         OleObjectBlob   =   "frmcharc.frx":0F5C
         TabIndex        =   9
         Top             =   240
         Width           =   4815
      End
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   6855
         Left            =   6720
         OleObjectBlob   =   "frmcharc.frx":2836
         TabIndex        =   42
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   11640
         TabIndex        =   46
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         Height          =   375
         Left            =   11640
         TabIndex        =   45
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmcharc.frx":4110
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label docu 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10080
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Reporte"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu fdlo332 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "FrmCharc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X As Integer

Private Sub cmdGrabar_Click()

    If Frame2.Visible = True Then
        fechai.SetFocus
        Exit Sub

    End If

    Frame2.Visible = True
    fechai = Format(Now, "dd/mm/yyyy") '"01/01/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    fechaii = Format(Now, "dd/mm/yyyy") '"01/01/" & Format(Year(Now), "0000")
    fechafi = Format(Now, "dd/mm/yyyy")

    tipo.ListIndex = 0
    moneda.ListIndex = 0
    vendedor.ListIndex = 0
    familia.ListIndex = 0
    Combo2.ListIndex = 0
    Label8.Visible = False
    Label9.Visible = False
    producto.Visible = False
    familia.Visible = False

    If Combo1 = "Productos" Then
        Label8.Visible = True
        Label9.Visible = True
        producto.Visible = True
        familia.Visible = True

    End If

    'fechai.SetFocus

End Sub

Private Sub cmdPrint_Click()

    If Frame2.Visible = True Then Exit Sub
    If MsgBox("Desea Imprimir", 1, "Aviso") <> 1 Then Exit Sub
    msChart1.EditCopy 'This Makes MSChart Control to be Copied
    DoEvents   ' may be needed for large datasets
    Printer.Print " "
    Picture1.PaintPicture Clipboard.GetData(), 0, 0
    Printer.EndDoc

End Sub

Private Sub CmdSaveas_Click()

    On Error GoTo Hell

    Dim strsavefile As String

    If Frame2.Visible = True Then Exit Sub

    With CommonDialog1
        .Filter = "Pictures (*.bmp)|*.bmp" ' You can Also Save the Pic in JPG/GIF/TIFF
        .DefaultExt = "bmp"
        .CancelError = False
        .ShowSave
        strsavefile = .FileName

        If strsavefile = "" Then Exit Sub

    End With

    msChart1.EditCopy
    SavePicture Clipboard.GetData, strsavefile ' File Saved
    Exit Sub
Hell:
    MsgBox Err.Description

End Sub

Private Sub Command1_Click()

    fdlo332_Click

End Sub

Private Sub Command2_Click()

    Dim buf          As String

    Dim I            As Integer

    Dim buf2         As String

    Dim mytablex     As New ADODB.Recordset

    Dim sdx          As Double

    Dim X            As Long

    Dim z            As Long

    Dim Y            As Integer

    Dim k            As Integer

    Dim Tmp          As String

    Dim sw           As Integer

    Dim cambuf       As String

    Dim dambuf       As String

    Dim xxx          As Integer

    Dim arraychart() As Single

    cambuf = "factura"
    dambuf = "detalle"
    'If deliveri.Value = True Then
    '   cambuf = "cpedidov"
    '   dambuf = "dpedidov"
    'End If
    buf = ""
    buf2 = ""

    If Combo1 = "Documentos" Or Combo1 = "Clientes" Or Combo1 = "Vendedor" Then
        If Combo2 = "Tipo" Then
            buf = "select Tipo as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Tipo"

        End If

        If Combo2 = "Usuario" Then
            buf = "select Usuario as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Usuario"

        End If

        If Combo2 = "Bodega" Then
            buf = "select Bodega as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Bodega"

        End If

        If Combo2 = "Caja" Then
            buf = "select Caja as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Caja"

        End If

        If Combo2 = "Vendedor" Then
            buf = "select Vendedor as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Vendedor"

        End If

        If Combo2 = "Anual" Then
            buf = "select year(fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "year(fecha)"

        End If

        If Combo2 = "Mensual" Then
            buf = "select month(fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "month(fecha)"

        End If

        If Combo2 = "Semanal" Then
            buf = "select DATENAME(weekday,fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "DATENAME(weekday,fecha)"

        End If

        If Combo2 = "Diario" Then
            buf = "select day(fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "day(fecha)"

        End If

        If Combo2 = "Horario" Then
            buf = "select left(hora,2) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = " left(hora,2) "

        End If

        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
   
        If tipo <> "%" Then
            buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"

        End If

        If codigo <> "%" Then
            buf = buf & " and codigo like '" & codigo & "'"

        End If
   
        If moneda <> "%" Then
            buf = buf & " and moneda like '" & moneda & "'"

        End If

        If caja <> "%" Then
            buf = buf & " and caja like '" & caja & "'"

        End If

        If turno <> "%" Then
            buf = buf & " and turno like '" & turno & "'"

        End If

        If cajero <> "%" Then
            buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

        End If

        If vendedor <> "%" Then
            buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

        End If

        If Combo3 <> "Todos" Then
            'If Combo3 = "Autoservicio" Then
            buf = buf & " and servicio='" & extra_loquesea(Combo3) & "'"

            'End If
            'If Combo3 = "Comanda" Then
            '   buf = buf & " and servicio='C'"
            'End If
            'If Combo3 = "Deliveri" Then
            '   buf = buf & " and servicio='D'"
            'End If
        End If

        If acu = "V" Then
            buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

        End If

        If acu = "C" Then
            buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

        End If

        buf = buf & "  group by " & buf2 & ",moneda "
      
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.EOF = True And mytablex.BOF = True Then
            Exit Sub

        End If

        X = 0
        X = mytablex.RecordCount
        msChart1.RowCount = X

        For I = 1 To X
            msChart1.Row = I
            msChart1.Data = Val("" & mytablex.Fields("xtotal"))
            msChart1.RowLabel = "" & mytablex.Fields("mes")
            mytablex.MoveNext
        Next I

        mytablex.Close
        msChart1.ColumnCount = 1
        msChart1.Title = Combo1 & " (" & Combo2 & " " & fechai & " " & fechaf & " )"
        Frame2.Visible = False

        If Not IsDate(fechaii) Then Exit Sub
        If Not IsDate(fechafi) Then Exit Sub
        graficos_paraali
        Exit Sub

    End If

    If Combo1 = "Productos" Then

        '--------------
        If Combo2 = "Tipo" Then
            buf = "select Tipo as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Tipo"

        End If

        If Combo2 = "Usuario" Then
            buf = "select Usuario as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Usuario"

        End If

        If Combo2 = "Bodega" Then
            buf = "select Bodega as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Bodega"

        End If

        If Combo2 = "Caja" Then
            buf = "select Caja as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Caja"

        End If

        If Combo2 = "Vendedor" Then
            buf = "select Vendedor as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Vendedor"

        End If

        If Combo2 = "Anual" Then
            buf = "select year(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "year(fecha)"

        End If

        If Combo2 = "Mensual" Then
            buf = "select month(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "month(fecha)"

        End If

        If Combo2 = "Semanal" Then
            buf = "select weekday(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "weekday(fecha)"

        End If

        If Combo2 = "Diario" Then
            buf = "select day(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "day(fecha)"

        End If

        If Combo2 = "Horario" Then
            buf = "select left(hora,2) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = " left(hora,2)"

        End If

        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "'"

        If tipo <> "%" Then
            buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"

        End If

        If moneda <> "%" Then
            buf = buf & " and moneda like '" & moneda & "'"

        End If

        If codigo <> "%" Then
            buf = buf & " and codigo like '" & codigo & "'"

        End If

        If caja <> "%" Then
            buf = buf & " and caja like '" & caja & "'"

        End If

        If turno <> "%" Then
            buf = buf & " and turno like '" & turno & "'"

        End If

        If cajero <> "%" Then
            buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

        End If

        If vendedor <> "%" Then
            buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

        End If

        If Combo1 = "Productos" Then
            If producto <> "%" Then
                buf = buf & " and producto like '" & producto & "'"

            End If

            If familia <> "%" Then
                buf = buf & " and familia like '" & extra_loquesea(producto) & "'"

            End If

        End If

        If dambuf <> "dpedidov" Then
            If acu = "V" Then
                buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F')"

            End If

            If acu = "C" Then
                buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O')"

            End If

        End If

        buf = buf & "  group by " & buf2 & ",moneda "

        'MsgBox buf
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.EOF = True And mytablex.BOF = True Then
            Exit Sub

        End If

        '-------------------------
        X = 0
        X = mytablex.RecordCount
        msChart1.RowCount = X

        For I = 1 To X
            msChart1.Row = I
            msChart1.Data = Val("" & mytablex.Fields("xtotal"))
            msChart1.RowLabel = "" & mytablex.Fields("mes")
            mytablex.MoveNext
        Next I

        mytablex.Close
        msChart1.ColumnCount = 1
        msChart1.Title = Combo1 & " (" & Combo2 & " " & fechai & " " & fechaf & " )"
        '-------------------------
        Frame2.Visible = False
    
        If Not IsDate(fechaii) Then Exit Sub
        If Not IsDate(fechafi) Then Exit Sub
        graficos_paraali
        Exit Sub

    End If

End Sub

Sub inicial()
    ReDim arraychart(1 To 1, 1 To 1) 'Array
    arraychart(1, 1) = 0
    msChart1.RowCount = 1
    msChart1.ColumnCount = 1
    msChart1.Title = ""
    msChart1.ChartData = arraychart

End Sub

Private Sub Command3_Click()
    Frame2.Visible = False
    inicial

End Sub

Private Sub fdlo332_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    FrmCharc.Hide
    Unload FrmCharc

End Sub

Private Sub Form_Activate()
    Frame1.Top = 10: Frame1.Left = 10
    Frame2.Top = 10: Frame2.Left = 10

    If docu = "1" Then
        Combo1.ListIndex = 1

    End If

    If docu = "2" Then  'clientes
        Combo1.ListIndex = 2

    End If

    inicial

End Sub

Private Sub Form_Load()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    carga_datos
    Combo3.Clear
    Combo3.AddItem "Todos"
    mytablex.Open "SELECT * FROM servicio ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        Combo3.AddItem "" & mytablex.Fields("servicio") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    Combo3.ListIndex = 0
    mytablex.Close
    Combo1.Clear
    Combo1.AddItem "Documentos"
    Combo1.AddItem "Productos"
    Combo1.AddItem "Clientes"
    'Combo1.AddItem "Vendedor"
    Combo1.ListIndex = 0
    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    Combo2.Clear
    Combo2.AddItem "Caja"
    Combo2.AddItem "Anual"
    Combo2.AddItem "Mensual"
    Combo2.AddItem "Semanal"
    Combo2.AddItem "Diario"
    Combo2.AddItem "Horario"
    Combo2.AddItem "Vendedor"
    Combo2.AddItem "Tipo"
    Combo2.AddItem "Usuario"
    Combo2.AddItem "Bodega"

    Combo2.ListIndex = 0
    producto = "%"
    codigo = "%"
    cmdGrabar_Click

End Sub

Sub carga_datos()

    Dim mytablex As New ADODB.Recordset

    tipo.Clear
    vendedor.Clear
    familia.Clear
    cajero.Clear
    caja.Clear
    turno.Clear
    familia.AddItem "%"
    tipo.AddItem "%"
    vendedor.AddItem "%"
    cajero.AddItem "%"
    caja.AddItem "%"
    turno.AddItem "%"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parameca", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        caja.AddItem "" & mytablex.Fields("caja")
        mytablex.MoveNext
    Loop

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from turno", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno")
        mytablex.MoveNext
    Loop

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from tipo", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from vendedor", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("NOMBRE")
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("NOMBRE")
        mytablex.MoveNext
    Loop

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from familia", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True And mytablex.BOF = True Then
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        familia.AddItem "" & mytablex.Fields("familia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    caja.ListIndex = 0
    turno.ListIndex = 0
    tipo.ListIndex = 0
    vendedor.ListIndex = 0
    familia.ListIndex = 0

End Sub

Sub graficos_paraali()

    Dim buf          As String

    Dim I            As Integer

    Dim buf2         As String

    Dim mytablex     As New ADODB.Recordset

    Dim sdx          As Double

    Dim X            As Long

    Dim z            As Long

    Dim Y            As Integer

    Dim k            As Integer

    Dim Tmp          As String

    Dim sw           As Integer

    Dim cambuf       As String

    Dim dambuf       As String

    Dim xxx          As Integer

    Dim arraychart() As Single

    cambuf = "factura"
    dambuf = "detalle"
    'If deliveri.Value = True Then
    '   cambuf = "cpedidov"
    '   dambuf = "dpedidov"
    'End If
    buf = ""
    buf2 = ""

    If Combo1 = "Documentos" Or Combo1 = "Clientes" Or Combo1 = "Vendedor" Then
        If Combo2 = "Tipo" Then
            buf = "select Tipo as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Tipo"

        End If

        If Combo2 = "Usuario" Then
            buf = "select Usuario as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Usuario"

        End If

        If Combo2 = "Bodega" Then
            buf = "select Bodega as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Bodega"

        End If

        If Combo2 = "Caja" Then
            buf = "select Caja as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Caja"

        End If

        If Combo2 = "Vendedor" Then
            buf = "select Vendedor as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "Vendedor"

        End If

        If Combo2 = "Anual" Then
            buf = "select year(fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "year(fecha)"

        End If

        If Combo2 = "Mensual" Then
            buf = "select month(fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "month(fecha)"

        End If

        If Combo2 = "Semanal" Then
            buf = "select DATENAME(weekday,fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "DATENAME(weekday,fecha)"

        End If

        If Combo2 = "Diario" Then
            buf = "select day(fecha) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = "day(fecha)"

        End If

        If Combo2 = "Horario" Then
            buf = "select left(hora,2) as mes,moneda,sum(total) as xtotal  from " & cambuf & " where "
            buf2 = " left(hora,2) "

        End If

        buf = buf & "  fecha>='" & Format(fechaii, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechafi, "YYYYMMDD") & "'"
   
        If tipo <> "%" Then
            buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"

        End If

        If codigo <> "%" Then
            buf = buf & " and codigo like '" & codigo & "'"

        End If
   
        If moneda <> "%" Then
            buf = buf & " and moneda like '" & moneda & "'"

        End If

        If caja <> "%" Then
            buf = buf & " and caja like '" & caja & "'"

        End If

        If turno <> "%" Then
            buf = buf & " and turno like '" & turno & "'"

        End If

        If cajero <> "%" Then
            buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

        End If

        If vendedor <> "%" Then
            buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

        End If

        If Combo3 <> "Todos" Then
            'If Combo3 = "Autoservicio" Then
            buf = buf & " and servicio='" & extra_loquesea(Combo3) & "'"

            'End If
            'If Combo3 = "Comanda" Then
            '   buf = buf & " and servicio='C'"
            'End If
            'If Combo3 = "Deliveri" Then
            '   buf = buf & " and servicio='D'"
            'End If
        End If

        If acu = "V" Then
            buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

        End If

        If acu = "C" Then
            buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

        End If

        buf = buf & "  group by " & buf2 & ",moneda "
      
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.EOF = True And mytablex.BOF = True Then
            Exit Sub

        End If

        X = 0
        X = mytablex.RecordCount
        MSChart2.RowCount = X

        For I = 1 To X
            MSChart2.Row = I
            MSChart2.Data = Val("" & mytablex.Fields("xtotal"))
            MSChart2.RowLabel = "" & mytablex.Fields("mes")
            mytablex.MoveNext
        Next I

        mytablex.Close
        MSChart2.ColumnCount = 1
        MSChart2.Title = Combo1 & " (" & Combo2 & " " & fechaii & " " & fechafi & " )"
        Frame2.Visible = False
        Exit Sub

    End If

    If Combo1 = "Productos" Then

        '--------------
        If Combo2 = "Tipo" Then
            buf = "select Tipo as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Tipo"

        End If

        If Combo2 = "Usuario" Then
            buf = "select Usuario as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Usuario"

        End If

        If Combo2 = "Bodega" Then
            buf = "select Bodega as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Bodega"

        End If

        If Combo2 = "Caja" Then
            buf = "select Caja as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Caja"

        End If

        If Combo2 = "Vendedor" Then
            buf = "select Vendedor as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "Vendedor"

        End If

        If Combo2 = "Anual" Then
            buf = "select year(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "year(fecha)"

        End If

        If Combo2 = "Mensual" Then
            buf = "select month(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "month(fecha)"

        End If

        If Combo2 = "Semanal" Then
            buf = "select weekday(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "weekday(fecha)"

        End If

        If Combo2 = "Diario" Then
            buf = "select day(fecha) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = "day(fecha)"

        End If

        If Combo2 = "Horario" Then
            buf = "select left(hora,2) as mes,moneda,sum(total) as xtotal  from " & dambuf & " where "
            buf2 = " left(hora,2)"

        End If

        buf = buf & "  fecha>='" & Format(fechaii, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechafi, "YYYYMMDD") & "'"

        If tipo <> "%" Then
            buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"

        End If

        If moneda <> "%" Then
            buf = buf & " and moneda like '" & moneda & "'"

        End If

        If codigo <> "%" Then
            buf = buf & " and codigo like '" & codigo & "'"

        End If

        If caja <> "%" Then
            buf = buf & " and caja like '" & caja & "'"

        End If

        If turno <> "%" Then
            buf = buf & " and turno like '" & turno & "'"

        End If

        If cajero <> "%" Then
            buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

        End If

        If vendedor <> "%" Then
            buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

        End If

        If Combo1 = "Productos" Then
            If producto <> "%" Then
                buf = buf & " and producto like '" & producto & "'"

            End If

            If familia <> "%" Then
                buf = buf & " and familia like '" & extra_loquesea(producto) & "'"

            End If

        End If

        If dambuf <> "dpedidov" Then
            If acu = "V" Then
                buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F')"

            End If

            If acu = "C" Then
                buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O')"

            End If

        End If

        buf = buf & "  group by " & buf2 & ",moneda "

        'MsgBox buf
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.EOF = True And mytablex.BOF = True Then
            Exit Sub

        End If

        '-------------------------
        X = 0
        X = mytablex.RecordCount
        MSChart2.RowCount = X

        For I = 1 To X
            MSChart2.Row = I
            MSChart2.Data = Val("" & mytablex.Fields("xtotal"))
            MSChart2.RowLabel = "" & mytablex.Fields("mes")
            mytablex.MoveNext
        Next I

        mytablex.Close
        MSChart2.ColumnCount = 1
        MSChart2.Title = Combo1 & " (" & Combo2 & " " & fechaii & " " & fechafi & " )"
        '-------------------------
        Frame2.Visible = False
        Exit Sub

    End If

End Sub

