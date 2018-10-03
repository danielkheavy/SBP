VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form facmesa 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturacion Mensual"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14415
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Busqueda"
      Height          =   8670
      Left            =   -5115
      TabIndex        =   12
      Top             =   7725
      Visible         =   0   'False
      Width           =   15210
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   7575
         Left            =   150
         TabIndex        =   16
         Top             =   750
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   13361
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
      Begin VB.ComboBox xbuffer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox cadena 
         Height          =   375
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin ChamaleonButton.ChameleonBtn ChaAceptar 
         Height          =   930
         Left            =   12705
         TabIndex        =   19
         Top             =   690
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1640
         BTYPE           =   5
         TX              =   "Aceptar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "facmesua.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChaCerrar 
         Height          =   750
         Left            =   12930
         TabIndex        =   20
         Top             =   2010
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1323
         BTYPE           =   5
         TX              =   "Cerrar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "facmesua.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn Command13 
         Height          =   495
         Left            =   6870
         TabIndex        =   21
         Top             =   165
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "Buscar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "facmesua.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condicion"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   6495
      Left            =   90
      TabIndex        =   11
      Top             =   1050
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   23
      FormatLocked    =   -1  'True
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "X"
         Caption         =   "X"
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
         DataField       =   "Estado"
         Caption         =   "E"
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
      BeginProperty Column02 
         DataField       =   "Local"
         Caption         =   "Local"
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
      BeginProperty Column03 
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
      BeginProperty Column04 
         DataField       =   "Serie"
         Caption         =   "Serie"
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
      BeginProperty Column05 
         DataField       =   "Numero"
         Caption         =   "Numero"
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
      BeginProperty Column06 
         DataField       =   "Nro"
         Caption         =   "Nro"
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
      BeginProperty Column07 
         DataField       =   "Fecha"
         Caption         =   "Fecha"
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
      BeginProperty Column08 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
      BeginProperty Column09 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
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
      BeginProperty Column10 
         DataField       =   "Moneda"
         Caption         =   "M"
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
      BeginProperty Column11 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2759.811
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1184.882
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proceder a facturar"
      Height          =   735
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   3735
   End
   Begin VB.TextBox nombre 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4665
      MaxLength       =   10
      TabIndex        =   7
      Top             =   690
      Width           =   3975
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   4680
      MaxLength       =   13
      TabIndex        =   5
      Top             =   285
      Width           =   1695
   End
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   3
      Top             =   675
      Width           =   1335
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   1
      Top             =   315
      Width           =   1335
   End
   Begin ChamaleonButton.ChameleonBtn Command1 
      Height          =   840
      Left            =   9585
      TabIndex        =   18
      Top             =   165
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1482
      BTYPE           =   5
      TX              =   "&Buscar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "facmesua.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   420
      Left            =   6510
      Picture         =   "facmesua.frx":0070
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Buscar cliente"
      Top             =   240
      Width           =   660
   End
   Begin VB.Label total 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   375
      Left            =   3105
      TabIndex        =   6
      Top             =   705
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   3090
      TabIndex        =   4
      Top             =   315
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   675
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   315
      Width           =   1575
   End
   Begin VB.Menu dlo22 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "facmesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public txmensual As New ADODB.Recordset

Private Sub cadena_DblClick()
    Command13_Click

End Sub

Private Sub cadena_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame3.Visible = False
        codigo.SetFocus
        Exit Sub

    End If

    Command13_Click

End Sub

Private Sub ChaAceptar_Click()
    codigo = dbgrid3.columns(1)
    nombre = dbgrid3.columns(0)
    Frame3.Visible = False
    codigo.SetFocus
    codigo_KeyPress 13

End Sub

Private Sub ChaCERRAR_Click()
    Frame3.Visible = False
    codigo.SetFocus

End Sub

Private Sub cmdBuscar_Click()
    consulta_cliente

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command1_Click

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_cliente

    End If

End Sub

Sub consulta_cliente()

    Dim found As Integer

    xbuffer.Clear
    xbuffer.AddItem "%"
    xbuffer.AddItem "Nombre"
    xbuffer.AddItem "Codigo"
    xbuffer.ListIndex = 1
    opcion1 = "1"
    Frame3.Visible = True
    Frame3.Enabled = True
    buffer = ""
    ejecuta 1
    dbgrid3.SetFocus

End Sub

Sub busca_facturas()

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Sub

    End If

    If Len(nombre) = 0 Then
        codigo.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        fechaf.SetFocus
        Exit Sub

    End If

    xbuffer.Clear
    xbuffer.AddItem "Tipo"
    opcion1 = "2"
    Frame3.Visible = True
    cadena = ""
    cadena.SetFocus
    Command13_Click

End Sub

Private Sub Command1_Click()
    consulta_cuentac

End Sub

Private Sub Command13_Click()
    ejecuta 1

End Sub

Private Sub Command2_Click()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim I        As Integer

    Dim sw       As Integer

    Dim buf      As String

    If MsgBox("Esta seguro?", 1, "Se borran datos") <> 1 Then Exit Sub
    sw = 0
    found = ir_inicio()

    If found = 0 Then
        MsgBox "No existen Datos ", 48, "Aviso"
        Exit Sub

    End If

    tdeliver.crucefa.Clear
    Do

        If txmensual.EOF Then Exit Do
        If "" & txmensual.Fields("x") = "S" Then
            buf = "select * from detalle where local='" & "" & txmensual.Fields("local") & "'"
            buf = buf & " and tipo='" & "" & txmensual.Fields("tipo") & "'"
            buf = buf & " and dua is null and serie='" & "" & txmensual.Fields("serie") & "'"
            buf = buf & " and numero='" & "" & txmensual.Fields("numero") & "'"
            mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount > 0 Then
                tdeliver.crucefa.AddItem "" & txmensual.Fields("local") & "|" & txmensual.Fields("tipo") & "|" & txmensual.Fields("serie") & "|" & txmensual.Fields("numero") & "|"
                tptovta.crucefa.AddItem "" & txmensual.Fields("local") & "|" & txmensual.Fields("tipo") & "|" & txmensual.Fields("serie") & "|" & txmensual.Fields("numero") & "|"
        
                Do

                    If mytablex.EOF Then Exit Do
                    sw = 1
                    tdeliver.Data2.Recordset.AddNew

                    For I = 0 To mytablex.Fields.count - 1
                        tdeliver.Data2.Recordset.Fields(I) = mytablex.Fields(I)
                    Next I

                    tdeliver.Data2.Recordset.Fields("usuario") = tdeliver.cajero
                    tdeliver.Data2.Recordset.Fields("caja") = tdeliver.caja
                    tdeliver.Data2.Recordset.Fields("turno") = tdeliver.turno
                    tdeliver.Data2.Recordset.Fields("local") = "" & mytable11.Fields("local")
                    tdeliver.Data2.Recordset.Update
                    mytablex.MoveNext
                Loop

            End If

            mytablex.Close

        End If

        txmensual.MoveNext
    Loop

    If sw = 1 Then
        tdeliver.codigo = codigo
        tdeliver.nombre = nombre
        MsgBox "Proceso Realizado con exito", 48, "Aviso"
        dlo22_Click
        Exit Sub

    End If

    MsgBox "No se cargado nigun documento", 48, "Aviso"
    Exit Sub

End Sub

Function ir_inicio()

    On Error GoTo cmd5612_err

    txmensual.MoveFirst
    ir_inicio = 1
    Exit Function
cmd5612_err:
    Exit Function

End Function

Private Sub dbgrid1_DblClick()

    On Error GoTo cmd21_err

    If "" & txmensual.Fields("x") = "S" Then
        txmensual.Fields("x") = ""
        txmensual.Update
        Exit Sub

    End If

    If "" & txmensual.Fields("x") = Null Or "" & txmensual.Fields("x") = "" Then
        txmensual.Fields("x") = "S"
        txmensual.Update
        Exit Sub

    End If

    Exit Sub
cmd21_err:
    MsgBox "Seleccione un registro", 48, "Aviso"
    Exit Sub

End Sub

Private Sub DBGrid11_Click()

End Sub

Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cadena.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        codigo = dbgrid3.columns(1)
        nombre = dbgrid3.columns(0)
        Frame3.Visible = False
        codigo.SetFocus
        codigo_KeyPress 13

    End If

End Sub

Private Sub dlo22_Click()
    facmesa.Hide
    Unload facmesa

End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    codigo.SetFocus

End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fechaf.SetFocus

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

    Frame3.Top = 0
    Frame3.Left = 0

End Sub

Sub ejecuta(sw As Integer)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If opcion1 = "1" Then
        If Len(cadena) = 0 Then
            buf = "select Nombre,Codigo from clientes "
        Else
            buf = "select Nombre,Codigo from  clientes where " & xbuffer & " like '%" & cadena & "%'"

        End If

    End If

    If opcion1 = "2" Then
        If Len(cadena) = 0 Then
            buf = "select Local,Tipo,Serie,Numero,Moneda as M,Total,Fecha,Usuario as Cajero,Caja,Turno,Vendedor,Bodega from  factura WHERE codigo='" & codigo & "'"
            buf = buf & "  and fecha>='" & Format(fechai, "YYYYMMDD") & "'"
            buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
            buf = buf & " and estado='2'"
        Else
            buf = "select Local,Tipo,Serie,Numero,Moneda as M,Total,Fecha,Usuario as Cajero,Caja,Turno,Vendedor,Bodega from  factura WHERE codigo='" & codigo & "'"
            buf = buf & "  and fecha>='" & Format(fechai, "YYYYMMDD") & "'"
            buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
            buf = buf & " and estado='2'"
            buf = buf & " and  " & xbuffer & " like '%" & cadena & "%'"

        End If

    End If

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablex

    If opcion1 = "1" Then
        dbgrid3.columns(0).Width = 4000
        dbgrid3.columns(1).Width = 2000

    End If

    If opcion1 = "2" Then
        dbgrid3.columns(0).Width = 500
        dbgrid3.columns(1).Width = 500
        dbgrid3.columns(2).Width = 500
        dbgrid3.columns(3).Width = 1100
        dbgrid3.columns(4).Width = 400
        dbgrid3.columns(5).Width = 1100
        dbgrid3.columns(6).Width = 1500
        dbgrid3.columns(7).Width = 1100
        dbgrid3.columns(8).Width = 700
        dbgrid3.columns(9).Width = 700
        dbgrid3.columns(9).Width = 900
        dbgrid3.columns(9).Width = 700

    End If

    If sw = 1 Then
        dbgrid3.SetFocus

    End If

End Sub

Sub consulta_cuentac()

    Dim buf As String

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        fechaf.SetFocus
        Exit Sub

    End If

    If txmensual.State = 1 Then
        txmensual.Close
        Set txmensual = Nothing

    End If

    'cn.Execute ("update cuentac set estado='0' where estado is null")
    buf = "select * from  cuentac  WHERE codigo='" & codigo & "'"
    buf = buf & "  and fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and estado='0'"
    buf = buf & " and saldo>0"
    txmensual.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txmensual
    suma_totales
    dbGrid1.SetFocus

End Sub

Sub suma_totales()

    Dim sdx As Double

    sdx = 0
    Do

        If txmensual.EOF Then Exit Do
        sdx = sdx + Val("" & txmensual.Fields("saldo"))
        txmensual.MoveNext
    Loop
    total = "" & sdx

End Sub
