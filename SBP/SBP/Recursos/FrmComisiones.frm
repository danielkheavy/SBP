VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmComisiones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "COMISIONES POR PRODUCTO POR TRABAJADOR"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7590
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8415
      Left            =   0
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   7330
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   2775
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   11880
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
      Begin ChamaleonButton.ChameleonBtn Command1 
         Height          =   495
         Left            =   4680
         TabIndex        =   14
         Top             =   480
         Width           =   1125
         _ExtentX        =   1984
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
         MICON           =   "FrmComisiones.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChaAceptar 
         Height          =   930
         Left            =   6000
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "FrmComisiones.frx":001C
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
         Left            =   6000
         TabIndex        =   16
         Top             =   2280
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "FrmComisiones.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   900
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmComisiones.frx":0054
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Borrar registro"
      Top             =   120
      Width           =   765
   End
   Begin MSDataGridLib.DataGrid DgvComisiones 
      Height          =   3495
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16744576
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
   Begin VB.CommandButton cmdSave 
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
      Height          =   645
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmComisiones.frx":1266
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Grabar registro"
      Top             =   120
      Width           =   765
   End
   Begin VB.CommandButton cmdBuscarVendedor 
      Height          =   420
      Left            =   6750
      Picture         =   "FrmComisiones.frx":2478
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Buscar producto"
      Top             =   1680
      Width           =   435
   End
   Begin VB.CommandButton Command2 
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
      Height          =   645
      Left            =   1680
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmComisiones.frx":2C26
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   765
   End
   Begin VB.TextBox comision 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   1680
      TabIndex        =   18
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label codigo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label descripcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod. Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label producto 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label nombre 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   5055
   End
End
Attribute VB_Name = "FrmComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dbpersonal           As New ADODB.Recordset

Dim dbConsultaComisiones As New ADODB.Recordset

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame1.Visible = False
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub ChaAceptar_Click()
    dbgrid1_KeyDown 13, 0

End Sub

Private Sub ChaCERRAR_Click()
    Frame1.Visible = False
    Frame1.Enabled = False
    cmdBuscarVendedor.SetFocus
    Exit Sub

End Sub

Private Sub cmdBuscarVendedor_Click()

    Dim found As Integer

    Frame1.Visible = True
    Frame1.Enabled = True
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    buffer = ""
    
    Frame1.Visible = True
    nombre.Caption = ""
    codigo.Caption = ""
    found = busca_vendedor

End Sub

Function ExistenciaVendedor(ByRef codVendedor)

    Dim rsexiste As New ADODB.Recordset

    rsexiste.Open "SELECT * FROM vendedorcomision where producto='" & Trim(producto) & "'  and codigo='" & Trim(codVendedor) & "' ", cn, adOpenKeyset, adLockOptimistic

    If rsexiste.RecordCount > 0 Then  'si existe
        MsgBox "Ya existe vendedor ", 48, "Aviso"
        ExistenciaVendedor = 1
        Exit Function

    End If

End Function

Sub LimpiaCampos()
    codigo = ""
    nombre = ""
    comision = ""

End Sub

Sub GuardaComisiones()
    dbConsultaComisiones.AddNew
    dbConsultaComisiones.Fields("producto") = producto
    dbConsultaComisiones.Fields("descripcion") = descripcion
    dbConsultaComisiones.Fields("codigo") = codigo
    dbConsultaComisiones.Fields("nombre") = nombre
    dbConsultaComisiones.Fields("comision") = Val(comision)
    dbConsultaComisiones.Update
    LimpiaCampos

End Sub

Sub EliminaComisiones()

    Dim buf  As String

    Dim buf2 As String

    On Error GoTo cmd656_err
    
    If MsgBox("Desea Borra a " + dbConsultaComisiones.Fields("nombre"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    dbConsultaComisiones.Delete
    dbConsultaComisiones.Update
    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub cmdDelete_Click()
    EliminaComisiones

End Sub

Private Sub cmdSave_Click()

    If codigo = "" Then
        MsgBox "Seleccione vendedor ", 48, "Aviso"
        cmdBuscarVendedor.SetFocus
        Exit Sub

    End If

    If comision = "" Then
        MsgBox "Agregue comisión ", 48, "Aviso"
        comision.SetFocus
        Exit Sub

    End If

    GuardaComisiones

End Sub

Private Sub Command1_Click()
    Call busca_vendedor

End Sub

Function busca_vendedor()

    Dim buf1 As String
  
    If Len(buffer) = 0 Then
        buf1 = "select Nombre,Codigo from vendedor"
           
    Else
        buf1 = "select Nombre,Codigo from vendedor where " & Combo1 & " like '%" & buffer & "%'"

    End If
              
    Set dbpersonal = Nothing

    If dbpersonal.State = 1 Then
        dbpersonal.Close
        Set dbpersonal = Nothing

    End If

    dbpersonal.Open buf1, cn, adOpenStatic, adLockOptimistic
    Set DBGrid1.DataSource = dbpersonal
    DBGrid1.refresh

    If dbpersonal.RecordCount = 0 Then
        buffer.SetFocus
        Exit Function

    End If
      
    DBGrid1.columns(0).Width = 3200
    DBGrid1.columns(1).Width = 1600
    busca_vendedor = 1
    Exit Function
cmd8912_err:
    MsgBox "Aviso en busca_vendedor " & error$, 48, "Aviso"
    buffer = ""

End Function

Function Lista_VendedorComision()

    If dbConsultaComisiones.State = 1 Then dbConsultaComisiones.Close
    dbConsultaComisiones.Open "select *from vendedorcomision where producto = '" & producto & "'", cn, adOpenStatic, adLockOptimistic
   
    Set DgvComisiones.DataSource = dbConsultaComisiones

    DgvComisiones.columns(0).Width = 0
    DgvComisiones.columns(1).Width = 0
    DgvComisiones.columns(2).Width = 1600
    DgvComisiones.columns(3).Width = 3100
    DgvComisiones.columns(4).Width = 1600

End Function

Private Sub Command2_Click()
    FrmComisiones.Hide
    Unload FrmComisiones

End Sub

Private Sub Form_Activate()
    Frame1.Top = 0: Frame1.Left = 0
    Lista_VendedorComision

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found    As Integer

    Dim xbuf     As String

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If KeyCode = 27 Then
        Frame1.Visible = False
        Exit Sub

    End If

    If KeyCode = 13 Then
    
        If dbpersonal.RecordCount = 0 Then
            Exit Sub

        End If
    
        If ExistenciaVendedor("" & dbpersonal.Fields("codigo")) = 1 Then Exit Sub

        codigo = "" & dbpersonal.Fields("codigo")
        nombre = "" & dbpersonal.Fields("nombre")
    
        Frame1.Visible = False
        Frame1.Enabled = False
        comision.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Form_Load()
    Lista_VendedorComision

End Sub
