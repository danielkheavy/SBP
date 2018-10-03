VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tincxc 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Cuentas Corrientes"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Consultas"
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
      Left            =   0
      TabIndex        =   50
      Top             =   600
      Visible         =   0   'False
      Width           =   9975
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   5400
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   0
         TabIndex        =   54
         Top             =   840
         Width           =   9735
         _ExtentX        =   17171
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
   End
   Begin VB.TextBox grupo 
      Height          =   375
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   47
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox turno 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   44
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox caja 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   42
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox usuario 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   40
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox local1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   38
      Text            =   "01"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox saldo 
      Height          =   375
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox abono 
      Height          =   375
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   14
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox interes 
      Height          =   375
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox total 
      Height          =   375
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Zona 
      Height          =   375
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   10
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox vendedor 
      Height          =   375
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox fechav 
      Height          =   375
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox fecha 
      Height          =   375
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox moneda 
      Height          =   375
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox nombre 
      Height          =   375
      Left            =   1920
      MaxLength       =   60
      TabIndex        =   6
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox tipoclie 
      Height          =   375
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox cuota 
      Height          =   375
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox numero 
      Height          =   375
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox serie 
      Height          =   375
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Tipo 
      Height          =   375
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11070
      TabIndex        =   16
      Top             =   0
      Width           =   11130
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
         Height          =   615
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tincxc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Grabar registro"
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tincxc.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label bandera 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label xcuentaco 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   6600
      Width           =   105
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.Credito A .nticipo  D.Deposito  O.Otros"
      Height          =   375
      Left            =   6120
      TabIndex        =   49
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grupo"
      Height          =   375
      Left            =   6120
      TabIndex        =   48
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label anticipo 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6240
      TabIndex        =   46
      Top             =   4920
      Width           =   105
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
      Height          =   375
      Left            =   240
      TabIndex        =   45
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      Height          =   375
      Left            =   240
      TabIndex        =   43
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label nombrev 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   37
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6240
      TabIndex        =   36
      Top             =   4560
      Width           =   105
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   6120
      TabIndex        =   33
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abono"
      Height          =   375
      Left            =   6120
      TabIndex        =   32
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interes"
      Height          =   375
      Left            =   6120
      TabIndex        =   31
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   6120
      TabIndex        =   30
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zona"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor"
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Vencim."
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha "
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      Height          =   375
      Left            =   6120
      TabIndex        =   25
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoClie (CPV)"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuota"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Menu jui12 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu dlo8912 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tincxc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim xcuentaco As String
Dim xnameclie As String

Private Sub abono_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    saldo.SetFocus

End Sub

Private Sub abono_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        interes.SetFocus
        Exit Sub

    End If

End Sub

Private Sub cmdExit_Click()
    dlo8912_Click

End Sub

Private Sub cmdSave_Click()

    Dim found As Integer

    found = graba_datos()

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(codigo) = 0 Then
        If tipoclie <> "C" And tipoclie <> "V" And tipoclie <> "P" Then
            tipoclie.SetFocus
            Exit Sub

        End If

        consulta_codigo
        Exit Sub

    End If

    nombre.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If tipoclie.Enabled = False Then
            cuota.SetFocus
            Exit Sub

        End If

        tipoclie.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        If tipoclie <> "C" And tipoclie <> "V" And tipoclie <> "P" Then
            If tipoclie.Enabled = True Then
                tipoclie.SetFocus

            End If

            Exit Sub

        End If

        consulta_codigo

    End If

End Sub

Private Sub Command1_Click()

    Dim buf       As String

    Dim xbuf      As String

    Dim rconsulta As New ADODB.Recordset

    If tipoclie = "C" Then
        xnameclie = "clientes"

    End If

    If tipoclie = "P" Then
        xnameclie = "proveedo"

    End If

    If tipoclie = "V" Then
        xnameclie = "Vendedor"

    End If

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Tipo,Serie,Anticipo from tipo where grupo='" & acu & "'"
        Else
            buf = "select Descripcio,Tipo,Serie,Anticipo from Tipo where grupo='" & acu & "' and " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "2" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from " & xnameclie
        Else
            buf = "select Nombre,Codigo  from " & xnameclie & " where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "3" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from vendedor "
        Else
            buf = "select Nombre,Codigo  from vendedor where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "4" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Zona from Zona "
        Else
            buf = "select Descripcio,Zona  from Zona where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    Set dbGrid1.DataSource = rconsulta
               
    If opcion1 = "1" Or opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Then
        dbGrid1.columns(0).Width = 3500
        dbGrid1.columns(1).Width = 900

    End If

    dbGrid1.SetFocus

End Sub

Private Sub cuota_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Val(cuota) = 0 Then
        cuota = "1"

    End If

    If tipoclie.Enabled = True Then
        tipoclie.SetFocus
        Exit Sub

    End If

    codigo.SetFocus

End Sub

Private Sub cuota_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If Numero.Enabled = False Then Exit Sub
        Numero.SetFocus
        Exit Sub

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            tipo = Trim(dbGrid1.columns(1))
            serie = Trim(dbGrid1.columns(2))
            anticipo = Trim(dbGrid1.columns(3))
            Frame1.Visible = False
            Frame1.Enabled = False
            tipo.SetFocus
            tipo_KeyPress 13

        End If

        If opcion1 = "2" Then
            codigo = Trim(dbGrid1.columns(1))
            nombre = Trim(dbGrid1.columns(0))
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            codigo_KeyPress 13

        End If

        If opcion1 = "3" Then
            vendedor = Trim(dbGrid1.columns(1))
            nombrev = Trim(dbGrid1.columns(0))
            Frame1.Visible = False
            Frame1.Enabled = False
            vendedor.SetFocus
            vendedor_KeyPress 13

        End If

        If opcion1 = "4" Then
            zona = Trim(dbGrid1.columns(1))
            Frame1.Visible = False
            Frame1.Enabled = False
            zona.SetFocus
            zona_KeyPress 13

        End If

    End If

End Sub

Private Sub dlo8912_Click()

    'If Frame2.Visible = True Then
    '   Frame2.Visible = False
    '   codigo.SetFocus
    '   Exit Sub
    'End If
    If Frame1.Visible = True Then
        If opcion1 = "1" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            tipo.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            Exit Sub

        End If

        If opcion1 = "3" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            vendedor.SetFocus
            Exit Sub

        End If

        If opcion1 = "4" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            zona.SetFocus
            Exit Sub

        End If
   
    End If

    tincxc.Hide
    Unload tincxc

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fecha) = 0 Then
        fecha = Format(Now, "dd/mm/yyyy")

    End If

    fechav.SetFocus

End Sub

Private Sub fecha_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nombre.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechav_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechav) = 0 Then
        fechav = Format(Now, "dd/mm/yyyy")

    End If

    vendedor.SetFocus

End Sub

Private Sub fechav_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fecha.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Form_Activate()

    If acu = "V" Then
        xnameclie = "clientes"

        'xcuentaco = "cuentac"
    End If

    If acu = "1" Then
        xnameclie = "clientes"

        'xcuentaco = "LETRACC"
    End If

    If acu = "2" Then
        xnameclie = "proveedo"

        'xcuentaco = "LETRApp"
    End If

    If acu = "C" Then
        xnameclie = "proveedo"

        'xcuentaco = "cuentap"
        'tipoclie = "P"
        'tipoclie.Enabled = False
    End If

End Sub

Private Sub interes_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    sumar
    abono.SetFocus

End Sub

Private Sub interes_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        total.SetFocus
        Exit Sub

    End If

End Sub

Private Sub jui12_Click()
    cmdSave_Click

End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    total.SetFocus

End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        zona.SetFocus
        Exit Sub

    End If

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fecha.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If codigo.Enabled = True Then
            codigo.SetFocus

        End If

        Exit Sub

    End If

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(Numero) = 0 Then Exit Sub
    found = valida_nuevo()

    If found = 1 Then
        If bandera = "NUEVO" Then
            MsgBox "Documento ya existe ", 48, "Aviso"
            Numero = ""
            Exit Sub

        End If

    End If

    cuota.SetFocus

End Sub

Private Sub numero_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        serie.SetFocus
        Exit Sub

    End If

End Sub

Private Sub saldo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        abono.SetFocus
        Exit Sub

    End If

    Grupo.SetFocus

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Numero.SetFocus

End Sub

Private Sub serie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        tipo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(tipo) = 0 Then
        consulta_tipo
        Exit Sub

    End If

    serie.SetFocus

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipo

    End If

End Sub

Private Sub tipoclie_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    codigo.SetFocus

End Sub

Private Sub tipoclie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If cuota.Enabled = True Then
            cuota.SetFocus

        End If

        Exit Sub

    End If

End Sub

Private Sub total_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Val(total) <= 0 Then
        total = ""
        total.SetFocus
        Exit Sub

    End If

    sumar
    interes.SetFocus

End Sub

Private Sub total_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        moneda.SetFocus
        Exit Sub

    End If

End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    zona.SetFocus

End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechav.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_vendedor

    End If

End Sub

Private Sub zona_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    moneda.SetFocus

End Sub

Private Sub zona_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        vendedor.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_zona

    End If

End Sub

Function graba_datos()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    mytablex.Open "select * from " & xcuentaco & " where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "' and cuota='" & cuota & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_registro mytablex
        mytablex.Update
    Else
        mytablex.AddNew
        pone_registro mytablex
        mytablex.Update

    End If

    '------------------------------------- ------------
    mytablex.Close
 
    dlo8912_Click

End Function

Sub pone_registro(mytablex As ADODB.Recordset)
    mytablex.Fields("usuario") = usuario
    mytablex.Fields("caja") = caja
    mytablex.Fields("anticipo") = anticipo
    mytablex.Fields("turno") = turno
    mytablex.Fields("grupo") = Grupo
    mytablex.Fields("local") = local1
    mytablex.Fields("tipo") = tipo
    mytablex.Fields("serie") = serie
    mytablex.Fields("numero") = Numero
    mytablex.Fields("cuota") = cuota
    mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
    mytablex.Fields("fechav") = Format(fechav, "dd/mm/yyyy")
    mytablex.Fields("tipoclie") = tipoclie
    mytablex.Fields("codigo") = codigo
    mytablex.Fields("nombre") = nombre
    mytablex.Fields("zona") = zona
    mytablex.Fields("vendedor") = vendedor
    mytablex.Fields("moneda") = moneda
    mytablex.Fields("total") = Val(total)
    mytablex.Fields("interes") = Val(interes)
    mytablex.Fields("abono") = Val(abono)
    mytablex.Fields("saldo") = Val(saldo)
    mytablex.Fields("estado") = "0"
    mytablex.Fields("fpago") = "C"

End Sub

Function valida()

    Dim found As Integer

    found = busca_tipo()

    If found = 0 Then

        'Tipo = ""
        If tipo.Enabled = True Then
            tipo.SetFocus

        End If

        Exit Function

    End If

    If Len(Numero) = 0 Then
        Numero.SetFocus
        Exit Function

    End If

    found = valida_nuevo()

    If found = 1 Then
        If bandera = "NUEVO" Then
            MsgBox "Documento ya existe ", 48, "Aviso"
            Numero = ""
            Numero.SetFocus
            Exit Function

        End If

    End If

    If Not IsNumeric(cuota) Then
        cuota = ""
        cuota.SetFocus
        Exit Function

    End If

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    found = busca_codigo()

    If found = 0 Then
        codigo = ""

        If codigo.Enabled = True Then
            codigo.SetFocus

        End If

        Exit Function

    End If

    If Len(vendedor) > 0 Then
        found = busca_vendedor()

        If found = 0 Then
            vendedor.SetFocus
            Exit Function

        End If

    End If

    If Len(zona) > 0 Then
        found = busca_zona()

        If found = 0 Then
            zona.SetFocus
            Exit Function

        End If

    End If

    If valida_fecha(fecha) = 0 Then
        fecha = ""
        fecha.SetFocus
        Exit Function

    End If

    If valida_fecha(fechav) = 0 Then
        fechav = ""
        fechav.SetFocus
        Exit Function

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda = ""
        moneda.SetFocus
        Exit Function

    End If

    If Val(total) <= 0 Then
        total = ""
        total.SetFocus
        Exit Function

    End If

    If Grupo <> "C" And Grupo <> "A" And Grupo <> "D" And Grupo <> "O" Then
        Grupo = ""
        Grupo.SetFocus
        Exit Function

    End If

    sdx = Val(total) + Val(interes) - Val(abono)
    saldo = Format(sdx, "0.00")
    valida = 1

End Function

Function valida_nuevo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xcuentaco & " where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "' and cuota='" & cuota & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_nuevo = 1

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Sub consulta_tipo()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Tipo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub consulta_codigo()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command1_Click

End Sub

Sub consulta_vendedor()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "3"
    Command1_Click

End Sub

Sub consulta_zona()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Zona"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "4"
    Command1_Click

End Sub

Function busca_tipo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tipo where tipo='" & tipo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        anticipo = "" & mytablex.Fields("anticipo")
        busca_tipo = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_codigo()

    Dim mytablex As New ADODB.Recordset

    If tipoclie = "C" Then
        xnameclie = "clientes"

    End If

    If tipoclie = "P" Then
        xnameclie = "proveedo"

    End If

    If tipoclie = "V" Then
        xnameclie = "Vendedor"

    End If

    mytablex.Open "select * from " & xnameclie & " where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_codigo = 1

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Function busca_vendedor()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & vendedor & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_vendedor = 1

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Function busca_zona()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from zona where zona='" & zona & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_zona = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub sumar()

    Dim sdx As Double

    sdx = Val(total) + Val(interes) - Val(abono)
    saldo = Format(sdx, "0.00")

End Sub

