VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form mmenua 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8295
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox local1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8235
      TabIndex        =   6
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton Cancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin ChamaleonButton.ChameleonBtn Add 
         Height          =   570
         Left            =   5400
         TabIndex        =   10
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1005
         BTYPE           =   5
         TX              =   "Procesar"
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
         MICON           =   "mmenua.frx":0000
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
   Begin VB.TextBox numero 
      Height          =   495
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox serie 
      Height          =   495
      Left            =   1890
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label local11 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Menu dfk992 
      Caption         =   "&Procesar"
   End
   Begin VB.Menu fl3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mmenua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()

End Sub

Private Sub Add_Click()

    '''27/07/2017 kenyo Testing Completo al Sistema
    If Me.Caption = "COPIA" Then

        Dim found3 As Integer

        found3 = valida()

        If found3 = 0 Then
            MsgBox "No existe Documento", 48, "Aviso"
            serie.SetFocus
            Exit Sub

        End If

        rrlocal11 = extra_loquesea(local1)
        rrtipo = extra_loquesea(tipo)
        rrserie = serie
        rrnumero = Numero
        mmenua.Hide
        Unload mmenua

    End If

    '''27/07/2017 kenyo Testing Completo al Sistema

    If Me.Caption = "CARGAPEDIDO" Then
        opcion1 = "1500" ' CARGA DETALLE DOCUMENTO ANTERIOR

        Dim found As Integer

        found = valida()

        If found = 0 Then
            MsgBox "No existe Documento", 48, "Aviso"
            serie.SetFocus
            Exit Sub

        End If

        If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
        found = tptovta.proceso_carga_doc_ant(extra_loquesea(local1), extra_loquesea(tipo), serie, Numero)
       
        If found = 0 Then
            MsgBox "Error de carga", 48, "Aviso"
            Exit Sub

        End If

        tptovta.Frame1.Visible = False
        tptovta.Frame1.Enabled = False
      
        mmenua.Hide
        Unload mmenua

    End If

    '    Dim found2 As Integer
    '    found2 = valida()
    '    If found2 = 0 Then
    '       MsgBox "No existe Documento", 48, "Aviso"
    '       serie.SetFocus
    '       Exit Sub
    '    End If
    '    rrlocal11 = extra_loquesea(local1)
    '    rrtipo = extra_loquesea(tipo)
    '    rrserie = serie
    '    rrnumero = numero
    '    mmenua.Hide
    '    Unload mmenua
    'End If

    'NOTA CREDITO
    If Me.Caption = "NOTACREDITO" Then

        opcion1 = "1500" ' CARGA DETALLE DOCUMENTO ANTERIOR

        Dim foundNC   As Integer

        Dim foundNCFP As Integer
    
        foundNC = valida()

        If foundNC = 0 Then
            MsgBox "No existe Documento", 48, "Aviso"
            serie.SetFocus
            Exit Sub

        End If

        foundNC = tptovta.proceso_carga_doc_ant(extra_loquesea(local1), extra_loquesea(tipo), serie, Numero)
        ' foundNCFP = tptovta.proceso_carga_doc_antFormaPago(extra_loquesea(local1), extra_loquesea(tipo), serie, numero)
    
        If foundNC = 0 Then
            MsgBox "Error de carga", 48, "Aviso"
            Exit Sub

        End If

        tptovta.Frame1.Visible = False
        tptovta.Frame1.Enabled = False

        mmenua.Hide
        Unload mmenua

    End If

    '16/06/2017 kenyo CORRECCION ANULA VENTA OTRA FECHA
    'ANULA OTRA FECHA
    If Me.Caption = "ANULA" Then

        Dim foundA As Integer

        foundA = valida()

        If foundA = 0 Then
            MsgBox "No existe Documento", 48, "Aviso"
            serie.SetFocus
            Exit Sub

        End If

        rrlocal11 = extra_loquesea(local1)
        rrtipo = extra_loquesea(tipo)
        rrserie = serie
        rrnumero = Numero
        mmenua.Hide
        Unload mmenua

    End If

    '16/06/2017 kenyo CORRECCION ANULA VENTA OTRA FECHA

End Sub

Function valida()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT local,tipo,serie,numero FROM factura where local='" & extra_loquesea(local1) & "' and tipo='" & extra_loquesea(tipo) & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida = 1

    End If

    mytablex.Close

End Function

Private Sub Cancel_Click()
    fl3434_Click

End Sub

Private Sub cmdCommand1_Click()

    opcion1 = "1500" ' CARGA DETALLE DOCUMENTO ANTERIOR

    Dim found As Integer

    found = valida()

    If found = 0 Then
        MsgBox "No existe Documento", 48, "Aviso"
        serie.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
    found = tptovta.proceso_carga_doc_ant(extra_loquesea(local1), extra_loquesea(tipo), serie, Numero)

    If found = 0 Then
        MsgBox "Error de carga", 48, "Aviso"
        Exit Sub

    End If

    tptovta.Frame1.Visible = False
    tptovta.Frame1.Enabled = False
      
    mmenua.Hide
    Unload mmenua

End Sub

Private Sub fl3434_Click()
    mmenua.Hide
    Unload mmenua

End Sub

Sub carga_inicial()

    Dim mytablex As New ADODB.Recordset

    local1.Clear
    mytablex.Open "SELECT *  FROM tlocal ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    tipo.Clear
    mytablex.Open "SELECT * FROM tipo WHERE ESTADOT='2'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do

        '''27/07/2017 kenyo Testing Completo al Sistema
        'If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "D" Then
        If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "E" Then

            '''27/07/2017 kenyo Testing Completo al Sistema
   
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0

End Sub

Private Sub Form_Activate()
   
    tptovta.cmdCobrar.Enabled = True
    tptovta.cmdCobrar.Caption = "EFECTIVO"
        
    tptovta.cmd50.Enabled = True
    tptovta.cmd50.Caption = "50"
        
    tptovta.cmd100.Enabled = True
    tptovta.cmd100.Caption = "100"
      
    tptovta.cmd200.Enabled = True
    tptovta.cmd200.Caption = "200"
        
End Sub

Private Sub Form_Load()
    carga_inicial
        
End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Numero.SetFocus

End Sub
