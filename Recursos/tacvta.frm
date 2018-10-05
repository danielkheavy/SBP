VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tacvta 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Empresa Datos Ventas"
   ClientHeight    =   7290
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Documentos Seleccionados"
      Height          =   7095
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "X"
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Procesar"
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "tacvta.frx":0000
         Height          =   5775
         Left            =   240
         OleObjectBlob   =   "tacvta.frx":0014
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   11175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total registros"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   6240
         Width           =   2055
      End
      Begin VB.Label registros 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   6240
         Width           =   2055
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox caja 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   3255
   End
   Begin VB.ComboBox tipodoc 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Buscar"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label empresa 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EmpresaActual"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde Caja"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDoc."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fechaf"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaI"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu doo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tacvta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo cmd37_err
Dim buf As String
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
buf = "select * from Factura where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
If tipodoc <> "%" Then
   buf = buf & " and tipo='" & extra_loquesea(tipodoc) & "'"
End If
If caja <> "%" Then
   buf = buf & " and caja='" & extra_loquesea(caja) & "'"
End If
buf = buf & " order by fecha"
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = "\kali\rp_orion.v2\001d\06"
               Data1.RecordSource = buf
               Data1.Refresh
Frame1.Visible = True
DBGrid1.SetFocus
Exit Sub
cmd37_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub Command2_Click()
Dim buf As String
Dim vr
Dim mytablex As Table
Dim mytablez As Table
Dim mytabley As Table
Dim sdx As Double
doo33.Enabled = False
Command2.Enabled = False
buf = "DELETE FROM factura WHERE tipo='" & extra_loquesea(tipodoc) & "'"
buf = buf & "  and fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
mydbxglo.Execute buf
buf = "DELETE FROM DETALLE WHERE tipo='" & extra_loquesea(tipodoc) & "'"
buf = buf & "  and fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
mydbxglo.Execute buf
buf = "DELETE FROM fpagov WHERE tipo='" & extra_loquesea(tipodoc) & "'"
buf = buf & "  and fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
mydbxglo.Execute buf
Data1.Refresh
Set mytablex = mydbxglo.OpenTable("factura")

Set mytabley = mydbxglo.OpenTable("detalle")
mytabley.Index = "tdetalle"

Set mydbz = OpenDatabase("\kali\rp_orion.v2\001d\06", False, False, "foxpro 2.5;")
Set mytablez = mydbz.OpenTable("detalle")
mytablez.Index = "tdetalle"
Do
If Data1.Recordset.EOF Then Exit Do
       vr = DoEvents()
       mytablex.AddNew
       For i = 0 To Data1.Recordset.Fields.count - 1
           mytablex.Fields(i) = Data1.Recordset.Fields(i)
       Next i
       mytablex.Fields("bodega") = "01"
       mytablex.Update
amki:
       mytabley.Seek "=", "" & Data1.Recordset.Fields("local"), "" & Data1.Recordset.Fields("tipo"), "" & Data1.Recordset.Fields("serie"), "" & Data1.Recordset.Fields("numero")
       If Not mytabley.NoMatch Then
          mytabley.Delete
          GoTo amki
       End If
        sdx = 0
        mytablez.Seek "=", "" & Data1.Recordset.Fields("local"), "" & Data1.Recordset.Fields("tipo"), "" & Data1.Recordset.Fields("serie"), "" & Data1.Recordset.Fields("numero")
         If Not mytablez.NoMatch Then
          Do
          If mytablez.EOF Then GoTo P1
          If "" & Data1.Recordset.Fields("local") = "" & mytablez.Fields("local") And "" & Data1.Recordset.Fields("tipo") = "" & mytablez.Fields("tipo") And "" & Data1.Recordset.Fields("serie") = "" & mytablez.Fields("serie") And "" & Data1.Recordset.Fields("numero") = "" & mytablez.Fields("numero") Then
             mytabley.AddNew
             For i = 0 To mytablez.Fields.count - 1
             mytabley.Fields(i) = mytablez.Fields(i)
             Next i
             mytabley.Fields("bodega") = "01"
             mytabley.Update
             sdx = sdx + 1
             Else: GoTo P1
          End If
          mytablez.MoveNext
          Loop
P1:
         'MsgBox sdx
         End If
       Data1.Recordset.MoveNext
Loop
mytabley.Close
mytablex.Close
mytablez.Close
doo33.Enabled = True
Command2.Enabled = True
MsgBox "Proceso Terminado ", 48, "Aviso"

End Sub

Private Sub Command3_Click()
doo33_Click
End Sub

Private Sub doo33_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
tacvta.Hide
Unload tacvta
End Sub

Private Sub Form_Load()
Dim mydbx As Database
empresa = menup.gempresa
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
Set mydbx = OpenDatabase("\kali\rp_orion.v2\001d\06", False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("tipo")
tipodoc.Clear
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Then
   tipodoc.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
End If
mytablex.MoveNext
Loop
tipodoc.ListIndex = 0
mytablex.Close

Set mytablex = mydbx.OpenTable("parameca")
caja.Clear
Do
If mytablex.EOF Then Exit Do
If empresa = "003D" Then
   If "" & mytablex.Fields("caja") = "01" Or "" & mytablex.Fields("caja") = "03" Or "" & mytablex.Fields("caja") = "05" Then
     If "" & mytablex.Fields("terminal") = "C" Then
        caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")
     End If
   End If
End If
If empresa = "004D" Then
   If "" & mytablex.Fields("caja") = "06" Then
        If "" & mytablex.Fields("terminal") = "C" Then
           caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")
        End If
   End If
End If
mytablex.MoveNext
Loop
caja.ListIndex = 0
mytablex.Close
mydbx.Close
End Sub
