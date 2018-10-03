VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form reporget 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   15645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Borra"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Adiciona"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7920
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   8280
      Width           =   3135
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   8280
      Width           =   2415
   End
   Begin VB.TextBox datos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   8280
      Width           =   2415
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   8280
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   30
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Campo"
         Caption         =   "Campo"
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
         DataField       =   "Tamano"
         Caption         =   "Tamano"
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
      BeginProperty Column03 
         DataField       =   "Condicion"
         Caption         =   "Condicion"
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
         DataField       =   "Dato"
         Caption         =   "Dato"
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
         DataField       =   "Logico"
         Caption         =   "Logico"
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
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2190.047
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reporte"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   5655
      Left            =   10200
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "And/Or"
      Height          =   375
      Left            =   8520
      TabIndex        =   13
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dato"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   120
      Top             =   7800
      Width           =   11535
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Criterio"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refresca por defecto"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Condiciones Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   11535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Reporte"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ArchivoReporte"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Vista"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label ejecutado 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label sentencia 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   9105
   End
   Begin VB.Label archivoreporte 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label NAMETABLA 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "reporget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rwreporte As New ADODB.Recordset



Function crear_view1()
Dim buf As String
Dim sw As Integer
On Error GoTo cmd671_err
buf = " "
sw = 0
If rwreporte.RecordCount > 0 Then
rwreporte.MoveFirst
Do
If rwreporte.EOF Then Exit Do
If "" & rwreporte.Fields("dato") <> "%" Then
   If sw = 0 Then
      buf = " where "
      sw = 1
   End If
   buf = buf & " " & rwreporte.Fields("campo") & " " & rwreporte.Fields("condicion") & " " & rwreporte.Fields("dato")
   If "" & rwreporte.Fields("logico") <> "%" Then
      buf = buf & " " & rwreporte.Fields("logico") & " "
   End If
End If
rwreporte.MoveNext
Loop
End If
'Si fuera con ODBC sería así
'CrystalReport1.Connect = = "Provider=MSDASQL;DSN=NombreX;UID=sa;PWD=CveX"

CrystalReport1.Connect = "Provider=SQLOLEDB;Server=" & menup.vservidor & ";Database=calipso;UID=sa;PWD="
'MsgBox globalpath & "\001d\06\reportes\" & archivoreporte
CrystalReport1.ReportFileName = globalpath & "\001d\06\reportes\" & archivoreporte
'MsgBox "SELECT * from " & NAMETABLA & " " & buf
CrystalReport1.DiscardSavedData = True
CrystalReport1.ProgressDialog = True
CrystalReport1.SQLQuery = "SELECT * from " & NAMETABLA & " " & buf
'cr2.PrinterSelect
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1

crear_view1 = 1
Exit Function
cmd671_err:
MsgBox "Aviso en " + error$, 48, "Aviso"
Exit Function
End Function





Private Sub Command4_Click()

crear_view1
End Sub

Private Sub Command5_Click()
Dim CAMPO1 As String
Dim CAMPO2 As String
Dim campo3 As String
Dim campo4 As String
Dim found As Integer

On Error GoTo cmd44512_err
If Combo4.Text = "%" Then
   MsgBox "Campo No definido ", 48, "Aviso"
   Exit Sub
End If
If Combo5.Text = "%" Then
   MsgBox "Condicion No definido ", 48, "Aviso"
   Exit Sub
End If
If Len(datos) = 0 Then
   MsgBox "Dato No definido ", 48, "Aviso"
   Exit Sub
End If
'If Combo6.Text <> "%" Then Exit Sub
CAMPO1 = ""
CAMPO2 = ""
campo3 = ""
campo4 = ""


found = extraer_camposx("" & Combo4.Text, CAMPO1, CAMPO2, campo3, campo4, "|")
'MsgBox campo1 & " " & campo2 & " " & campo3 & " " & campo4
       rwreporte.AddNew
       rwreporte.Fields("campo") = "" & CAMPO1
       rwreporte.Fields("tamano") = "" & CAMPO2
       rwreporte.Fields("tipo") = "" & campo3
       rwreporte.Fields("condicion") = "" & Combo5.Text
       rwreporte.Fields("dato") = "" & datos
       rwreporte.Fields("logico") = "" & Combo6.Text
       rwreporte.Update
       Combo4.ListIndex = 0
       Combo5.ListIndex = 0
       datos = ""
       Combo6.ListIndex = 0
Exit Sub
cmd44512_err:
MsgBox "No se puede Grabar " + error$, 48, "Aviso"
End Sub

Private Sub Command6_Click()
On Error GoTo cmd8911_err
rwreporte.Delete
Exit Sub
cmd8911_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub


Private Sub flo44_Click()
reporget.Hide
Unload reporget
End Sub

Function extraer_camposx(campo As String, CAMPO1 As String, CAMPO2 As String, campo3 As String, campo4 As String, Flags As String)
Dim i As Integer
Dim j As Integer
Dim temp As String
i = 0
temp = Trim$(campo)
If Len(temp) = 0 Then Exit Function
Do
   j = InStr(temp, Flags)
   If j > 0 Then
      i = i + 1
      'MsgBox Mid$(temp, 1, j - 1)
      Select Case i
             Case 1: CAMPO1 = Mid$(temp, 1, j - 1)
             Case 2: CAMPO2 = Mid$(temp, 1, j - 1)
             Case 3: campo3 = Mid$(temp, 1, j - 1)
             Case 4: campo4 = Mid$(temp, 1, j - 1)
             'Case 5: campo5 = Mid$(temp, 1, J - 1)
      End Select
      temp = Trim$(Mid$(temp, j + 1))
     Else
     Exit Function
   End If
Loop
   Exit Function
End Function

Private Sub Form_Activate()
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd54_err
   Dim cad As String
   If ejecutado <> "S" Then
   cad = "select * from " & NAMETABLA
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   abre_tablax mytablex
   mytablex.Close
   consulta_sql
   End If
   ejecutado = "S"
   Exit Sub
cmd54_err:
   MsgBox "Proceso No realizado ", 48, "Aviso"
   Exit Sub
End Sub
Sub abre_tablax(mytablex As ADODB.Recordset)
   Dim i As Integer
   Dim cad As String
   
   
   Combo4.Clear 'como debe salir el nombre
   Combo5.Clear
   Combo6.Clear
   
   Combo5.AddItem "%"
   Combo5.AddItem "Like"
   Combo5.AddItem "="
   Combo5.AddItem "<>"
   Combo5.AddItem ">"
   Combo5.AddItem "<"
   Combo5.AddItem ">="
   Combo5.AddItem "<="
   
   
   Combo6.AddItem "%"
   Combo6.AddItem "AND"
   Combo6.AddItem "OR"
   
   Combo4.AddItem "%"
   'MsgBox Trim(mytablex.Fields(0).DefinedSize)
   For i = 0 To mytablex.Fields.count - 1
       Combo4.AddItem Trim(mytablex.Fields(i).Name) & "|" & mytablex.Fields(i).DefinedSize & "|" & mytablex.Fields(i).Type & "|"
   Next i
   Combo4.ListIndex = 0
   Combo5.ListIndex = 0
   Combo6.ListIndex = 0
End Sub


Sub consulta_sql()
If rwreporte.State = 1 Then rwreporte.Close
rwreporte.Open "select * from  reporteg", cn, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = rwreporte
DataGrid1.Refresh
End Sub

