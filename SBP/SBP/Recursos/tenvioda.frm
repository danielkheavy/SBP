VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tenvioda 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de Actualizaciones Remotas"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Procesar"
      Height          =   1215
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   4335
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Descripcio"
         Caption         =   "Descripcio"
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
         DataField       =   "Ip"
         Caption         =   "Ip"
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
            ColumnWidth     =   3690.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2099.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label procesos 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10800
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Menu lfo9933 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tenvioda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mips As New ADODB.Recordset

Dim cnn  As New ADODB.Connection

Option Explicit

Private Sub Command1_Click()

    Dim found As Integer

    If InputBox("LLave de Paso", "", "") <> "CUIDADO" Then Exit Sub
    Do

        If mips.EOF Then Exit Do
        If Len(Trim("" & mips.Fields("ip"))) > 0 Then
            found = conecta_productos(Trim("" & mips.Fields("ip")))

            If found = 0 Then
                MsgBox "No existe conexion ", 48, "Aviso"

            End If

            If found = 1 Then
                found = envio_data(Trim("" & mips.Fields("ip")))

            End If

        End If

        mips.MoveNext
    Loop

End Sub

Function envio_data(buf As String)

    Dim found As Integer

    Dim vr

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    procesos = "Familia"
    vr = DoEvents
   
    mytablex.Open "select * from familia where familia like '%'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If
   
    cnn.Execute ("delete from familia")
    mytabley.Open "select * from familia where familia like '%'", cnn, adOpenStatic, adLockOptimistic
   
    Do

        If mytablex.EOF Then Exit Do
   
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
   
    'productos
    procesos = "Producto"
    vr = DoEvents
   
    mytablex.Open "select * from producto where producto like '%'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    cnn.Execute ("delete from producto where producto like '%'")
    mytabley.Open "select * from producto where producto like '%'", cnn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
   
    'subfamilia
    procesos = "Subfamilia"
    vr = DoEvents
   
    mytablex.Open "select * from subfamil where subfamilia like '%'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    cnn.Execute ("delete from subfamil")
    mytabley.Open "select * from subfamil where subfamilia like '%'", cnn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
   
    'marca
    procesos = "Marca"
    vr = DoEvents
   
    mytablex.Open "select * from marca where marca like '%'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    cnn.Execute ("delete from marca")
    mytabley.Open "select * from marca where marca like '%'", cnn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    cnn.Close
    MsgBox "Proceso Terminado ", 48, "Aviso"

End Function

Private Sub Form_Activate()

    If mips.RecordCount = 0 Then
        Command1.Enabled = False

    End If

End Sub

Private Sub Form_Load()

    If mips.State = 1 Then mips.Close
    mips.Open "select * from ip ", cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mips

    List1.Clear
    List1.AddItem "Productos"
    List1.AddItem "Ventas"
    List1.ListIndex = 0
  
End Sub

Function conecta_productos(buf As String)

    On Error GoTo cmd8912_err

    cnn.CursorLocation = adUseClient
    cnn.Open "Driver={SQL Server};Server=" & Trim(buf) & ";Database=calipso;uid=sa"
    conecta_productos = 1
    MsgBox "Conexion establecida"
    Exit Function
cmd8912_err:
    MsgBox "No se conecta  " + error$, 48, "Aviso"
    Exit Function

End Function

Private Sub lfo9933_Click()
    tenvioda.Hide
    Unload tenvioda

End Sub
