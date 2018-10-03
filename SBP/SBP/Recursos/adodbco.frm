VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form adodbco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejemplos"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8493
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   15
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
      BeginProperty Column01 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5745.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "adodbco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset

Private Sub dbgrid1_AfterColEdit(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 1
            MsgBox "afterCOLEDIT:" & dbGrid1.columns(1)

            If dbGrid1.columns(1) <> "HOLA" Then
                MsgBox "error "
                Cancel = True
                Exit Sub

            End If
            
    End Select

End Sub

Private Sub dbgrid1_AfterColUpdate(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 1
            MsgBox "aftercolupdate:" & dbGrid1.columns(1)

            If dbGrid1.columns(1) <> "HOLA" Then
                MsgBox "error "
                Cancel = True
                Exit Sub

            End If
            
    End Select

End Sub

Private Sub dbgrid1_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    Select Case ColIndex

        Case 1

            If KeyAscii = 13 Then
                MsgBox "BEFORECOLEDIT:" & dbGrid1.columns(1)

                If dbGrid1.columns(1) <> "HOLA" Then
                    MsgBox "error "
                    Cancel = True
                    Exit Sub

                End If

            End If
            
    End Select

End Sub

Private Sub dbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Select Case ColIndex

        Case 1
            MsgBox "Beforecolupdate:" & dbGrid1.columns(1) & " oldvalue=" & OldValue

            If dbGrid1.columns(1) <> "HOLA" Then
                MsgBox "error "
                Cancel = True
                Exit Sub

            End If

    End Select

End Sub

Private Sub Form_Load()
    Set mytablex = New ADODB.Recordset
 
    mytablex.Open "Select *  From clientes ", cn, adOpenDynamic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex
 
End Sub
