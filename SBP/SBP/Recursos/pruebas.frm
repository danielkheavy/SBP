VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form pruebas 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "aDD nEW RECORD"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "uPDATE DBGRID"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "uPDATE"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cREATE"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "pruebas.frx":0000
      Height          =   975
      Left            =   120
      OleObjectBlob   =   "pruebas.frx":0014
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
   End
End
Attribute VB_Name = "pruebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      Dim rs1 As Recordset
      Dim db As Database
      Dim td As TableDef
      Dim fl As Field
      Dim igRow As Integer, igColumn As Integer
      Dim iFields As Integer, iRecords As Integer
      Dim vargBookmark As Variant

      Private Sub Command1_Click()
         ' The Create Database button: By clicking this button, you create
         ' a database with four fields and five records.

         Set db = CreateDatabase("C:\test.mdb", dbLangGeneral)
         Set td = db.CreateTableDef("Table1")

         ' After you create the database, you need to add fields to it.
         For iFields = 1 To 4 ' The last number can be changed to the
                              ' number of fields you want in the database.
            Set fl = td.CreateField("Field " & CStr(iFields), dbText)
            td.Fields.Append fl
         Next iFields

         db.TableDefs.Append td

         ' Now that you have added fields to the database, you need to add
         ' some records through a recordset.
         Set rs1 = db.OpenRecordset("Table1", dbOpenTable)
         For iRecords = 1 To 5  'For each row
            rs1.AddNew          'Add a new record

            For iFields = 1 To 4        ' For each field in the record, add
               rs1("Field " & CStr(iFields)) = CStr(iFields) ' a number.
            Next iFields

         rs1.Update
         Next iRecords

         ' Close both the recordset and database.
         rs1.Close
         db.Close

         ' Populate the DBGrid control with the contents of the Recordset.
         Set db = OpenDatabase("C:\test.mdb")
         Set rs1 = db.OpenRecordset("Select * from Table1")
         Set Data1.Recordset = rs1

         Command1.Visible = False
         Command2.Visible = True
         Command4.Visible = True
      End Sub

      Private Sub Command2_Click()
         ' The Update Database button: By clicking this button, you save
         ' the contents of the text box to the database. Since the contents
         ' of the recordset are being modified, the contents are saved to
         ' the database after you execute the Update method.

         Data1.Recordset.Edit
         Data1.Recordset.Fields(igColumn) = Text1.Text
         Data1.Recordset.Update
      End Sub

      Private Sub Command3_Click()
         ' The Update DBGrid button: By clicking this button, you execute
         ' the UpdateControls method on the Data control to demonstrate
         ' that changing the cell in a bound DBGrid control does not save
         ' the new information to the database. To save these changes, you
         ' must modify the underlying recordset from the Data control.

         Data1.UpdateControls
      End Sub

      Private Sub Command4_Click()
         ' The Add New Record button: By clicking this button, you add new
         ' records to the recordset. Use the following code to add a new
         ' record to the DBGrid control.

         ' Set DBGrid and Data Control Properties to allow new records to
         ' be added.
         DBGrid1.AllowAddNew = True
         Data1.EOFAction = vbAddNew
         Data1.Recordset.MoveLast
         Data1.Recordset.MoveNext
         DBGrid1.Row = DBGrid1.VisibleRows - 1
         Data1.Recordset.AddNew
         For iFields = 1 To 4    ' For each field in the record,
                                 ' add the contents of the text box.
            Data1.Recordset("Field " & CStr(iFields)) = Text1.Text
         Next iFields
         Data1.Recordset.Update

      End Sub

      Private Sub DBGrid1_Change()
         Command3.Visible = True
      End Sub

      Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, _
                                  X As Single, Y As Single)
         Command2.Visible = True
         igColumn = DBGrid1.ColContaining(X)
         igRow = DBGrid1.RowContaining(Y)
         vargBookmark = DBGrid1.RowBookmark(igRow)

         Text1.Text = DBGrid1.Columns(igColumn).CellValue(vargBookmark)

      End Sub

      Private Sub Form_Load()
         Command1.Visible = False
         Command2.Visible = False
         Command3.Visible = False
         Command4.Visible = False
         Command1.Caption = "Create Database"
         Command2.Caption = "Update Database"
         Command3.Caption = "Update DBGrid"
         Command4.Caption = "Add New Record"

         ' If the database does not exist, show the Create Database button.
         If Dir("C:\test.mdb") = "" Then
            Command1.Visible = True
         Else
            ' Open an existing database.
            Set db = OpenDatabase("C:\test.mdb")
            Set rs1 = db.OpenRecordset("Select * from Table1")
            Set Data1.Recordset = rs1
            Command4.Visible = True
         End If

      End Sub

      Private Sub Text1_Change()
         Command2.Visible = True
      End Sub

    

