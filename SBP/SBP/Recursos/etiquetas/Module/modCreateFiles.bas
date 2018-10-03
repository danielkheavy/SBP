Attribute VB_Name = "modCreateFiles"
Option Explicit

Public Sub CreateDAT()

  Dim TBL  As ADOX.Table
  Dim INDX As ADOX.Index
  Dim Mydb As ADODB.Connection, MySet As ADODB.Recordset
  Dim CAT  As ADOX.Catalog

   On Error GoTo CreateDatERROR

   Set CAT = New ADOX.Catalog

   '/* Engine Type = 4; (Access97)
   '/* Engine Type = 5; (Access2000)

   CAT.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & AppPath & App.Title & ".dat;" & _
      "Jet OLEDB:Database Password=;" & _
      "Jet OLEDB:Engine Type=4;"

   '/* Create Table 'Settings' */
   Set TBL = New ADOX.Table
   Set TBL.ParentCatalog = CAT
   With TBL
      .Name = "Settings"
      .Columns.Append "ID", adInteger, 0
      .Columns("ID").Properties("AutoIncrement") = True
      .Columns("ID").Properties("NullAble") = True

      .Columns.Append "Description", adVarWChar, 50
      .Columns("Description").Properties("NullAble") = True
      .Columns("Description").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "TopMargin", adSingle, 0
      .Columns("TopMargin").Properties("NullAble") = True

      .Columns.Append "SideMargin", adSingle, 0
      .Columns("SideMargin").Properties("NullAble") = True

      .Columns.Append "VPitch", adSingle, 0
      .Columns("VPitch").Properties("NullAble") = True

      .Columns.Append "HPitch", adSingle, 0
      .Columns("HPitch").Properties("NullAble") = True

      .Columns.Append "NoAcross", adSmallInt, 0
      .Columns("NoAcross").Properties("NullAble") = True

      .Columns.Append "NoDown", adSmallInt, 0
      .Columns("NoDown").Properties("NullAble") = True

      .Columns.Append "FontSize", adSmallInt, 0
      .Columns("FontSize").Properties("NullAble") = True

      .Columns.Append "FontName", adVarWChar, 50
      .Columns("FontName").Properties("NullAble") = True
      .Columns("FontName").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "FontStyle", adVarWChar, 50
      .Columns("FontStyle").Properties("NullAble") = True
      .Columns("FontStyle").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "FontUnderline", adBoolean, 2
      '.Columns("FontUnderline").Properties("NullAble") = True

      .Columns.Append "FontStrikeThru", adBoolean, 2
      '.Columns("FontStrikeThru").Properties("NullAble") = True

   End With
   CAT.Tables.Append TBL

   '/* Create Index 'PrimaryKey' */
   Set INDX = New ADOX.Index
   With INDX
      .Name = "PrimaryKey"
      .Columns.Append "ID"
      .PrimaryKey = True
      .Unique = True
      .Clustered = False
      .IndexNulls = adIndexNullsDisallow
   End With
   CAT.Tables("Settings").Indexes.Append INDX
   Set INDX = Nothing

   '/* Create Index 'Description' */
   Set INDX = New ADOX.Index
   With INDX
      .Name = "Description"
      .Columns.Append "Description"
      .PrimaryKey = False
      .Unique = False
      .Clustered = False
      .IndexNulls = adIndexNullsAllow
   End With
   CAT.Tables("Settings").Indexes.Append INDX
   Set INDX = Nothing

   '/* Create Index 'ID' */
   Set INDX = New ADOX.Index
   With INDX
      .Name = "ID"
      .Columns.Append "ID"
      .PrimaryKey = False
      .Unique = False
      .Clustered = False
      .IndexNulls = adIndexNullsAllow
   End With
   CAT.Tables("Settings").Indexes.Append INDX
   Set INDX = Nothing

   Set CAT = Nothing

   Call OpenDB(Mydb, , AppPath & App.Title & ".dat")
   Call OpenRS(MySet, "Select * From Settings", Mydb)
   MySet.AddNew
   MySet!Description = "Default"
   MySet!TopMargin = 0.65
   MySet!SideMargin = 0.15
   MySet!VPitch = 1
   MySet!HPitch = 2.73
   MySet!NoAcross = 3
   MySet!NoDown = 10
   MySet!FontSize = 10
   MySet!FontName = "Times New Roman"
   MySet!FontStyle = "regular"
   MySet!FontUnderline = False
   MySet!FontStrikethru = False
   MySet.Update
   MySet.Close
   Mydb.Close
   Set Mydb = Nothing
   Set MySet = Nothing

Exit Sub


CreateDatERROR:
   MsgBox Err.Description

End Sub

Public Sub CreateMDL(ByVal dbPathFilename As String)

  Dim CAT  As ADOX.Catalog
  Dim TBL  As ADOX.Table
  Dim INDX As ADOX.Index

   On Error GoTo CreateERROR


   '/* Jet OLEDB:Engine Type = 4; (Access 97 database)
   '/* Jet OLEDB:Engine Type = 5; (Access 2000 database)

   Set CAT = New ADOX.Catalog
   CAT.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & dbPathFilename & ";" & _
      "Jet OLEDB:Database Password=;" & _
      "Jet OLEDB:Engine Type=5;"

   '/* Create Table
   Set TBL = New ADOX.Table
   Set TBL.ParentCatalog = CAT
   With TBL
      .Name = "Labels"

      .Columns.Append "ID", adInteger, 0
      .Columns("ID").Properties("AutoIncrement") = True

      .Columns.Append "LINE1", adVarWChar, 30
      .Columns("LINE1").Properties("NullAble") = True
      .Columns("LINE1").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "LINE2", adVarWChar, 30
      .Columns("LINE2").Properties("NullAble") = True
      .Columns("LINE2").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "LINE3", adVarWChar, 30
      .Columns("LINE3").Properties("NullAble") = True
      .Columns("LINE3").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "LINE4", adVarWChar, 30
      .Columns("LINE4").Properties("NullAble") = True
      .Columns("LINE4").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "ZIPCODE", adVarWChar, 10
      .Columns("ZIPCODE").Properties("NullAble") = True
      .Columns("ZIPCODE").Properties("Jet OLEDB:Allow Zero Length") = True
   End With
   CAT.Tables.Append TBL
   Set TBL = Nothing

   '/* Create Index
   Set INDX = New ADOX.Index
   With INDX
      .Name = "ID"
      .Columns.Append "ID"
      .PrimaryKey = True
      .Unique = True
      .Clustered = False
      .IndexNulls = adIndexNullsDisallow
   End With
   CAT.Tables("Labels").Indexes.Append INDX
   Set INDX = Nothing

   Set CAT = Nothing

Exit Sub


CreateERROR:

End Sub

