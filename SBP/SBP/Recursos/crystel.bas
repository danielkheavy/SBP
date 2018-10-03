Attribute VB_Name = "Module14"
Private Sub MyReport(ByVal Crpt As CrystalReport, ByVal Location_rpt As String, ByVal Sformula As String)
        With Crpt 'name of your crystal report
            .DiscardSavedData = True
            .Reset
            .WindowState = crptMaximized
            .ReportFileName = Location_rpt 'set the location of the report
            .SelectionFormula = Sformula   'set SQL statement
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .RetrieveDataFiles
            .Action = 1
        End With
End Sub

