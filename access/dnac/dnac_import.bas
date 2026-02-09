Option Compare Database
Option Explicit

Private Sub cmdRunDnacImport_Click()
    On Error GoTo ErrHandler

    Dim fd As Office.FileDialog
    Dim csvPath As String

    '--- Pick the DNAC CSV file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select DNAC Subreport CSV"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False

        If .Show <> -1 Then
            MsgBox "Cancelled. No file selected.", vbInformation
            Exit Sub
        End If

        csvPath = .SelectedItems(1)
    End With

    '--- Step 1: Clear existing data (prevents duplicates + mismatches)
    CurrentDb.Execute "DELETE * FROM CiscoRaw;", dbFailOnError
    CurrentDb.Execute "DELETE * FROM In_Cisco_Final;", dbFailOnError

    '--- Step 2: Import into CiscoRaw using your saved import specification
    ' Spec name must match exactly: DNAC_Staging
    DoCmd.TransferText _
        TransferType:=acImportDelim, _
        SpecificationName:="DNAC_Staging", _
        TableName:="CiscoRaw", _
        FileName:=csvPath, _
        HasFieldNames:=True

    '--- Step 3: Transform / load into final table
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryCisco_RawToFinal"
    DoCmd.SetWarnings True

    MsgBox "DNAC Import complete ✅", vbInformation

    'Optional: open results
    DoCmd.OpenTable "In_Cisco_Final", acViewDatasheet

ExitHere:
    On Error Resume Next
    DoCmd.SetWarnings True
    Set fd = Nothing
    Exit Sub

ErrHandler:
    DoCmd.SetWarnings True
    MsgBox "Import failed ❌" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitHere
End Sub