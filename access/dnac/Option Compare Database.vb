Option Compare Database
Option Explicit

' =========================
' LOCAL CONFIG
' =========================
Private Const WORK_FOLDER As String = "C:\CiscoLoaderTest\"
Private Const CLEAN_FILE As String = "C:\CiscoLoaderTest\In_Cisco_clean.csv"
Private Const OUT_FILE As String = "C:\CiscoLoaderTest\In_Cisco.csv"

' DNAC raw export naming patterns
Private Const INPUT_PREFIX1 As String = "Subreport1_"
Private Const INPUT_PREFIX2 As String = "APReport_"

' =========================
' MAIN - RUN THIS
' =========================
Public Sub Run_Cisco_Local_All()

    On Error GoTo ErrHandler

    Dim srcFile As String

    ' 1) Find newest DNAC raw CSV
    srcFile = GetNewestCsv(WORK_FOLDER, INPUT_PREFIX1, INPUT_PREFIX2)
    If srcFile = "" Then
        MsgBox "No DNAC CSV found in: " & WORK_FOLDER & vbCrLf & _
               "Expected a file starting with: " & INPUT_PREFIX1 & " or " & INPUT_PREFIX2, vbExclamation
        Exit Sub
    End If

    ' 2) Clean DNAC CSV (remove garbage rows; keep real header row)
    CleanDnacCsv srcFile, CLEAN_FILE

    ' 3) Clear CiscoRaw then import cleaned CSV into CiscoRaw
    CurrentDb.Execute "DELETE FROM CiscoRaw;", dbFailOnError

    DoCmd.TransferText _
        TransferType:=acImportDelim, _
        SpecificationName:="", _
        TableName:="CiscoRaw", _
        FileName:=CLEAN_FILE, _
        HasFieldNames:=True

    ' 4) Build prod-shaped output in In_Cisco_Final
    '    Option A: If you already have an append query, run it here.
    '    Replace qryCisco_RawToFinal with your real query name if different.
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryCisco_RawToFinal"
    DoCmd.SetWarnings True

    ' 5) Export In_Cisco_Final to required filename In_Cisco.csv (LOCAL)
    '    If your final table name is different, change "In_Cisco_Final"
    DoCmd.TransferText _
        TransferType:=acExportDelim, _
        SpecificationName:="", _
        TableName:="In_Cisco_Final", _
        FileName:=OUT_FILE, _
        HasFieldNames:=True

    MsgBox "DONE ✅" & vbCrLf & _
           "Raw: " & srcFile & vbCrLf & _
           "Clean: " & CLEAN_FILE & vbCrLf & _
           "Exported: " & OUT_FILE, vbInformation

    Exit Sub

ErrHandler:
    DoCmd.SetWarnings True
    MsgBox "FAILED ❌ " & Err.Number & " - " & Err.Description, vbCritical

End Sub

' =========================
' HELPERS
' =========================

Private Function GetNewestCsv(ByVal folderPath As String, ByVal prefix1 As String, ByVal prefix2 As String) As String

    Dim fso As Object, fldr As Object, fil As Object
    Dim newestFile As String
    Dim newestDate As Date
    Dim nm As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        GetNewestCsv = ""
        Exit Function
    End If

    Set fldr = fso.GetFolder(folderPath)
    newestFile = ""
    newestDate = #1/1/2000#

    For Each fil In fldr.Files
        nm = fil.Name

        If LCase(fso.GetExtensionName(nm)) = "csv" Then
            If LCase(Left(nm, Len(prefix1))) = LCase(prefix1) _
               Or LCase(Left(nm, Len(prefix2))) = LCase(prefix2) Then

                If fil.DateLastModified > newestDate Then
                    newestDate = fil.DateLastModified
                    newestFile = fil.Path
                End If
            End If
        End If
    Next fil

    GetNewestCsv = newestFile

End Function


' Finds the real header row (contains AP MAC Address + Device Name + IP Address)
' and writes that line + everything after it.
Private Sub CleanDnacCsv(ByVal inPath As String, ByVal outPath As String)

    Dim fso As Object
    Dim tsIn As Object, tsOut As Object
    Dim line As String
    Dim foundHeader As Boolean
    Dim norm As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tsIn = fso.OpenTextFile(inPath, 1, False)   'ForReading
    Set tsOut = fso.CreateTextFile(outPath, True)   'overwrite

    foundHeader = False

    Do While Not tsIn.AtEndOfStream
        line = tsIn.ReadLine

        norm = LCase$(Trim$(line))
        norm = Replace(norm, """", "")

        If Not foundHeader Then
            If InStr(1, norm, "ap mac address", vbTextCompare) > 0 _
               And InStr(1, norm, "device name", vbTextCompare) > 0 _
               And InStr(1, norm, "ip address", vbTextCompare) > 0 Then

                foundHeader = True
                tsOut.WriteLine Replace(Trim$(line), """", "")
            End If
        Else
            tsOut.WriteLine Replace(line, """", "")
        End If
    Loop

    tsIn.Close
    tsOut.Close

    If Not foundHeader Then
        Err.Raise vbObjectError + 1000, "CleanDnacCsv", _
                  "Could not find header row in: " & inPath
    End If

End Sub