Sub ImportPrimavera()
    Dim XERFilePath As String
    Dim FSO As Object
    Dim TS As Object
    Dim LineData As String
    Dim ParsedData As Variant
    Dim i As Long

    ' Specify the path to the .xer file
    XERFilePath = Application.GetOpenFilename(FileFilter:="XER Files (*.xer), *.xer", Title:="Select Primavera P6 .xer File")

    If XERFilePath = "False" Then
        MsgBox "No file selected!"
        Exit Sub
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TS = FSO.OpenTextFile(XERFilePath, 1)

    ' Clear existing data in the worksheet
    Sheets("Data").Cells.Clear

    ' Loop through each line in the XER file
    i = 1
    Do While Not TS.AtEndOfStream
        LineData = TS.ReadLine
        ParsedData = Split(LineData, ",") ' Assuming comma-separated values in the file

        ' Output parsed data to the worksheet
        For j = LBound(ParsedData) To UBound(ParsedData)
            Sheets("Data").Cells(i, j + 1).Value = ParsedData(j)
        Next j
        i = i + 1
    Loop

    TS.Close
    MsgBox "Import complete!"
End Sub