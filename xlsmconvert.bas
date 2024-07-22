Sub ConvertXlsxToXlsm()
    Dim FSO As Object, oFolder As Object, oFile As Object
    Dim WB As Workbook
    Dim NewN As String
    Dim sourceFolder As String
    Dim destinationFolder As String
    Dim CalcMode As XlCalculation
    Dim DisplayAlertsState As Boolean

    ' Store the current calculation mode and display alerts state
    CalcMode = Application.Calculation
    DisplayAlertsState = Application.DisplayAlerts

    ' Set calculation mode to manual and disable screen updating, events, and alerts
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Define the source and destination folders
    sourceFolder = "[FOLDER NAME]"
    destinationFolder = sourceFolder & "converted\"

    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = FSO.GetFolder(sourceFolder)

    ' Check if destination folder exists, create if not
    If Not FSO.FolderExists(destinationFolder) Then
        FSO.CreateFolder (destinationFolder)
    End If

    ' Loop through each file in the folder
    For Each oFile In oFolder.Files
        If LCase(Right(oFile.Name, 5)) = ".xlsx" Then
            ' Open the workbook
            Set WB = Workbooks.Open(oFile.Path, ReadOnly:=True, UpdateLinks:=False)

            ' Define the new file name
            NewN = destinationFolder & FSO.GetBaseName(oFile.Name) & ".xlsm"

            ' Check if the .xlsm file already exists, if not save as .xlsm
            If Dir(NewN) = "" Then
                WB.SaveAs NewN, xlOpenXMLWorkbookMacroEnabled
            End If

            ' Close the workbook
            WB.Close SaveChanges:=False
            
            ' Clear memory
            Set WB = Nothing
            DoEvents
        End If
    Next oFile

    ' Re-enable calculation, screen updating, events, and alerts
    Application.Calculation = CalcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = DisplayAlertsState

    ' Clean up
    Set oFolder = Nothing
    Set FSO = Nothing

    MsgBox "Conversion complete!"
End Sub
