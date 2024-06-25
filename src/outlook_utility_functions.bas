Attribute VB_Name = "outlook_utility_functions"
Option Explicit

Private Function oneDArrayConvertToTwoDArray(inputArray As Variant) As Variant
    Dim outputArray As Variant

    ReDim outputArray(LBound(inputArray) To UBound(inputArray), 1 To 1)

    Dim i As Long
    For i = LBound(inputArray) To UBound(inputArray)
        outputArray(i, 1) = inputArray(i)
    Next i

    oneDArrayConvertToTwoDArray = outputArray
End Function


Private Function DeleteSubdirectoriesAndFiles(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    ' Delete all subdirectories and files recursively
    DeleteSubdirectories folder
    DeleteFiles folder
    
    ' Clean up
    Set folder = Nothing
    Set fso = Nothing
End Function

Private Function DeleteSubdirectories(ByVal folder As Object)
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        DeleteSubdirectories subfolder ' Recursive call to delete subdirectories
        subfolder.Delete ' Delete the subdirectory
    Next subfolder
End Function

Private Function DeleteFiles(ByVal folder As Object)
    Dim file As Object
    For Each file In folder.Files
        file.Delete ' Delete the file
    Next file
End Function


Private Function RemoveInvalidChars(ByVal inputString As String) As String
    Dim invalidChars As String
    invalidChars = " ~`!@#$%^&*()-+=[]\{}|;':"",./<>?"
    
    Dim resultString As String
    Dim i As Long
    
    For i = 1 To Len(inputString)
        Dim currentChar As String
        currentChar = Mid(inputString, i, 1)
        
        If InStr(invalidChars, currentChar) = 0 Then
            resultString = resultString & currentChar
        End If
    Next i
    
    RemoveInvalidChars = resultString
End Function

Private Function buyerNameIfExistInDir(currentBuyer As Variant)
'handle inside buyer name space or spcial charecter mismatch

'    Dim currentBuyer As Variant
'    currentBuyer = "A.K.M KNIT WEAR LTD"

    Dim currentBuyerTrimed As Variant
    currentBuyerTrimed = Application.Run("outlook_utility_functions.RemoveInvalidChars", currentBuyer)   'remove all invalid characters for use dic keys
    
    Dim buyerNameExistingDict As Object
    Set buyerNameExistingDict = CreateObject("Scripting.Dictionary")
    
    Dim existBuyerArr As Variant
    'existing mismatch buyer
    existBuyerArr = Array( _
    "A.K.M. KNIT WEAR LTD", _
    "THATS IT SPORTS WEAR LTD", _
    "S.F FASHION WEARS LTD", _
    "S.F. DENIM APPARELS LTD", _
    "S.F. JEANS LTD")
    
    Set buyerNameExistingDict = Application.Run("dictionary_utility_functions.AddKeysAndValueSame", buyerNameExistingDict, existBuyerArr)
    
    If buyerNameExistingDict.Exists(currentBuyerTrimed) Then
        buyerNameIfExistInDir = buyerNameExistingDict(currentBuyerTrimed)
    Else
        buyerNameIfExistInDir = currentBuyer
    End If
    
End Function

Private Function PrintPDFDuplexUsingSumatraPDF(ByVal pdfFilePath As String)
    
    ' Specify the full path to SumatraPDF.exe if it's not in the system PATH
    Shell "C:\Users\Humayun\AppData\Local\SumatraPDF\SumatraPDF.exe -print-to-default -print-settings ""duplex=long"" """ & pdfFilePath & """", vbHide
    
End Function


Private Function printPdfUsingAdobeSdk(ByVal filePath As String)
    Dim acroApp As acroApp
    Dim avDoc As AcroAVDoc
    Dim pdDoc As AcroPDDoc

    Set acroApp = New acroApp
    Set avDoc = New AcroAVDoc
    Set pdDoc = New AcroPDDoc

    Dim methodeReturn As Variant
    '  Dim avDoc As Object
    ' Set avDoc = CreateObject("AcroExch.AVDoc")
    
    methodeReturn = acroApp.Hide() ' this methode must call bfore call "Exit()" methode

    If avDoc.Open(filePath, "") Then

        Set pdDoc = avDoc.GetPDDoc()
        methodeReturn = avDoc.PrintPagesSilent(0, pdDoc.GetNumPages - 1, 2, 0, 0)
        
    End If
    
    methodeReturn = acroApp.CloseAllDocs() ' this methode must call bfore call "Exit()" methode
    
    methodeReturn = acroApp.Exit()
    
    Debug.Print "printing by acroApp"
    ' methodeReturn = acroApp.Show() ' if this methode call "Exit()" methode not work

End Function


Private Function ExtractStringLeftOfComma(ByVal inputText As String) As Variant

    Dim commaPosition As Long
    Dim extractedString As String

    ' Find the position of the first comma
    commaPosition = InStr(inputText, ",")

    If commaPosition > 0 Then
        ' Extract the string left of the first comma
        extractedString = Left(inputText, commaPosition - 1)
        ExtractStringLeftOfComma = extractedString
    Else
        ExtractStringLeftOfComma = Null
    End If

End Function


Private Function columnTrimAsB2bOrDashboard(ByVal firstSlotCol As String, ByVal secondSlotCol As String, ByVal isB2B As Boolean)
    
    ' isB2B take boolean value if true it assume fun call for B2B status and add "?" on last column
        Dim curRegionObj As Range
        Set curRegionObj = Range("A1").CurrentRegion
'        set border
        curRegionObj.Borders(xlDiagonalDown).LineStyle = xlNone
        curRegionObj.Borders(xlDiagonalUp).LineStyle = xlNone
        With curRegionObj.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With curRegionObj.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With curRegionObj.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With curRegionObj.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With curRegionObj.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With curRegionObj.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
              
'        delete rows & column
        curRegionObj.Rows(1).EntireRow.Delete Shift:=xlUp
        curRegionObj.Columns(firstSlotCol).EntireColumn.Delete Shift:=xlToLeft
        curRegionObj.Columns(secondSlotCol).EntireColumn.Delete Shift:=xlToLeft

        Dim curRegionVal As Variant
        curRegionVal = curRegionObj.Value

            ' take buyer & bank name only
            Dim i As Long
            For i = 2 To UBound(curRegionVal, 1)
                curRegionVal(i, 1) = i - 1 'sl. no
                curRegionVal(i, 2) = Application.Run("outlook_utility_functions.ExtractStringLeftOfComma", curRegionVal(i, 2)) 'buyer
                curRegionVal(i, 3) = Application.Run("outlook_utility_functions.ExtractStringLeftOfComma", curRegionVal(i, 3)) 'bank
                If isB2B Then
                    curRegionVal(i, UBound(curRegionVal, 2)) = "?" ' for B2B
                End If
            Next i

        curRegionObj.Value = curRegionVal
        
        With Cells
            .Columns.AutoFit
            .Rows.AutoFit
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
        End With

End Function


Private Function returnSelectedFilesFullPathArr(ByVal initialPath As String) As Variant
    Dim fileDialog As Object
    Dim selectedFiles As Variant
    Dim i As Long
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select Files"
        .AllowMultiSelect = True
         .InitialFileName = initialPath
        If .Show = -1 Then
            ReDim selectedFiles(1 To .SelectedItems.Count)
            For i = 1 To .SelectedItems.Count
                selectedFiles(i) = .SelectedItems.Item(i)
            Next i
        End If
    End With

    returnSelectedFilesFullPathArr = selectedFiles
End Function

