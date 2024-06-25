Attribute VB_Name = "outlook_save_attachement"
Option Explicit

Sub SaveOutlookAttachments()

    'This early-binding version requires a reference to the Outlook and Scripting Runtime object libraries
    
    Dim ol As outlook.Application
    Dim ns As outlook.Namespace
    Dim fol As outlook.folder
    Dim i As Object
    Dim mi As outlook.MailItem
    Dim at As outlook.Attachment
    Dim atForTempDir As outlook.Attachment
    Dim fso As Scripting.FileSystemObject
    Dim dir As Scripting.folder
    Dim dirName As String
    
    Set fso = New Scripting.FileSystemObject
    
    Set ol = New outlook.Application
    Set ns = ol.GetNamespace("MAPI")
    Set fol = ns.Folders(1).Folders("Inbox").Folders("Working")  'be changed if need
    
    Dim baseDir As String
    baseDir = "G:\PDL Customs\Export LC, Import LC & UP\Export LC With Related Doc\" 'be changed if need
    
    Dim todaysBaseDir As String
    todaysBaseDir = "C:\Users\Humayun\Desktop\Todays Mail Attachment\" 'be changed if need
    
    Dim tempDir As String
    tempDir = todaysBaseDir & "temp\"
           
'    Dim printingIntervalTime As String
'    printingIntervalTime = InputBox("Enter Printing interval time in seconds:", "User Input")

    Dim printedFileDict As Object
    Set printedFileDict = CreateObject("Scripting.Dictionary")
    printedFileDict("Sl.") = "Subject " & "& File Name"
    Dim mailCount As Integer
    mailCount = 0
    
    Dim previousYear, currentYear, year As Variant
    
    previousYear = "2023" ' be changed if year changed
    currentYear = "2024" ' be changed if year changed
        
    Dim buyerName, lcNo As Variant
    
    Dim extension As String
    
    Dim atFileSavePath As String
    
    Dim blankPdfPrinting As Boolean
    
    Application.Run "outlook_utility_functions.DeleteSubdirectoriesAndFiles", todaysBaseDir ' delete all the files and folders in todays folder
    
    For Each i In fol.Items
    
        If i.Class = olMail Then
        
            Set mi = i
            
            If mi.Attachments.Count > 0 Then
                'Debug.Print mi.SenderName, mi.ReceivedTime, mi.Attachments.Count
                                
                    Dim lcAndBuyerName As Variant
                    lcAndBuyerName = extractMailSubject(mi.Subject)
                    
                    If Not IsNull(lcAndBuyerName) Then
                    
                        lcNo = lcAndBuyerName(0) 'regEx extract
                        buyerName = lcAndBuyerName(1) 'regEx extract
                        buyerName = Application.Run("outlook_utility_functions.buyerNameIfExistInDir", buyerName) ' handle inside buyer name space or spcial character conflict
                                        
                        If fso.FolderExists(baseDir & previousYear & Application.PathSeparator & buyerName & Application.PathSeparator & lcNo) Then
                        
                            ' if same LC No. arrived then year goes to previousYear, so it's handle here
                            
                            Application.Run CreateNestedDirectory(todaysBaseDir, Array("temp"))  'create temp directory
                            
                            Application.Run "outlook_utility_functions.DeleteSubdirectoriesAndFiles", tempDir ' delete all the files and folders in temp folder
                                                            
                            For Each atForTempDir In mi.Attachments
                                                    
                                ' Get the file extension
                                extension = Right$(atForTempDir, 3)
                                
                                If LCase(extension) = "pdf" Then
                                    
                                    atForTempDir.SaveAsFile tempDir & Application.PathSeparator & atForTempDir.Filename
                                        
                                End If
                            Next atForTempDir
                                                            
                            If isAnyDuplicateFileInTwoFolders(tempDir, baseDir & previousYear & Application.PathSeparator & buyerName & Application.PathSeparator & lcNo) Then
                                year = previousYear
                            Else
                                year = currentYear
                            End If
                            
                            Else
                                year = currentYear
                            End If
        
                        
                        Application.Run CreateNestedDirectory(baseDir, Array(year, buyerName, lcNo))  'create directory
                        
                        dirName = baseDir & year & Application.PathSeparator & buyerName & Application.PathSeparator & lcNo
                        
                        
                        If fso.FolderExists(dirName) Then
                            Set dir = fso.GetFolder(dirName)
                        Else
                            Set dir = fso.CreateFolder(dirName)
                        End If
                        
                        blankPdfPrinting = False ' reset
                        
                        For Each at In mi.Attachments
                                                        
                            ' Get the file extension
                                                        
                            extension = Right$(at, 3)    'Hard code take only 3 character and The dollar sign ($) in VBA is used to explicitly declare that a function is returning a string
                            
                            If LCase(extension) = "pdf" Then
                                
                                atFileSavePath = dir.path & Application.PathSeparator & at.Filename
                                
                                If Not fso.FileExists(atFileSavePath) Then
                                
                                    'Debug.Print vbTab, at.DisplayName, at.Size
                                    at.SaveAsFile atFileSavePath ' file save to as a required directory
                                    
                                    If LCase(Left$(at, 2)) = "ud" Then
                                    
'                                        Application.Wait Now + TimeValue("00:00:" & CInt(printingIntervalTime) + 30)    ' time delay here for proper printing
                                    
                                    Else
                                    
'                                        Application.Wait Now + TimeValue("00:00:" & printingIntervalTime)  ' time delay here for proper printing
                                    
                                    End If
                                    
'                                    Application.Run "outlook_utility_using_win_api.PrintPDF", atFileSavePath ' print current at file

'                                    Application.Run "outlook_utility_functions.PrintPDFDuplexUsingSumatraPDF", atFileSavePath ' print current at file
                                    
                                    Application.Run "outlook_utility_functions.printPdfUsingAdobeSdk", atFileSavePath ' print current at file
                                    
                                    printedFileDict(printedFileDict.Count) = "Sub: " & mi.Subject & " & File: " & at.Filename
                                    
                                    blankPdfPrinting = True ' if exist new attachments then set true
                                    
                                End If
                            End If
                        Next at
                        
                        If blankPdfPrinting Then ' if new attachments saved then printing all the attachments of a mail then print a blank pdf
                        
'                            Application.Wait Now + TimeValue("00:00:" & printingIntervalTime) ' time delay here for proper printing
'                            Application.Run "outlook_utility_using_win_api.PrintPDF", "G:\PDL Customs\Export LC, Import LC & UP\Export LC With Related Doc\Blank_PDF_One_Page.pdf" ' print a blank pdf file to separate each mail
                            
'                            Application.Run "outlook_utility_functions.PrintPDFDuplexUsingSumatraPDF", "G:\PDL Customs\Export LC, Import LC & UP\Export LC With Related Doc\Blank_PDF_One_Page.pdf" ' print a blank pdf file to separate each mail
                            
                            Application.Run "outlook_utility_functions.printPdfUsingAdobeSdk", "G:\PDL Customs\Export LC, Import LC & UP\Export LC With Related Doc\Blank_PDF_One_Page.pdf" ' print a blank pdf file to separate each mail
                            
                            mailCount = mailCount + 1
                            
                            printedFileDict(printedFileDict.Count) = "Sub: " & mi.Subject & " & File: " & "# Blank Page # " & "(End of Mail " & mailCount & ")"
                        
                        End If
                        
                        ' todays save attachedment handel here

                        Application.Run CreateNestedDirectory(todaysBaseDir, Array(year, buyerName, lcNo))  'create todays directory
                        Application.Run CopyFilesModifiedToday(dirName, todaysBaseDir & year & Application.PathSeparator & buyerName & Application.PathSeparator & lcNo)

                
                    Else
                        ' this block if not UD LC related subject or regex mismatch
'                        Debug.Print mi.Subject
                        MsgBox mi.Subject & " Mail subject not extracted"
                    End If
            End If
            
        End If
    
    Next i
    
'    Application.Wait Now + TimeValue("00:00:" & printingIntervalTime) ' time delay here for proper printing

    If printedFileDict.Count > 1 Then
        ActiveSheet.Cells.Clear
        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ActiveSheet.Range("a1"), printedFileDict, True, True, True
    End If
    
    MsgBox "All attachment saved"
    
End Sub


Private Function CreateNestedDirectory(baseDir As String, directoryNames As Variant)
    ' function take a base directory string and an array that contain all the sub directory name

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim dirPath As String
    dirPath = baseDir

    Dim level As Integer
    Dim currentDir As Object

    ' Loop through each level and check if it already exists, or create it if it doesn't
    For level = LBound(directoryNames) To UBound(directoryNames)

       If fso.FolderExists(dirPath & Application.PathSeparator & directoryNames(level)) Then
           Set currentDir = fso.GetFolder(dirPath & Application.PathSeparator & directoryNames(level))
       Else
           Set currentDir = fso.CreateFolder(dirPath & Application.PathSeparator & directoryNames(level))
       End If

        ' Update the dirPath for the next iteration
        dirPath = currentDir.path

    Next level

    ' Clean up the FileSystemObject
    Set currentDir = Nothing
    Set fso = Nothing
    
End Function


Private Function extractMailSubject(mailSubject As Variant) As Variant
    Dim regEx As New RegExp
    regEx.Global = True
    regEx.MultiLine = True
    regEx.IgnoreCase = True

'    Dim mailSubject As Variant
'    mailSubject = "LC-1886-STERLING DENIMS LTD"

    
    Dim lcOrScAndBuyer As Variant
    
    regEx.Pattern = "(lc\-.+((ltd)|(limited)))|(((sc\s*no)|(sc\-)).+((ltd)|(limited)))" 'LC/SC + buyer name
    
    If Not regEx.Test(mailSubject) Then
        extractMailSubject = Null
        Exit Function
    End If
    
    Set lcOrScAndBuyer = regEx.Execute(mailSubject)
    lcOrScAndBuyer = lcOrScAndBuyer.Item(0)

    Dim lcNo As Variant

    regEx.Pattern = "(LC\-\d+\-L)|(LC\-\d+)|((sc\s*no)|(sc\-))((\-\d+(\-\d+)?)|(\d+))" 'LC/SC
    Set lcNo = regEx.Execute(lcOrScAndBuyer)
    lcNo = lcNo.Item(0)

    Dim buyerName As Variant
    
    buyerName = regEx.Replace(Trim(lcOrScAndBuyer), "")
    regEx.Pattern = "^\-"
    buyerName = regEx.Replace(Trim(buyerName), "")

    Dim result(0 To 1) As Variant
    result(0) = lcNo
    result(1) = buyerName
    
    extractMailSubject = result
    
End Function


Private Function CopyFilesModifiedToday(sourceDir As String, destDir As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the source directory exists
    If Not fso.FolderExists(sourceDir) Then
        MsgBox "Source directory does not exist."
        Exit Function
    End If
    
    ' Check if the destination directory exists
    If Not fso.FolderExists(destDir) Then
        MsgBox "Destination directory does not exist."
        Exit Function
    End If
    
    Dim sourceFolder As Object, file As Object
    Set sourceFolder = fso.GetFolder(sourceDir)
    
    ' Loop through each file in the source directory
    For Each file In sourceFolder.Files
        ' Check if the file was modified today
        If DateValue(file.DateLastModified) = DateValue(Now) Or DateValue(file.DateCreated) = DateValue(Now) Then
            ' Build the destination file path
            Dim destFilePath As String
            destFilePath = destDir & Application.PathSeparator & fso.GetFileName(file.path)
            
            ' Copy the file to the destination directory
            fso.CopyFile file.path, destFilePath, True
           
        End If
    Next file
    
    ' Clean up the FileSystemObject
    Set file = Nothing
    Set sourceFolder = Nothing
    Set fso = Nothing
End Function


Private Function isTwoPdfFileSame(ByVal filePath1 As String, ByVal filePath2 As String) As Boolean
    ' this function compare pdf file same or not
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the files exist
    If Not fso.FileExists(filePath1) Or Not fso.FileExists(filePath2) Then
        MsgBox "One or both files do not exist."
        isTwoPdfFileSame = False
        Exit Function
    End If
    
    Dim file1 As Object, file2 As Object
    Set file1 = fso.GetFile(filePath1)
    Set file2 = fso.GetFile(filePath2)
    
    ' Compare the file sizes
    If file1.Size <> file2.Size Then
        isTwoPdfFileSame = False
        Exit Function
    End If
    
    ' Open the files in binary mode
    Dim stream1 As Object, stream2 As Object
    Set stream1 = file1.OpenAsTextStream(1, -2)
    Set stream2 = file2.OpenAsTextStream(1, -2)
    
    ' Read the file contents
    Dim content1 As String, content2 As String
    content1 = stream1.ReadAll
    content2 = stream2.ReadAll
    
    ' Compare the file contents
    isTwoPdfFileSame = (content1 = content2)
    
    ' Clean up
    stream1.Close
    stream2.Close
    Set stream1 = Nothing
    Set stream2 = Nothing
    Set file1 = Nothing
    Set file2 = Nothing
    Set fso = Nothing
End Function


 Private Function GetFileNamesInDirectory(ByVal folderPath As String) As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    Dim file As Object
    Dim fileNames() As String
    Dim i As Integer
    
    If folder.Files.Count = 0 Then
        GetFileNamesInDirectory = Null
        Exit Function
    End If
    
    ReDim fileNames(1 To folder.Files.Count)
    
    i = 1
    For Each file In folder.Files
        fileNames(i) = file.Name
        i = i + 1
    Next file
    
    ' Clean up
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    GetFileNamesInDirectory = fileNames
End Function


Function isAnyDuplicateFileInTwoFolders(ByVal sourcePath As String, ByVal comparedPath As String) As Boolean
    
    Dim sourceFilesName As Variant
    Dim comparedFilesName As Variant
    
    sourceFilesName = GetFileNamesInDirectory(sourcePath)
    comparedFilesName = GetFileNamesInDirectory(comparedPath)
    
    Dim sourceFileName  As String
    Dim comparedFileName As String
    
    Dim i As Integer
    Dim j As Integer

    For i = LBound(sourceFilesName) To UBound(sourceFilesName)

        sourceFileName = sourcePath & "\" & sourceFilesName(i)

        For j = LBound(comparedFilesName) To UBound(comparedFilesName)

            comparedFileName = comparedPath & "\" & comparedFilesName(j)
            isAnyDuplicateFileInTwoFolders = isTwoPdfFileSame(sourceFileName, comparedFileName)

            If isAnyDuplicateFileInTwoFolders Then
                
                Exit Function
                
            End If

        Next j

    Next i

    isAnyDuplicateFileInTwoFolders = False

End Function



Private Function DeleteAllFilesInDirectory(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    Dim file As Object
    For Each file In folder.Files
        fso.DeleteFile file.path, True ' Delete the file
    Next file
    
    ' Clean up
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Function

