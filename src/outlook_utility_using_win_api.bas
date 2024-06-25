Attribute VB_Name = "outlook_utility_using_win_api"
Option Explicit


Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr


    Private Function PrintPDF(ByVal pdfFilePath As String)
        Dim result As LongPtr
    
        ' Open the PDF file with the default associated application (usually a PDF reader)
        result = ShellExecute(0, "print", pdfFilePath, vbNullString, vbNullString, vbNormalNoFocus)
    
        If result <= 32 Then
            Debug.Print "Failed to print the PDF file.", vbExclamation
        Else
            Debug.Print "Printing the PDF file...", vbInformation
        End If
    End Function
    
 
