Attribute VB_Name = "outlook_send_mail"
Option Explicit

Sub sendUpIssuingStatus()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim wordDoc As Object

    ' Create Outlook application
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0) ' 0 represents a new mail item

    ' Create a Word document within the email
    Set wordDoc = OutMail.GetInspector.WordEditor

    ' Email details
    With wordDoc

        .Range(0, 2).Delete 'delete 2 unwanted line break
        .Range(11, 11).InsertAfter Text:="Fyi..."

    End With

    ' Additional email details
    With OutMail

        .To = "customs2@pioneerdenim.com;rashid.harun54@gmail.com" ' Recipient email address
        .CC = "customs@pioneerdenim.com"
        .Subject = "UP Pending List" ' Email subject
        ' Add attachments
        .Attachments.Add "G:\PDL Customs\Customs Audit 2024-2025\UP Issuing Status # 2024-2025\UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx"

        ' Display or send the email
        .Display

    End With

    ' Clean up
    Set wordDoc = Nothing
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Sub sendDashboardStatus()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim wordDoc As Object

    ' Create Outlook application
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0) ' 0 represents a new mail item

    ' Create a Word document within the email
    Set wordDoc = OutMail.GetInspector.WordEditor
    
    Application.Run "outlook_utility_functions.columnTrimAsB2bOrDashboard", "K:AA", "L:Y", False 'this function current region formated for mail
    
    Range("A1").CurrentRegion.Copy
    
    ' Email details
    With wordDoc
    
        .Range(0, 2).Delete 'delete 2 unwanted line break
        .Range(11, 11).InsertAfter Text:="Please take the necessary step." & vbNewLine
        .Range(44, 44).Paste

    End With
    
    Application.CutCopyMode = False
    ' Additional email details
    With OutMail

        .To = "commercial@pioneerdenim.com;commercial1@pioneerdenim.com;commercial2@pioneerdenim.com" ' Recipient email address
        .CC = "customs@pioneerdenim.com;customs2@pioneerdenim.com"
        .Subject = "Bangladesh bank dashboard" ' Email subject

        ' Display or send the email
        .Display

    End With

    ' Clean up
    Set wordDoc = Nothing
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Sub sendB2bStatus()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim wordDoc As Object

    ' Create Outlook application
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0) ' 0 represents a new mail item

    ' Create a Word document within the email
    Set wordDoc = OutMail.GetInspector.WordEditor

    Application.Run "outlook_utility_functions.columnTrimAsB2bOrDashboard", "K:S", "L:Y", True 'this function current region formated for mail

    Range("A1").CurrentRegion.Copy

    ' Email details
    With wordDoc

        .Range(0, 2).Delete 'delete 2 unwanted line break
        .Range(11, 11).InsertAfter Text:="Below LC We can't proceed in UP, Due to not knowing the back-to-back LC status. Please inform us." & vbNewLine
        .Range(110, 110).Paste

    End With

    Application.CutCopyMode = False
    ' Additional email details
    With OutMail

        .To = "commercial@pioneerdenim.com;commercial1@pioneerdenim.com;commercial2@pioneerdenim.com" ' Recipient email address
        .CC = "customs@pioneerdenim.com;customs2@pioneerdenim.com"
        .Subject = "Back to Back Status" ' Email subject

        ' Display or send the email
        .Display

    End With

    ' Clean up
    Set wordDoc = Nothing
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Sub sendApprovedUp()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim wordDoc As Object

    Dim i As Long

    ' Create Outlook application
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0) ' 0 represents a new mail item

    ' Create a Word document within the email
    Set wordDoc = OutMail.GetInspector.WordEditor
    
    Dim attachedItems As Variant
    attachedItems = Application.Run("outlook_utility_functions.returnSelectedFilesFullPathArr", "G:\PDL Customs\Export LC, Import LC & UP\UP\UP 2024") ' all attachments path

    ' Email details
    With wordDoc

        .Range(0, 2).Delete 'delete 2 unwanted line break
        .Range(11, 11).InsertAfter Text:="Attached approved UP."

    End With

    ' Additional email details
    With OutMail

        .To = "commercial@pioneerdenim.com;" & _
        "commercial1@pioneerdenim.com;" & _
        "commercial2@pioneerdenim.com;" & _
        "incentive@pioneerdenim.com;" & _
        "fabricstore@pioneerdenim.com;" & _
        "yarnstore@pioneerdenim.com;" & _
        "chemicalstore@pioneerdenim.com;" & _
        "pioneerfabric1@gmail.com;" ' Recipient email address
        
        .CC = "customs@pioneerdenim.com;customs2@pioneerdenim.com"
        
        .Subject = "UP- (2024)" ' Email subject

            For i = 1 To UBound(attachedItems, 1)

                ' Add attachments
                .Attachments.Add attachedItems(i)

            Next i

        ' Display or send the email
        .Display

    End With

    ' Clean up
    Set wordDoc = Nothing
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



