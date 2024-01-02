Sub ExcelOutlook()

    'FORMAT CELLS TO TEXT TO PREVENT EXCEL FROM DELETING LEADING ZEROES

    'OPENS WORD DOCUMENT TEMPLATE AND INPUTS SPECIFIED VALUES

    Dim book1 As Word.Application
    Dim sheet1 As Word.Document

    'Set saved location for your Blank Email Template , ex. C:\Users\user\Desktop\Templates
    Email_Template_Blank = Range("D2")

    'Opens up word document
Set book1 = CreateObject("word.application")
book1.Visible = True
Set sheet1 = book1.Documents.Open(Email_Template_Blank)

    'Set fields you want to edit and change
    Word_EMAIL = Range("B1")
    Word_Example1 = Range("B3")
    Word_Example2 = Range("B4")
    Word_Example3 = Range("B5")
    Word_Example4 = Range("B6")
    Word_Example5 = Range("B7")
    Word_DATE = Date

    'Updates the Date automatically to current system date

    With sheet1.Content.Find
        .Text = "DATE:"
        .Replacement.Text = "DATE: " & Word_DATE
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
End With

    'Replace first field in a specific location within your template, simply copy/paste preceeding text in template  
    'essentially giving code a place to control + f (find) where you want to replace in template.

    With sheet1.Content.Find
        .Text = "Example 1: "
        .Replacement.Text = "Example 1: " & Word_Example1
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
End With

    With sheet1.Content.Find
        .Text = "Example 2: "
        .Replacement.Text = "Example 2: " & Word_Example2
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
End With

    With sheet1.Content.Find
        .Text = "Example 3: "
        .Replacement.Text = "Example 3: " & Word_Example3
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
End With

    With sheet1.Content.Find
        .Text = "Example 4: "
        .Replacement.Text = "Example 4: " & Word_Example4
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
End With

    With sheet1.Content.Find
        .Text = "Example 5: "
        .Replacement.Text = "Example 5: " & Word_Example5
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
End With



    'SAVES WORD DOC TO PDF IN LOCAL FOLDER

    Dim strClient As String
    Dim strPDFname As String
    Dim strClientEmail As String


    'Location you want to save pdf 
    strClientEmail = Range("D12")
    strPDFname = "PDF_Name_" & Example1


    'Saves word document and converts to PDF
    book1.ActiveDocument.ExportAsFixedFormat OutputFileName:=
                                       strClientEmail & strPDFname & ".pdf",
                                       ExportFormat:=wdExportFormatPDF,
                                       OpenAfterExport:=False,
                                       OptimizeFor:=wdExportOptimizeForPrint,
                                       Range:=wdExportAllDocument,
                                       IncludeDocProps:=True,
                                       CreateBookmarks:=wdExportCreateWordBookmarks,
                                       BitmapMissingFonts:=True

'AFTER FIRST SAVE FILE, MUST SET WORD DOCUMENT TO ACTIVE OR ERROR OCCURS


    book1.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges

'OPENS OUTLOOK, SETS TO , SUBJECT , AND ADD ATTACHMENT FIELDS

    Dim OutlookApp As Outlook.Application
    Dim OutlookMail As Outlook.MailItem

  Set OutlookApp = New Outlook.Application
  Set OutlookMail = OutlookApp.CreateItem(olMailItem)
  
  With OutlookMail
        .BodyFormat = olFormatHTML
        .Display
        email_TO = Range("B1")
        .To = email_TO
        .Subject = "EMAIL SUBJECT"
        .Attachments.Add strClientEmail & strPDFname & ".pdf"


  End With
End Sub