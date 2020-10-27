Attribute VB_Name = "Extra_"
Sub getEmaiList()
lrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim EmailAdress()
ReDim EmailAdress(0 To lrow - 2)
i = 2
e = 0
Z = Range("A1").SpecialCells(xlCellTypeLastCell).Address
While Cells.Range("A" & i).Value <> ""
    EmailAdress(e) = Cells.Range("A" & i).Value
    i = i + 1
    e = e + 1
Wend
Sheets("Mail Template").Activate
For Z = LBound(EmailAdress) To UBound(EmailAdress)
    Cells.Range("B3").Value = Cells.Range("B3").Value & ";" & EmailAdress(Z) 'bcc
Next
End Sub

Sub sendEmail()
Dim OutlookApp As Outlook.Application
Dim OutlookMail As Outlook.MailItem

Set OutlookApp = New Outlook.Application
Set OutlookMail = OutlookApp.CreateItem(olMailItem)

With OutlookMail
    .Display
    .HTMLBody = Cells.Range("B6") & "<br><br>" & .HTMLBody
    .BCC = Cells.Range("B3")
    .Subject = Cells.Range("B4")
    '.Send ' uncomment to autosend
End With
End Sub


Sub Main()
Application.ScreenUpdating = False
Call getEmaiList
Call sendEmail
Application.ScreenUpdating = True
End Sub
