Attribute VB_Name = "RPAChallenge"
Dim driver As New ChromeDriver

Sub RPAChallenge()

'Variables
Dim Sheet_Base      As Worksheet
Dim i               As Long
Dim Qtd_Rows        As Long

Set Sheet_Base = ThisWorkbook.Sheets("BASE")
Let Qtd_Rows = WorksheetFunction.CountA(Sheet_Base.Columns(1))

'Open the browser

With driver
    .Start
    .Get "https://rpachallenge.com/"
    .Window.Maximize
End With

'Click on the button Start
driver.FindElementByXPath("//button[contains(text(),'Start')]").Click

For i = 2 To Qtd_Rows

    
    With driver
        
        'First Name
        .FindElementByXPath("//input[@ng-reflect-name='labelFirstName']").Click
        .SendKeys Sheet_Base.Cells(i, 1)
        
        'Last Name
        .FindElementByXPath("//input[@ng-reflect-name='labelLastName']").Click
        .SendKeys Sheet_Base.Cells(i, 2)
        
        'Company Name
        .FindElementByXPath("//input[@ng-reflect-name='labelCompanyName']").Click
        .SendKeys Sheet_Base.Cells(i, 3)
        
        'Role in Company
        .FindElementByXPath("//input[@ng-reflect-name='labelRole']").Click
        .SendKeys Sheet_Base.Cells(i, 4)
        
        'Address
        .FindElementByXPath("//input[@ng-reflect-name='labelAddress']").Click
        .SendKeys Sheet_Base.Cells(i, 5)

        'Email
        .FindElementByXPath("//input[@ng-reflect-name='labelEmail']").Click
        .SendKeys Sheet_Base.Cells(i, 6)
        
        'Phone Number
        .FindElementByXPath("//input[@ng-reflect-name='labelPhone']").Click
        .SendKeys Sheet_Base.Cells(i, 7)
        
        'Click on the button Submit
        .FindElementByXPath("//input[@type='submit']").Click
        
    End With

Next i



End Sub
