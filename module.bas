Option Explicit

Dim ch As Selenium.ChromeDriver



Sub SendWhatsAppMessage()

    
    
 Set ch = New Selenium.ChromeDriver
 
 Dim ws As Worksheet
 
 Set ws = ThisWorkbook.Worksheets("whatsappbot")

 Dim name As String
 Dim message As String
 Dim rep As Double
 
 
    
If ws.Range("c7").Value <> "" And ws.Range("c9").Value <> "" And ws.Range("c11").Value <> "" Then

    
    name = ws.Range("c7").Value
    message = ws.Range("c9").Value
    rep = ws.Range("c11").Value
 
    MsgBox "Browser will open and it will open whatsapp web then you will scan the QR code then you will come back to excel after scaning QR code and you will click ok in excel then it will continue sending message", vbOKOnly, "MANUAL INFO"
    
    
    ch.Start , baseUrl:="https://web.whatsapp.com/"
    ch.Get "/"
    

    
    MsgBox "SEND MESSAGES ? ", vbOKOnly, "CONFIRMATION"
    
    ch.FindElementByXPath("//*[@id='side']/div[1]/div/div/div[2]/div/div[2]").click

 

 
 ch.SendKeys (name)
 ch.Wait (300)
 Dim ks As New Keys
 
 ch.SendKeys (ks.Enter)
 
 ch.Wait (300)
 
 
 Dim i As Integer
 For i = 1 To rep
  

    ch.SendKeys (message)
    ch.SendKeys (ks.Enter)
    ch.Wait (50)
    
    Next i
    
     ch.Wait (5000)
     ch.Wait (5000)
     'ch.Quit
     
     MsgBox "MACRO EXECUTED SUCCESFULLY", vbOKOnly, "SUCCESS"
     
 Else
 
 ws.Activate
 
 MsgBox "Plz Enter Name, Message and Iteration", vbOKOnly, "ERROR, MISSING INPUT"
 
 End If
 
 
 
    
End Sub