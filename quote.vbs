
Option Explicit

Dim WithEvents objInspectors As Outlook.Inspectors
Dim WithEvents objEmail As Outlook.MailItem

Private Sub Application_Startup()
  Set objInspectors = Application.Inspectors
End Sub

Private Sub objInspectors_NewInspector(ByVal Inspector As Inspector)
  If Inspector.CurrentItem.Class <> olMail Then Exit Sub
  Set objEmail = Inspector.CurrentItem
End Sub

Private Sub objEmail_Open(Cancel As Boolean)
   If objEmail.LastModificationTime <> "1/1/4501" Then Exit Sub
   Dim strHTML
   strHTML = objEmail.HTMLBody
   Dim p
   p = InStr(1, strHTML, "%QUOTE%", vbTextCompare)
   If (p <> 0) Then
      
      Dim s1, s2
      s1 = Left(strHTML, p - 1)
      s2 = Mid(strHTML, p + 7)
      objEmail.HTMLBody = s1 & vbCrLf & RandomQuote() & s2
      objEmail.Save
   End If
   End Sub

Private Sub Application_Quit()
     Set objInspectors = Nothing
     Set objEmail = Nothing
     
End Sub

Public Function RandomQuote() As String
    Dim sLines As String
    Dim sQuote As String
    Dim FileNumber, RandomLineNumber, i
    FileNumber = FreeFile
    Open "d:\quotes.txt" For Input As #FileNumber
      Line Input #FileNumber, sLines
      Randomize
      RandomLineNumber = Int(Rnd * CInt(sLines)) + 1
    For i = 1 To RandomLineNumber
       Line Input #FileNumber, sQuote
    Next
    RandomQuote = sQuote
    Close #FileNumber
End Function

Public Sub AddQuote()

    Dim myItem As Outlook.MailItem
    Set myItem = Application.ActiveInspector.CurrentItem
    myItem.Body = myItem.Body & vbCrLf & RandomQuote() & vbCrLf
    myItem.Save

End Sub
