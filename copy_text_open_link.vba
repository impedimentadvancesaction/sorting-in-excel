Sub ButtonClick()

    Dim clipText As String
    Dim DataObj As Object

    ' Build the text
    clipText = "This is the fixed text that must be copied." & vbCrLf & _
               "Value from A4: " & Range("A4").Value

    ' Copy to clipboard
    Set DataObj = CreateObject("MSForms.DataObject")
    DataObj.SetText clipText
    DataObj.PutInClipboard

    ' Inform the user
    MsgBox "The required text has been copied to your clipboard." & vbCrLf & _
           "Click OK to continue.", vbInformation, "Copied"

    ' Open link in Chrome
    Shell "cmd /c start chrome ""https://www.example.com""", vbHide

End Sub
