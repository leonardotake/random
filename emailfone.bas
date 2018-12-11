Attribute VB_Name = "Módulo1"
Sub contatcxtract()
Dim i, j, col As Integer
Dim texto As String
Dim sht As Worksheet
Set sht = ActiveSheet
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

col = 17
mail = "[a-zA-Z0-9]*(.[a-zA-Z0-9-]+)@[a-zA-Z0-9-]+(.[A-Za-z0-9-]+)*(.[A-Za-z]{2,4})"
fone = "([0-9]{4})(.)+?([0-9]{4})|([0-9]{8})"
lastrow = sht.UsedRange.Rows(sht.UsedRange.Rows.count).Row

    
        
For i = 1 To lastrow
    With regex
        .Pattern = fone
        .Global = False
        .MultiLine = True
    End With
    
    texto = Cells(i, col).Value
    If regex.Test(texto) Then
         texto = Cells(i, col).Value
         Cells(i, col + 1).Value = regex.Execute(texto)(0)
         'Debug.Print i & " " & regex.Execute(texto)(0)
    
    End If
 
    
    With regex
        .Pattern = mail
        .Global = False
        .MultiLine = True
    End With
    
    texto = Cells(i, col).Value
    If regex.Test(texto) Then
         texto = Cells(i, col).Value
         Cells(i, col).Value = regex.Execute(texto)(0)
         Debug.Print i & " " & regex.Execute(texto)(0)
    
    
    
    End If

Next i

End Sub
