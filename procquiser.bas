Attribute VB_Name = "Módulo1"
Sub seisdozedezoito()
Dim i, j, col As Integer
Dim texto, exames As String
Dim sht1 As Worksheet
Set sht1 = ThisWorkbook.Sheets("Plan2")
Dim sht2 As Worksheet
Set sht2 = ThisWorkbook.Sheets("Plan3")

lastrow1 = sht1.UsedRange.Rows(sht1.UsedRange.Rows.count).Row
lastrow2 = sht2.UsedRange.Rows(sht2.UsedRange.Rows.count).Row

For i = 1 To lastrow1
    cnpja = Sheets("Plan2").Cells(i, 1).Value
    '''Debug.Print i & " "
    For j = 2 To lastrow2
        
        cnpjb = Sheets("Plan3").Cells(j, 1).Value
        exames = Sheets("Plan3").Cells(j, 3).Value
        valor = Sheets("Plan3").Cells(j, 4).Value
    
        If (cnpja = cnpjb) And exames = "AUDIO" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 7).Value = valor
        
        ElseIf (cnpja = cnpjb) And exames = "CONSULTA" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 8).Value = valor
        ElseIf (cnpja = cnpjb) And exames = "CONSULTA ESPECIALISTA" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 9).Value = valor
        ElseIf (cnpja = cnpjb) And exames = "EXAMES" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 10).Value = valor
        ElseIf (cnpja = cnpjb) And exames = "INCLUSÃO PGR" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 11).Value = valor
        ElseIf (cnpja = cnpjb) And exames = "INCLUSÃO PPRA" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 12).Value = valor
        ElseIf (cnpja = cnpjb) And exames = "PSICO" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 13).Value = valor
        ElseIf (cnpja = cnpjb) And exames = "EXAME TOXICOLOGICO PARA MOTORISTA" Then
                Debug.Print j, valor, 1
                Sheets("plan2").Cells(i, 14).Value = valor
       
        End If
    
    Next j
Next i




End Sub
