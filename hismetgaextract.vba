Attribute VB_Name = "Módulo5"
Sub hismetga()
Dim i, j, final As Integer
Dim count As Long
Dim texto, data, id, nome As String
Dim subtexto
Dim sht As Worksheet
Set sht = ActiveSheet
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

data = "([0-9]{2})(\/)([0-9]{2})(\/)([0-9]{4})"
id = "[0-9]{6}"
artefatos = "[^0-9]{3}"

final = sht.UsedRange.Rows(sht.UsedRange.Rows.count).Row


For i = 1 To final
    
    'deleta celulas vazias e retira espaços
    If Cells(i, 1).Value = "" Then
        Cells(i, 1).Delete
    End If
    texto = Cells(i, 1).Value
    Cells(i, 1).Value = Trim(texto)
    
Next i

final = sht.UsedRange.Rows(sht.UsedRange.Rows.count).Row

For i = 1 To final
 
 'deleta celulas vazias e retira espaços
    If Cells(i, 1).Value = "" Then
       Cells(i, 1).Delete
   
    End If
    
    count = 0
    texto = Cells(i, 1).Value
    
    'substitui linhas vazias maiores que dois
        For j = 1 To Len(texto)
            
            If Mid(texto, j, 1) = " " Then
                count = count + 1
            Else
                count = 0
            End If
            If count > 2 Then
                Mid(texto, j, 1) = "-"
            End If
         
         Next j
     
     texto = Replace(texto, "-", "", 1)
     texto = Replace(texto, "  ", "|", 1)
     Cells(i, 1) = texto
     subtexto = Split(texto, "|", -1)
     
    'etapa de expressao regulares
 
     With regex
        .Pattern = id
        .Global = False
        .MultiLine = False
        End With
        
        If regex.Test(texto) Then
            If Len(subtexto(0)) = 6 Then
           nome = subtexto(1)
        Else
            Cells(i, 1).Delete
          
           End If
        End If
     With regex
     .Pattern = data
     .Global = False
     .MultiLine = False
     End With
        
        If regex.Test(texto) Then
            If subtexto(1) = "PSICOLOGA" Then
            Cells(i, 1).Value = nome
            Cells(i, 2).Value = subtexto(3)
            Cells(i, 3).Value = subtexto(4)
            
            
            
            ElseIf subtexto(2) = "ELETROCARDIOGRAMA" Then
            Cells(i, 1).Value = nome
            Cells(i, 2).Value = subtexto(3)
            Cells(i, 3).Value = subtexto(4)
            
            Else
            Cells(i, 1).Value = nome
            Cells(i, 2).Value = subtexto(2)
            Cells(i, 3).Value = subtexto(3)
            
            End If
        End If
                       
                      
Next i

final = sht.UsedRange.Rows(sht.UsedRange.Rows.count).Row

For j = 1 To 3
    For i = 1 To final
        If (Cells(i, 1).Value = "" Or Cells(i, 2).Value = "" Or Cells(i, 3).Value = "") Then
         Cells(i, 1).EntireRow.Delete
        End If
    Next i
Next j

Cells(1, 1).EntireRow.Insert
Cells(1, 1).Value = "Nome"
Cells(1, 2).Value = "Exame"
Cells(1, 3).Value = "Valor"


End Sub
