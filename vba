Sub balancete_pia()

Dim ws As Worksheet
    Dim ultimalinha As Long
    Dim i As Long
    Dim j As Long
    Dim rngVisivel As Range
    Dim cel As Range
    Dim cc As String
    
    
    Set ws = ActiveSheet
    
 ' Desativa atualização de tela e cálculo
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
 'Copiar o valor da célula A4 e armazenar
  
  ws.Range("A5").Formula = "=TEXTAFTER(A4,"": "")"
  cc = CStr(ws.Range("A5").Value)
  
  
 
    
'Deletar primeiras linhas
For A = 1 To 7
    ws.Rows(1).Delete
    Next A

'Tirar a mescla
ws.Cells.UnMerge
   
'Fonte
ws.Cells.Font.Name = "Arial"
ws.Cells.Font.Size = 8
    
' Última linha baseada na coluna H:M
    ultimalinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "I").End(xlUp).Row > ultimalinha Then ultimalinha = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "J").End(xlUp).Row > ultimalinha Then ultimalinha = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "K").End(xlUp).Row > ultimalinha Then ultimalinha = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "L").End(xlUp).Row > ultimalinha Then ultimalinha = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "M").End(xlUp).Row > ultimalinha Then ultimalinha = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    
' Mover para a coluna E
    For i = 2 To ultimalinha
        For j = 8 To 13
            If ws.Cells(i, j).Value <> "" Then
                ws.Cells(i, "E").Value = ws.Cells(i, j).Value
                ws.Cells(i, j).ClearContents '
                Exit For '
            End If
        Next j
    Next i
    
'Deletar colunas
Columns("B:C").Delete
Columns("D:Q").Delete
Columns("E").Delete
Columns("F").Delete
Columns("G:I").Delete

With ws
.Range("A1").Value = "Reduzido"
.Range("B1").Value = "Extenso"
.Range("C1").Value = "Descrição"
.Range("H1").Value = "Saldo"
.Range("I1").Value = "Dif"
.Range("J1").Value = "Obs"
End With

Columns("C").ColumnWidth = 50
Columns("A").ColumnWidth = 9
Columns("A:A").HorizontalAlignment = xlLeft
Rows(1).Font.Bold = True

'Filtrar linhas em negrito
If ws.AutoFilterMode Then ws.AutoFilterMode = False
      
ultimalinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
ws.Range("B1:B" & ultimalinha).AutoFilter
For i = 2 To ultimalinha
        If ws.Cells(i, "B").Font.Bold = False Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
For i = 2 To ultimalinha
    If Not ws.Rows(i).Hidden Then
        ws.Rows(i).Interior.Color = RGB(200, 200, 200) ' cinza
        ws.Rows(i).Font.Bold = True
    End If
Next i

ws.Columns.Hidden = False

'Congelar a linha 1
Rows(2).Select
ActiveWindow.FreezePanes = True

' Limpa o filtro e formata colunas
ws.AutoFilterMode = False
ws.Range("A1:J1").AutoFilter
Columns("B").ColumnWidth = 14
ws.Columns("D:G").AutoFit
ws.Columns("I").AutoFit
Columns("H").ColumnWidth = 16
If ws.AutoFilterMode Then
    ws.AutoFilterMode = False
    End If

'Ajustar Coluna D removendo o caractere especial
Columns("D:D").Insert
ultimalinha = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
For i = 2 To ultimalinha
    ws.Cells(i, "D").Formula = "=VALUE(E" & i & ")"
    Next i
ws.Range("D2:D" & ultimalinha).Value = ws.Range("D2:D" & ultimalinha).Value
Columns("E").Delete

'Ajustar Coluna G removendo o caractere especial
Columns("G:G").Insert
ultimalinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
For i = 2 To ultimalinha
        ws.Cells(i, "G").Formula = "=VALUE(H" & i & ")"
    Next i
ws.Range("G2:G" & ultimalinha).Value = ws.Range("G2:G" & ultimalinha).Value
Columns("H").Delete

If ws.AutoFilterMode Then
    ws.AutoFilterMode = False
    End If
    
With ws
.Columns("I").AutoFit
.Columns("D").AutoFit
.Rows(1).AutoFilter
.Range("D1").Value = "Saldo anterior"
.Range("G1").Value = "Saldo atual"
End With


'Adicionar se é D ou C
Columns("I:I").Insert
ws.Range("H1").Value = "C/D"
ultimalinha = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
For i = 2 To ultimalinha
    ws.Cells(i, "H").Formula = "=IFS(G" & i & ">0,""D"",G" & i & "<0,""C"",TRUE,""0"")"
    Next i
ws.Range("H2:G" & ultimalinha).Value = ws.Range("H2:G" & ultimalinha).Value

'Transformar em positivo
For i = 2 To ultimalinha
    ws.Cells(i, "G").Value = Abs(ws.Cells(i, "G").Value)
Next i

'Fórmula de diferença
ultimalinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
ws.Range("A1:Y" & ultimalinha).AutoFilter Field:=10, Operator:=xlFilterNoFill

On Error Resume Next
Set rngVisivel = ws.Range("J2:J" & ultimalinha).SpecialCells(xlCellTypeVisible)
On Error GoTo 0
If Not rngVisivel Is Nothing Then
    For Each cel In rngVisivel
        cel.Formula = "=G" & cel.Row & "-I" & cel.Row
    Next cel
End If

With ws
.Columns("H").AutoFit
.Columns("J").AutoFit
.Range("I1").Value = "Saldo"
End With


If ws.AutoFilterMode Then
    ws.ShowAllData
End If

Columns("E:H").HorizontalAlignment = xlLeft

'Cria a coluna ordem
Columns("A:A").Insert
Columns("A:A").Insert
ws.Range("A1").Value = "Ordem"
ws.Range("B1").Value = "CC"
ws.Range("B1").Value = cc



contador = 1
ultimalinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
For i = 2 To ultimalinha
    ws.Cells(i, "A").Value = contador
    contador = contador + 1
Next i


 ' Reativa atualização de tela e cálculo
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
     
MsgBox ("Formatação concluída")
End Sub
