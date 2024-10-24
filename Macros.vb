Sub Entrada_Estoque()
'
' Entrada_Estoque Macro
'

'
        Call Verifica_Vazio
        
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
    
        Sheets("Rastreabilidade").Rows("2:2").Insert Shift:=xlDown
    
        Sheets("Ajustes").Unprotect Password:="8494"
        
        Sheets("Rastreabilidade").Range("A2") = Sheets("Ajustes").Range("C2")
        Sheets("Rastreabilidade").Range("B2") = Sheets("Ajustes").Range("B5")
        Sheets("Rastreabilidade").Range("C2") = Sheets("Ajustes").Range("C5")
        Sheets("Rastreabilidade").Range("D2") = Sheets("Ajustes").Range("D5")
        Sheets("Rastreabilidade").Range("E2") = Sheets("Ajustes").Range("E5")
        Sheets("Rastreabilidade").Range("F2") = Sheets("Ajustes").Range("F5")
    
        Sheets("Ajustes").Range("B5:F5").ClearContents
    
        ActiveWorkbook.RefreshAll
    
        ActiveSheet.Protect Password:="8494"
    
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
End Sub

Sub Saida_Estoque()
'
' Saida_Estoque Macro
'

'
    Call Verifica_Vazio
    Call Valida_Estoque
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Sheets("Rastreabilidade").Rows("2:2").Insert Shift:=xlDown
    
    Sheets("Ajustes").Unprotect Password:="8494"
    
    Sheets("Rastreabilidade").Range("A2") = Sheets("Ajustes").Range("C2")
    Sheets("Rastreabilidade").Range("B2") = Sheets("Ajustes").Range("B5")
    Sheets("Rastreabilidade").Range("C2") = Sheets("Ajustes").Range("C5").Value * (-1)
    Sheets("Rastreabilidade").Range("D2") = Sheets("Ajustes").Range("D5")
    Sheets("Rastreabilidade").Range("E2") = Sheets("Ajustes").Range("E5")
    Sheets("Rastreabilidade").Range("F2") = Sheets("Ajustes").Range("F5")
    
    Sheets("Ajustes").Range("B5:F5").ClearContents
    
    ActiveWorkbook.RefreshAll
    
    ActiveSheet.Protect Password:="8494"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Sub Verifica_Vazio()
    If IsEmpty(Sheets("Ajustes").Range("B5").Value) Or _
        IsEmpty(Sheets("Ajustes").Range("C5").Value) Or _
        IsEmpty(Sheets("Ajustes").Range("D5").Value) Or _
        IsEmpty(Sheets("Ajustes").Range("E5").Value) Then
            MsgBox "Há uma célula obrigatória sem valor! Verifique os campos e tente novamente.", _
               vbExclamation + vbOKOnly, "Aviso de Preenchimento"
            End
    End If
    
End Sub

Sub Valida_Estoque()

    Dim wsAjustes As Worksheet
    Dim wsSubtotal As Worksheet
    Dim codigoSubtotal As Range
    Dim osSubtotal As Range
    Dim localSubtotal As Range
    Dim codigoAjustes As String
    Dim osAjustes As String
    Dim localAjustes As String
    Dim qtdAjustes As Double
    Dim estoqueSubtotal As Double
    Dim resultado As Range
    
    ' Definir as planilhas
    Set wsAjustes = Sheets("Ajustes")
    Set wsSubtotal = Sheets("Subtotal")
    
    ' Obter os valores das células B5, D5, E5 e C5 da planilha Ajustes
    codigoAjustes = wsAjustes.Range("B5").Value
    osAjustes = wsAjustes.Range("D5").Value
    localAjustes = wsAjustes.Range("E5").Value
    qtdAjustes = wsAjustes.Range("C5").Value
    
    ' Definir os intervalos de busca nas colunas específicas da planilha Subtotal
    Set codigoSubtotal = wsSubtotal.Range("B:B") ' Coluna B para procurar código (B5)
    Set osSubtotal = wsSubtotal.Range("C:C")    ' Coluna C para procurar OS (D5)
    Set localSubtotal = wsSubtotal.Range("E:E")  ' Coluna E para procurar local (E5)
    
    ' Procurar o código (B5) na coluna B da planilha Subtotal
    Set resultado = codigoSubtotal.Find(What:=codigoAjustes, LookIn:=xlValues, LookAt:=xlWhole)
    If resultado Is Nothing Then
        MsgBox "Produto (" & codigoAjustes & ") não encontrado nos registros.", vbExclamation, "Erro na Busca"
        End
    End If
    
    ' Se código foi encontrado, procurar a OS (D5) na coluna C
    Set resultado = osSubtotal.Find(What:=osAjustes, LookIn:=xlValues, LookAt:=xlWhole)
    If resultado Is Nothing Then
        MsgBox "OS (" & osAjustes & ") desse produto não encontrada nos registros.", vbExclamation, "Erro na Busca"
        End
    End If
    
    ' Se OS foi encontrada, procurar o local (E5) na coluna E
    Set resultado = localSubtotal.Find(What:=localAjustes, LookIn:=xlValues, LookAt:=xlWhole)
    If resultado Is Nothing Then
        MsgBox "Local (" & localAjustes & ") não encontrado para esse produto e OS.", vbExclamation, "Erro na Busca"
        End
    End If
    
    ' Se local foi encontrado, verificar se a quantidade (C5) é maior que o valor na coluna D (estoque)
    estoqueSubtotal = wsSubtotal.Cells(resultado.Row, 4).Value ' Obter o valor na coluna D (estoque)
    
    If qtdAjustes > estoqueSubtotal Then
        MsgBox "Quantidade(" & qtdAjustes & ") é maior que o estoque(" & estoqueSubtotal & ").", vbExclamation, "Quantidade Excedente"
        End
    End If

End Sub
