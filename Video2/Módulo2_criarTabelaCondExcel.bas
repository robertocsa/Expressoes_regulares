Attribute VB_Name = "Módulo2_criarTabelaCondExcel"
Sub criarTabelaCondicionalParaExcel()
    'Define a variavel do limite para o crédito (somente selecionará as linhas que
    'contiverem crédito maior do que o valor escolhido
    Dim valorLimite As Currency
        
    valorLimite = InputBox("Digite o valor limite", "Valor Mínimo de crédito", 85000)
    
    'Define a variável que representará o documento de origem (texto Word original)
    Dim docOrig As Document
    'Atribui a essa variável o documento atual (o do texto original)
    Set docOrig = ActiveDocument
    
    'Este bloco abaixo foi copiado da gravação da macro de busca, conforme mostrado no vídeo
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "A pessoa*R$ [0-9,]{4;10}."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    ' Aqui acima termina o bloco de busca
    
    'Prepara o cursor do documento de origem para a iteração (leva-o para o início)
    'Desative (comente com uma aspa simples) esta linha abaixo se não quiser que o cursor inicie
    'a partir do início do texto do documento
    docOrig.Select
    Selection.HomeKey Unit:=wdStory
    
    'Define a variável que guardará o documento a ser escrito no final
    '(texto Doc com os dados a exportar para o Excel)
    Dim doc As Document
    
    'Atribui a essa variável o documento recém criado
    Set doc = Documents.Add
    
    'Define a variável que guardará em memória a região que está sendo escrita no
    'documento de destino, em cada passagem
    Dim regiao As Range
    
    'Guarda nessa variável temporária, o conteúdo da região atual do documento
    'de destino
    Set regiao = doc.Range
    
    'Escreve a linha de título da tabela, no texto de destino, com a informação
    'do valor mínimo limite:
    regiao.Text = "Dados com valor de crédito acima de " & FormatCurrency(valorLimite, 2) & vbCrLf
    
    'Prepara o cursor para a inserção (leva-o para o final da região)
    regiao.Collapse wdCollapseEnd
               
    'Escreve a linha de cabecalho, no texto de destino:
    regiao.Text = "Nome" + vbTab + "Valor do crédito" + vbCrLf
    
    'Coloca em negrito o cabecalho
    doc.Range.Bold = True
    
    'Prepara o cursor para a inserção (leva-o para o final da região)
    regiao.Collapse wdCollapseEnd
        
    'Ativa o documento original, ou seja, leva o foco do cursor para ele
    docOrig.Activate
    
    'Bloco de iteração (looping) que executa várias vezes a localização da expressão
    'de interesse
    Do While (Selection.Find.Execute = True)
        'Define a variável que guardará em memória o nome das pessoas, em cada passagem
        Dim nome As String
        'Define a variável que guardará em memória o valor de crédito, em cada passagem
        Dim valor As Currency
        
        'Bloco IF -> controla o fluxo condicional. Somente entra nesse bloco se a
        'condição entre parênteses for satisfeita, ou seja, somente se o valor do
        'crédito for superior ao valor determinado nessa linha de código
        If (getValorCredito(Selection.Text) > valorLimite) Then
            'Atribui à variável nome o nome de cada pessoa encontrada
            'Chama a função getNome para retornar esse nome
            nome = getNome(Selection.Text)
            'Atribui à variável valor o valor do crédito de cada pessoa encontrada
            'Chama a função getValorCredito para retornar esse valor
            valor = getValorCredito(Selection.Text)
            'Debug.Print getValorCredito(Selection.Text)
            'Debug.Print nome
            
            'Ativa (leva o foco do cursor para) o documento de destino (o que está sendo
            'escrito)
            doc.Activate
            'Guarda nessa variável temporária, o conteúdo da região atual do documento
            'de destino
            Set regiao = doc.Range
            
            'Prepara o cursor para a inserção (leva-o para o final da região)
            regiao.Collapse wdCollapseEnd
            
            'Atribui a string de nome + tabulação+ valor + marca de final de parágrafo
            'à região do texto do documento de destino, na posição onde deve ser escrita
            regiao.Text = nome & vbTab & FormatCurrency(valor, 2) & vbCrLf
            
            'Retira o negrito que foi posto no cabecalho
            regiao.Bold = False
            
            'Prepara o cursor para a próxima iteração
            regiao.Collapse wdCollapseEnd

        End If
        'Retorna o foco do cursor para o documento original, para que o loop se reinicie
        'da posição de onde parou (onde localizou a mais recente instância da expressão
        'de busca
        docOrig.Activate
    Loop  ' Final do bloco de iteração (looping)
    
    doc.Activate
    doc.Range.Copy
    
    'Para o funcionamento das linhas de código abaixo, deve-se, antes, incluir a referência
    'à biblioteca do Excel. Para isso, abra o menu ferramenta->referências e selecione a
    'caixa de seleção que contiver o nome Microsoft Excel .. Object Library ou algo parecido
    
    'Define a variável que guardará o objeto de documento do Excel
    Dim planExcel As Excel.Workbook
    
    'Fecha todos os documentos abertos do Excel
    Dim contaDoc As Integer
    contaDoc = 1
    Do While (Excel.Workbooks.Count > 0)
        Excel.Workbooks(contaDoc).Close SaveChanges:=False
        contaDoc = contaDoc + 1
    Loop
        
    'Atribui à variável acima um novo documento do Excel (Workbook do Excel)
    Set planExcel = Excel.Workbooks.Add
    
    'Ativa essa planilha
    planExcel.Activate
    
    'Cola os dados que estão na área de transferência na planilha 1 do
    'documento do Excel recém aberto
    planExcel.Worksheets(1).Paste
    
    'Ajusta a largura automaticamente ao conteúdo (autofit)
    planExcel.Worksheets(1).Range("A:B").Columns.AutoFit
    
    'Aplica o tipo Moeda (Currency) à coluna de valores (coluna B)
    planExcel.Worksheets(1).Columns("B:B").NumberFormat = "$ #,##0.00"
    
    'Salva na pasta documentos
    'planExcel.Save
    
    planExcel.ActiveSheet.Visible = True
    
    planExcel.Application.DisplayFullScreen = True
        
    AppActivate planExcel.Application.Caption
    
    planExcel.Application.DisplayFullScreen = False
    
    'Salva na pasta que você definir (pode ser na pasta atual onde se situa este arquivo)
    Dim strData As String
    strData = Format(Now(), "dd mm yyyy - hh mm ss")
    
    Dim desejaSalvar As Integer
 
    desejaSalvar = MsgBox("Deseja salvar a planilha do Excel " & vbCrLf & _
        "com os dados exportados?", _
        vbQuestion + vbApplicationModal + vbCritical + vbYesNo _
        + vbDefaultButton2 + vbSystemModal, _
        "Clique para salvar os dados exportados")
        
    If (desejaSalvar = vbYes) Then

        planExcel.SaveAs (docOrig.Path & "\Dados Exportados " & strData & ".xlsx")
        doc.SaveAs2 (docOrig.Path & "\Dados Exportados " & strData & ".docx")
    
    End If
    
    'Fecha tanto a planilha quanto o documento de destino Word (texto temporário com os dados
    'do documento Word exportados para a planilha do Excel)
    If (Not planExcel Is Nothing) Then
        'Fecha a planilha Excel que contém os dados exportados
        planExcel.Close SaveChanges:=False
        'Fecha o documento Word
        doc.Close SaveChanges:=False
    End If
        
End Sub 'Final do bloco do procedimento SUB

'Funcão destinada a obter o nome da pessoa, dentro da string (cadeia de caracteres)
'completa da expressão encontrada (ou seja, no caso do exemplo, procura o nome
'dentro da expressão "A pessoa José Pereira tem um crédito de R$ 77474,01." e
'retorna "José Pereira"
Function getNome(texto) As String
    'definição, em memória, da variável que encontrará a posição da palavra "pessoa"
    Dim posPessoa As Integer
    'definição, em memória, da variável que encontrará a posição da palavra "tem"
    Dim posTem As Integer
    'localiza, dentro da expressão completa, a palavra pessoa e guarda sua posição
    posPessoa = InStr(1, texto, "pessoa")
    'localiza, dentro da expressão completa, a palavra tem e guarda sua posição
    posTem = InStr(1, texto, "tem")
    'Definição da variável que guardará a diferença entre as posições
    'das palavras tem e pessoa
    Dim diferenca As Integer
    'Calcula a diferença entre as palavras tem e pessoa
    diferenca = posTem - posPessoa
    'Obtem o nome somente da pessoa em cada iteração
    getNome = Mid(texto, posPessoa + 7, diferenca - 7)
End Function

'Função para obter o valor de crédito em cada expressão encontrada
Function getValorCredito(texto) As Currency
    'Definição da posição onde for encontrado o subtexto (substring) "R$"
    Dim posCifrao As Integer
    'Guarda a posição encontrada do R$ dentro da expressão localizada
    posCifrao = InStr(1, texto, "R$")
    'Guarda, na variável de retorno da função, o valor de crédito encontrado
    getValorCredito = Mid(texto, posCifrao + 3, Len(texto) - posCifrao - 3)
    'Debug.Print getValorCredito

End Function

