Attribute VB_Name = "M�dulo2_criarTabelaCondExcel"
Sub criarTabelaCondicionalParaExcel()
    'Define a variavel do limite para o cr�dito (somente selecionar� as linhas que
    'contiverem cr�dito maior do que o valor escolhido
    Dim valorLimite As Currency
        
    valorLimite = InputBox("Digite o valor limite", "Valor M�nimo de cr�dito", 85000)
    
    'Define a vari�vel que representar� o documento de origem (texto Word original)
    Dim docOrig As Document
    'Atribui a essa vari�vel o documento atual (o do texto original)
    Set docOrig = ActiveDocument
    
    'Este bloco abaixo foi copiado da grava��o da macro de busca, conforme mostrado no v�deo
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
    
    'Prepara o cursor do documento de origem para a itera��o (leva-o para o in�cio)
    'Desative (comente com uma aspa simples) esta linha abaixo se n�o quiser que o cursor inicie
    'a partir do in�cio do texto do documento
    docOrig.Select
    Selection.HomeKey Unit:=wdStory
    
    'Define a vari�vel que guardar� o documento a ser escrito no final
    '(texto Doc com os dados a exportar para o Excel)
    Dim doc As Document
    
    'Atribui a essa vari�vel o documento rec�m criado
    Set doc = Documents.Add
    
    'Define a vari�vel que guardar� em mem�ria a regi�o que est� sendo escrita no
    'documento de destino, em cada passagem
    Dim regiao As Range
    
    'Guarda nessa vari�vel tempor�ria, o conte�do da regi�o atual do documento
    'de destino
    Set regiao = doc.Range
    
    'Escreve a linha de t�tulo da tabela, no texto de destino, com a informa��o
    'do valor m�nimo limite:
    regiao.Text = "Dados com valor de cr�dito acima de " & FormatCurrency(valorLimite, 2) & vbCrLf
    
    'Prepara o cursor para a inser��o (leva-o para o final da regi�o)
    regiao.Collapse wdCollapseEnd
               
    'Escreve a linha de cabecalho, no texto de destino:
    regiao.Text = "Nome" + vbTab + "Valor do cr�dito" + vbCrLf
    
    'Coloca em negrito o cabecalho
    doc.Range.Bold = True
    
    'Prepara o cursor para a inser��o (leva-o para o final da regi�o)
    regiao.Collapse wdCollapseEnd
        
    'Ativa o documento original, ou seja, leva o foco do cursor para ele
    docOrig.Activate
    
    'Bloco de itera��o (looping) que executa v�rias vezes a localiza��o da express�o
    'de interesse
    Do While (Selection.Find.Execute = True)
        'Define a vari�vel que guardar� em mem�ria o nome das pessoas, em cada passagem
        Dim nome As String
        'Define a vari�vel que guardar� em mem�ria o valor de cr�dito, em cada passagem
        Dim valor As Currency
        
        'Bloco IF -> controla o fluxo condicional. Somente entra nesse bloco se a
        'condi��o entre par�nteses for satisfeita, ou seja, somente se o valor do
        'cr�dito for superior ao valor determinado nessa linha de c�digo
        If (getValorCredito(Selection.Text) > valorLimite) Then
            'Atribui � vari�vel nome o nome de cada pessoa encontrada
            'Chama a fun��o getNome para retornar esse nome
            nome = getNome(Selection.Text)
            'Atribui � vari�vel valor o valor do cr�dito de cada pessoa encontrada
            'Chama a fun��o getValorCredito para retornar esse valor
            valor = getValorCredito(Selection.Text)
            'Debug.Print getValorCredito(Selection.Text)
            'Debug.Print nome
            
            'Ativa (leva o foco do cursor para) o documento de destino (o que est� sendo
            'escrito)
            doc.Activate
            'Guarda nessa vari�vel tempor�ria, o conte�do da regi�o atual do documento
            'de destino
            Set regiao = doc.Range
            
            'Prepara o cursor para a inser��o (leva-o para o final da regi�o)
            regiao.Collapse wdCollapseEnd
            
            'Atribui a string de nome + tabula��o+ valor + marca de final de par�grafo
            '� regi�o do texto do documento de destino, na posi��o onde deve ser escrita
            regiao.Text = nome & vbTab & FormatCurrency(valor, 2) & vbCrLf
            
            'Retira o negrito que foi posto no cabecalho
            regiao.Bold = False
            
            'Prepara o cursor para a pr�xima itera��o
            regiao.Collapse wdCollapseEnd

        End If
        'Retorna o foco do cursor para o documento original, para que o loop se reinicie
        'da posi��o de onde parou (onde localizou a mais recente inst�ncia da express�o
        'de busca
        docOrig.Activate
    Loop  ' Final do bloco de itera��o (looping)
    
    doc.Activate
    doc.Range.Copy
    
    'Para o funcionamento das linhas de c�digo abaixo, deve-se, antes, incluir a refer�ncia
    '� biblioteca do Excel. Para isso, abra o menu ferramenta->refer�ncias e selecione a
    'caixa de sele��o que contiver o nome Microsoft Excel .. Object Library ou algo parecido
    
    'Define a vari�vel que guardar� o objeto de documento do Excel
    Dim planExcel As Excel.Workbook
    
    'Fecha todos os documentos abertos do Excel
    Dim contaDoc As Integer
    contaDoc = 1
    Do While (Excel.Workbooks.Count > 0)
        Excel.Workbooks(contaDoc).Close SaveChanges:=False
        contaDoc = contaDoc + 1
    Loop
        
    'Atribui � vari�vel acima um novo documento do Excel (Workbook do Excel)
    Set planExcel = Excel.Workbooks.Add
    
    'Ativa essa planilha
    planExcel.Activate
    
    'Cola os dados que est�o na �rea de transfer�ncia na planilha 1 do
    'documento do Excel rec�m aberto
    planExcel.Worksheets(1).Paste
    
    'Ajusta a largura automaticamente ao conte�do (autofit)
    planExcel.Worksheets(1).Range("A:B").Columns.AutoFit
    
    'Aplica o tipo Moeda (Currency) � coluna de valores (coluna B)
    planExcel.Worksheets(1).Columns("B:B").NumberFormat = "$ #,##0.00"
    
    'Salva na pasta documentos
    'planExcel.Save
    
    planExcel.ActiveSheet.Visible = True
    
    planExcel.Application.DisplayFullScreen = True
        
    AppActivate planExcel.Application.Caption
    
    planExcel.Application.DisplayFullScreen = False
    
    'Salva na pasta que voc� definir (pode ser na pasta atual onde se situa este arquivo)
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
    
    'Fecha tanto a planilha quanto o documento de destino Word (texto tempor�rio com os dados
    'do documento Word exportados para a planilha do Excel)
    If (Not planExcel Is Nothing) Then
        'Fecha a planilha Excel que cont�m os dados exportados
        planExcel.Close SaveChanges:=False
        'Fecha o documento Word
        doc.Close SaveChanges:=False
    End If
        
End Sub 'Final do bloco do procedimento SUB

'Func�o destinada a obter o nome da pessoa, dentro da string (cadeia de caracteres)
'completa da express�o encontrada (ou seja, no caso do exemplo, procura o nome
'dentro da express�o "A pessoa Jos� Pereira tem um cr�dito de R$ 77474,01." e
'retorna "Jos� Pereira"
Function getNome(texto) As String
    'defini��o, em mem�ria, da vari�vel que encontrar� a posi��o da palavra "pessoa"
    Dim posPessoa As Integer
    'defini��o, em mem�ria, da vari�vel que encontrar� a posi��o da palavra "tem"
    Dim posTem As Integer
    'localiza, dentro da express�o completa, a palavra pessoa e guarda sua posi��o
    posPessoa = InStr(1, texto, "pessoa")
    'localiza, dentro da express�o completa, a palavra tem e guarda sua posi��o
    posTem = InStr(1, texto, "tem")
    'Defini��o da vari�vel que guardar� a diferen�a entre as posi��es
    'das palavras tem e pessoa
    Dim diferenca As Integer
    'Calcula a diferen�a entre as palavras tem e pessoa
    diferenca = posTem - posPessoa
    'Obtem o nome somente da pessoa em cada itera��o
    getNome = Mid(texto, posPessoa + 7, diferenca - 7)
End Function

'Fun��o para obter o valor de cr�dito em cada express�o encontrada
Function getValorCredito(texto) As Currency
    'Defini��o da posi��o onde for encontrado o subtexto (substring) "R$"
    Dim posCifrao As Integer
    'Guarda a posi��o encontrada do R$ dentro da express�o localizada
    posCifrao = InStr(1, texto, "R$")
    'Guarda, na vari�vel de retorno da fun��o, o valor de cr�dito encontrado
    getValorCredito = Mid(texto, posCifrao + 3, Len(texto) - posCifrao - 3)
    'Debug.Print getValorCredito

End Function

