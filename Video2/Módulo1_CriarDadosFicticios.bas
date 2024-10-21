Attribute VB_Name = "Módulo1_CriarDadosFicticios"
'###################################################################################
'Cria um documento inicialmente vazio e aplica os comandos de rand e de lorem ipsum do word
'para criar textos fictícios para testes.
Public Sub CriaNovoTextoFicticioEmDocWord()
'
' CriaNovoTextoFicticio Macro
'
'
    Dim doc As Document
    Set doc = Documents.Add(Template:="Normal", NewTemplate:=False, DocumentType:=0)
    doc.Activate
    Dim i As Integer
    Do While i < 50
        SendKeys "=rand{(}2,4{)}" & vbCr
        SendKeys "=lorem{(}2,4{)}" & vbCr
        i = i + 1
    Loop
    'Call criarDadosFicticios
    Call AddReferencesWord(doc)
    
End Sub


'Função destinada a criar as frases contendo os dados fictícios
Public Sub criarDadosFicticios()
    'A expressão de interesse terá o padrão (sem as aspas):
    '" A pessoa <nome> tem um crédito de R$ <valor>."
    'Define a lista (ou array) que guardará os nomes próprios fictícios
    Dim listaNomes(9) As String
    'Atribui nomes a cada item da lista (de 0 a 9, ou seja, 10 nomes)
    
    listaNomes(0) = "Roberto"
    listaNomes(1) = "Andrea"
    listaNomes(2) = "José"
    listaNomes(3) = "Nice"
    listaNomes(4) = "Carlos"
    listaNomes(5) = "Mariana"
    listaNomes(6) = "Felipe"
    listaNomes(7) = "Ada"
    listaNomes(8) = "Joana"
    listaNomes(9) = "Nicole"
            
    'Define a lista (ou array) que guardará os sobrenomes fictícios
    Dim listaSobreNomes(9) As String
    'Atribui sobrenomes a cada item da lista (de 0 a 9, ou seja, 10 sobrenomes)
    listaSobreNomes(0) = "Santos"
    listaSobreNomes(1) = "Silva"
    listaSobreNomes(2) = "Campos"
    listaSobreNomes(3) = "Pereira"
    listaSobreNomes(4) = "Gonçalves"
    listaSobreNomes(5) = "Nunes"
    listaSobreNomes(6) = "Alencar"
    listaSobreNomes(7) = "Vargas"
    listaSobreNomes(8) = "Oliveira"
    listaSobreNomes(9) = "Moreira"
    
    'Define o parágrafo de cada passagem da iteração (do loop)
    Dim parg As Paragraph
    'Define a variável que trabalhará com a coleção de parágrafos do documento de origem
    Dim parags As Paragraphs
    
    'Atribui à variável parags a coleção de parágrafos do documento atual (original)
    Set parags = ActiveDocument.Paragraphs
    
    'Define a variável que guardará a frase final a ser inserida no texto, ou seja, a
    'frase padrão, modificada em cada iteração para adicionar um nome e um valor
    'aleatórios e fictícios
    Dim fraseFinalParagrafo As String
    
    'Atribui à variável seguinte a frase padrão a ser trabalhada
    fraseFinalParagrafo = " A pessoa <nome> tem um crédito de R$ <valor>."
    
    'Define o documento original
    Dim doc As Document
    'Atribui o documento ativo à essa variável
    Set doc = ActiveDocument
    
    'Define a região do documento original onde será inserida a frase padrão após modificada
    Dim regiao As Range
    
    'Início do bloco de iterações (looping) - em cada passagem, trabalha com um dos
    'parágrafos do documento original
    For Each parg In parags
        'Debug.Print Right(parg.Range.Text, 20)
        'Atribui à variável temporária regiao a região do parágrafo em cada iteração
        Set regiao = parg.Range
        'Reduz 1 caracter dessa região, para excluir a marca de parágrafo no final
        Set regiao = doc.Range(parg.Range.Start, parg.Range.End - 1)
        'chama a função que modifica a frase padrão, para incluir nomes e valores
        frase = montaFrase(fraseFinalParagrafo, nome(listaNomes, listaSobreNomes), valores())
        'Debug.Print frase
        'Leva o cursor para o final do parágrafo atual
        regiao.Collapse wdCollapseEnd
        'Insere a frase padrão já modificada
        regiao.Text = frase
        'Realça, em amarelo, o texto recém inserido
        regiao.HighlightColorIndex = wdYellow
    Next
    
    'Debug.Print nome(listaNomes, listaSobreNomes)
    'Debug.Print valores()

End Sub

'Função para montar a frase padrão
Function montaFrase(fraseFinalParagrafo, nome, valor)
    'Debug.Print nome
    'Debug.Print fraseFinalParagrafo
    'Troca o texto "<nome>" pelo nome guardado na variável nome
    montaFrase = Replace(fraseFinalParagrafo, "<nome>", nome)
    'Debug.Print montaFrase
    'Troca o texto "<valor>" pelo valor guardado na variável valor
    montaFrase = Replace(montaFrase, "<valor>", valor)
    'Debug.Print montaFrase
End Function

'Função destinada a gerar nome+espaço+sobrenome aleatório, dentre aqueles contidos nas
'listas de nome e de sobrenome
Function nome(listaNomes, listaSobreNomes) As String
    'Define a variável que guardará o valor aleatório para o índice de nomes
    Dim numAleat1 As Byte
    'Define a variável que guardará o valor aleatório para o índice de sobrenomes
    Dim numAleat2 As Byte
    'Nas duas linhas abaixo, são gerados índices de 0 a 9
    numAleat1 = Int(Rnd(10) * 10)
    numAleat2 = Int(Rnd(10) * 10)
       
    'Atribui à variável de retorno nome a string nome+espaço+sobrenome
    nome = listaNomes(numAleat1) & " " & listaSobreNomes(numAleat2)
End Function

'Função para retornar um valor fictício entre R$0,00 a R$ 999999,99
Function valores() As Currency
    'Atribui à variável de retorno da função o valor gerado aleatoriamente
    valores = Rnd(100000) * 100000
End Function

