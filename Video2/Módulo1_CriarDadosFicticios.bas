Attribute VB_Name = "M�dulo1_CriarDadosFicticios"
'###################################################################################
'Cria um documento inicialmente vazio e aplica os comandos de rand e de lorem ipsum do word
'para criar textos fict�cios para testes.
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


'Fun��o destinada a criar as frases contendo os dados fict�cios
Public Sub criarDadosFicticios()
    'A express�o de interesse ter� o padr�o (sem as aspas):
    '" A pessoa <nome> tem um cr�dito de R$ <valor>."
    'Define a lista (ou array) que guardar� os nomes pr�prios fict�cios
    Dim listaNomes(9) As String
    'Atribui nomes a cada item da lista (de 0 a 9, ou seja, 10 nomes)
    
    listaNomes(0) = "Roberto"
    listaNomes(1) = "Andrea"
    listaNomes(2) = "Jos�"
    listaNomes(3) = "Nice"
    listaNomes(4) = "Carlos"
    listaNomes(5) = "Mariana"
    listaNomes(6) = "Felipe"
    listaNomes(7) = "Ada"
    listaNomes(8) = "Joana"
    listaNomes(9) = "Nicole"
            
    'Define a lista (ou array) que guardar� os sobrenomes fict�cios
    Dim listaSobreNomes(9) As String
    'Atribui sobrenomes a cada item da lista (de 0 a 9, ou seja, 10 sobrenomes)
    listaSobreNomes(0) = "Santos"
    listaSobreNomes(1) = "Silva"
    listaSobreNomes(2) = "Campos"
    listaSobreNomes(3) = "Pereira"
    listaSobreNomes(4) = "Gon�alves"
    listaSobreNomes(5) = "Nunes"
    listaSobreNomes(6) = "Alencar"
    listaSobreNomes(7) = "Vargas"
    listaSobreNomes(8) = "Oliveira"
    listaSobreNomes(9) = "Moreira"
    
    'Define o par�grafo de cada passagem da itera��o (do loop)
    Dim parg As Paragraph
    'Define a vari�vel que trabalhar� com a cole��o de par�grafos do documento de origem
    Dim parags As Paragraphs
    
    'Atribui � vari�vel parags a cole��o de par�grafos do documento atual (original)
    Set parags = ActiveDocument.Paragraphs
    
    'Define a vari�vel que guardar� a frase final a ser inserida no texto, ou seja, a
    'frase padr�o, modificada em cada itera��o para adicionar um nome e um valor
    'aleat�rios e fict�cios
    Dim fraseFinalParagrafo As String
    
    'Atribui � vari�vel seguinte a frase padr�o a ser trabalhada
    fraseFinalParagrafo = " A pessoa <nome> tem um cr�dito de R$ <valor>."
    
    'Define o documento original
    Dim doc As Document
    'Atribui o documento ativo � essa vari�vel
    Set doc = ActiveDocument
    
    'Define a regi�o do documento original onde ser� inserida a frase padr�o ap�s modificada
    Dim regiao As Range
    
    'In�cio do bloco de itera��es (looping) - em cada passagem, trabalha com um dos
    'par�grafos do documento original
    For Each parg In parags
        'Debug.Print Right(parg.Range.Text, 20)
        'Atribui � vari�vel tempor�ria regiao a regi�o do par�grafo em cada itera��o
        Set regiao = parg.Range
        'Reduz 1 caracter dessa regi�o, para excluir a marca de par�grafo no final
        Set regiao = doc.Range(parg.Range.Start, parg.Range.End - 1)
        'chama a fun��o que modifica a frase padr�o, para incluir nomes e valores
        frase = montaFrase(fraseFinalParagrafo, nome(listaNomes, listaSobreNomes), valores())
        'Debug.Print frase
        'Leva o cursor para o final do par�grafo atual
        regiao.Collapse wdCollapseEnd
        'Insere a frase padr�o j� modificada
        regiao.Text = frase
        'Real�a, em amarelo, o texto rec�m inserido
        regiao.HighlightColorIndex = wdYellow
    Next
    
    'Debug.Print nome(listaNomes, listaSobreNomes)
    'Debug.Print valores()

End Sub

'Fun��o para montar a frase padr�o
Function montaFrase(fraseFinalParagrafo, nome, valor)
    'Debug.Print nome
    'Debug.Print fraseFinalParagrafo
    'Troca o texto "<nome>" pelo nome guardado na vari�vel nome
    montaFrase = Replace(fraseFinalParagrafo, "<nome>", nome)
    'Debug.Print montaFrase
    'Troca o texto "<valor>" pelo valor guardado na vari�vel valor
    montaFrase = Replace(montaFrase, "<valor>", valor)
    'Debug.Print montaFrase
End Function

'Fun��o destinada a gerar nome+espa�o+sobrenome aleat�rio, dentre aqueles contidos nas
'listas de nome e de sobrenome
Function nome(listaNomes, listaSobreNomes) As String
    'Define a vari�vel que guardar� o valor aleat�rio para o �ndice de nomes
    Dim numAleat1 As Byte
    'Define a vari�vel que guardar� o valor aleat�rio para o �ndice de sobrenomes
    Dim numAleat2 As Byte
    'Nas duas linhas abaixo, s�o gerados �ndices de 0 a 9
    numAleat1 = Int(Rnd(10) * 10)
    numAleat2 = Int(Rnd(10) * 10)
       
    'Atribui � vari�vel de retorno nome a string nome+espa�o+sobrenome
    nome = listaNomes(numAleat1) & " " & listaSobreNomes(numAleat2)
End Function

'Fun��o para retornar um valor fict�cio entre R$0,00 a R$ 999999,99
Function valores() As Currency
    'Atribui � vari�vel de retorno da fun��o o valor gerado aleatoriamente
    valores = Rnd(100000) * 100000
End Function

