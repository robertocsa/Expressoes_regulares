Attribute VB_Name = "Módulo_macros_diversas"
Sub CopiaTabelas()
    'Fonte em inglês obtido a partir do portal Tips.Net
    'https://wordribbon.tips.net/T013338_Copying_All_Tables_to_a_New_Document.html
    
    'Definições dos objeto documento
    Dim docOrigem As Document
    Dim docDestino As Document
    
    'Definição dos objetos tabela
    Dim tabelaAtual As Table
    
    'Definição do objeto região
    Dim regiao As Range
    
    'Definição do objeto contador de tabelas
    Dim contaTabelas As Integer

    'Atribuição dos documentos Word que serão utilizados
    'Atribuição do documento de origem (o texto original)
    Set docOrigem = ActiveDocument
    'Atribuição do documento de destino (o texto onde ficarão as tabelas)
    Set docDestino = Documents.Add

    'Looping de iteração de busca das tabelas
    For Each tabela In docOrigem.Tables
        'Incremento da contagem das tabelas
        contaTabelas = contaTabelas + 1
        
        'Atribuição da região onde será colada a tabela copiada
        Set regiao = docDestino.Range
        
        'Leva o cursor de inserção para o final da região
        regiao.Collapse wdCollapseEnd
        
        'As duas próximas linhas escrevem o título da tabela atual (em cada passagem)
        'regiao.Text = vbCrLf & "Tabela" & contaTabelas & vbCrLf
        'regiao.Collapse wdCollapseEnd
        
        'Copia a tabela de origem, formatada, para o final do texto do documento de
        'destino
        regiao.FormattedText = tabela.Range.FormattedText
        
        'Leva o cursor do documento de destino para o final do texto
        'preparando-o para a próxima iteração
        regiao.Collapse wdCollapseEnd
    Next
End Sub


Sub MacroSubstParags()
Attribute MacroSubstParags.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro utilizada para retirar as marcas de parágrafo excessivas.
'
'
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    
End Sub

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "código*(R$*,[0-9]{2})"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.Find.Execute
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
    Dim i As Long
    Selection.Find.ClearFormatting
    'Selection.Find.Style = ActiveDocument.Styles("")
    Selection.Find.Replacement.ClearFormatting
    'Selection.Find.Replacement.Style = ActiveDocument.Styles("")
    'Selection.Find.Replacement.Highlight = True
    With Selection.Find
        .Text = "O código*R$ [0-9.,]{4;20}"
        '.Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        '.Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Do While Selection.Find.Execute = True
        i = i + 1
        If (valorFaturamento(Selection.Text) > 100000#) Then
            'Debug.Print i & " " & Selection.Text
            Selection.Select
            Selection.Font.Bold = False
        End If
    Loop
End Sub
Function valorFaturamento(texto) As Currency
    posicaoCifrao = InStr(texto, "R$ ")
    If posicaoCifrao > 0 Then
        strValor = Mid(texto, posicaoCifrao + 3)
    End If
    valorFaturamento = Int(Trim(strValor))
    'Debug.Print valorFaturamento
    
End Function

Sub testeAleatorio()
    Debug.Print Int(Rnd(10000) * 10000)
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' Macro3 Macro
'
'
    Dim doc As Document
    Set doc = ActiveDocument
    Dim regiao As Range
    
    Dim parag As Paragraph
    
    For Each parag In doc.Paragraphs
        'Set regiao = doc.Range(parag.Range.Start, parag.Range.End - 1)
        Set regiao = parag.Range
        regiao.InsertAfter (nome())
        Debug.Print Right(regiao.Text, 20)
    Next
    
    'Selection.MoveDown Unit:=wdParagraph, Count:=2, Extend:=wdExtend
End Sub

Function nome() As String
    Dim lista1(2) As String
    Dim aleat1 As Integer
    lista1(0) = "Teste0"
    lista1(1) = "Teste1"
    lista1(2) = "Teste2"
    aleat1 = Int(Rnd(3) * 3)
    nome = lista1(aleat1)
    'Debug.Print lista1(aleat1), aleat1
    
End Function

Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' Macro4 Macro
'
'
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    Selection.HomeKey Unit:=wdStory
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro5"
'
' Macro5 Macro
'
'
   Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "A pessoa*R$ [0-9.,]{4;10}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
    Selection.Find.Execute
End Sub
Sub CriaNovoTextoFicticio()
Attribute CriaNovoTextoFicticio.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.CriaNovoTextoFicticio"
'
' CriaNovoTextoFicticio Macro
'
'
    Dim doc As Document
    Set doc = Documents.Add(Template:="Normal", NewTemplate:=False, DocumentType:=0)
    doc.Activate
    SendKeys "=rand{(}100,4{)}" & vbCr
    SendKeys "=lorem{(}100,4{)}" & vbCr
End Sub
