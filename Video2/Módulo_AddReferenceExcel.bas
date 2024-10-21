Attribute VB_Name = "Módulo_AddReferenceExcel"
'As macros abaixo só funcionam para aplicativos Excel
'Servem para incluir referencias automaticamente, via programação VBA
'Mas devem ser utilizadas quando estiver trabalhando com macros VBA do Excel
'Ao final, farei adaptações para versões similares funcionarem no Word

'Sub AddReferences(wbk As Workbook)
    ' Run DebugPrintExistingRefs in the immediate pane, to show guids of existing references
'    AddRef wbk, "{00025E01-0000-0000-C000-000000000046}", "DAO"
'    AddRef wbk, "{00020905-0000-0000-C000-000000000046}", "Word"
'    AddRef wbk, "{91493440-5A91-11CF-8700-00AA0060263B}", "PowerPoint"
'End Sub

'Sub AddRef(wbk As Workbook, sGuid As String, sRefName As String)
'    Dim i As Integer
'    On Error GoTo EH
'    With wbk.VBProject.References
'        For i = 1 To .Count
'            If .Item(i).Name = sRefName Then
'               Exit For
'            End If
'        Next i
'        If i > .Count Then
'           .AddFromGuid sGuid, 0, 0 ' 0,0 should pick the latest version installed on the computer
'        End If
'    End With
'EX: Exit Sub
'EH: MsgBox "Error in 'AddRef'" & vbCrLf & vbCrLf & err.Description
'    Resume EX
'    Resume ' debug code
'End Sub

'Macro para mostrar as GUID IDs de arquivos do Excel (só funciona em aplicativo Excel)
'Public Sub DebugPrintExistingRefs()
'    Dim i As Integer
'    With Application.ThisWorkbook.VBProject.References
'        For i = 1 To .Count
'            Debug.Print "    AddRef wbk, """ & .Item(i).GUID & """, """ & .Item(i).Name & """"
'        Next i
'    End With
'End Sub
'######################################################################################
'Macro para mostrar as GUID IDs de arquivos do Word (adaptação das macros acima, para funcionar em aplicativo Word)
Sub DebugPrintExistingRefsWord()
    Dim i As Integer
    With Application.ActiveDocument.VBProject.References
        For i = 1 To .Count
            Debug.Print "    AddRef doc, """ & .Item(i).GUID & """, """ & .Item(i).Name & """"
        Next i
    End With
End Sub

Public Sub AddReferencesWord(doc As Document)
    ' Run DebugPrintExistingRefs in the immediate pane, to show guids of existing references
    AddRefWord doc, "{00025E01-0000-0000-C000-000000000046}", "DAO"
    'AddRefWord doc, "{00020905-0000-0000-C000-000000000046}", "Word"
    'AddRefWord doc, "{91493440-5A91-11CF-8700-00AA0060263B}", "PowerPoint"
    AddRefWord doc, "{000204EF-0000-0000-C000-000000000046}", "VBA"
    AddRefWord doc, "{00020430-0000-0000-C000-000000000046}", "stdole"
    AddRefWord doc, "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "Office"
    AddRefWord doc, "{00020813-0000-0000-C000-000000000046}", "Excel"
    
End Sub

Sub AddRefWord(doc As Document, sGuid As String, sRefName As String)
    Dim i As Integer
    On Error GoTo EH
    With doc.VBProject.References
        For i = 1 To .Count
            If .Item(i).Name = sRefName Then
               Exit For
            End If
        Next i
        If i > .Count Then
           .AddFromGuid sGuid, 0, 0 ' 0,0 should pick the latest version installed on the computer
        End If
    End With
EX: Exit Sub
EH: MsgBox "Error in 'AddRef'" & vbCrLf & vbCrLf & err.Description
    Resume EX
    Resume ' debug code
End Sub
