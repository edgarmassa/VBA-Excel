'DECLARA A VARIÁVEL
Dim objShape As Shape

'LOOP ATRAVÉS DE TODOS SHAPES(IMAGENS) DENTRO DO SEU SHEET
For Each objShape In Sheets("nome da sua sheet").Shapes

    'VOCÊ PODE COLOCAR UMA CONDIÇÃO SE QUISER COM BASE NO NOME DO SHAPE LÁ NA SHEET
    'If objShape.Name <> "Icone" Then
    '    objShape.Delete
    'End If

    'OU SIMPLESMENTE DELETAR
    objShape.Delete
Next