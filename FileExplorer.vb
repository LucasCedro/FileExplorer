Option Private Module

'Função que busca por arquivos "Files"
Public Function browseFilePath()
    On Error GoTo err
    Dim FileExplorer As FileDialog
    Set FileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    Dim filePath As String

    'Desativa multiseleção de objetos
    FileExplorer.AllowMultiSelect = False


    With FileExplorer
    
        'Título da Janela
        .Title = "Selecionar arquivo destino..."

        'Filtro de arquivos
        .Filters.Clear
        .Filters.Add "Excel", "*.xls*"
        .Filters.Add "All Files", "*.*"
    
        If .Show = -1 Then 'Se qualquer arquivo foi escolhido
            filePath = .SelectedItems.Item(1)
        Else 'Usuário clicou em "Cancelar"
            filePath = "" 'Define Path como null pois a importação foi cancelada pelo usuário
        End If
    End With
    
    browseFilePath = filePath
    
err:
    Exit Function
End Function

'Função que busca por pastas "Folders"
Public Function browseFolderPath()
    On Error GoTo err
    Dim FileExplorer As FileDialog
    Set FileExplorer = Application.FileDialog(msoFileDialogFolderPicker)

    Dim folderPath As String

    'To allow or disable to multi select
    FileExplorer.AllowMultiSelect = False

    With FileExplorer
        
        'Título da Janela
        .Title = "Selecionar pasta destino..."

        If .Show = -1 Then 'Se qualquer arquivo foi escolhido
            folderPath = .SelectedItems.Item(1)
        Else 'Usuário clicou em "Cancelar"
            MsgBox "Importação Cancelada"
            folderPath = "" 'Define Path como null pois a importação foi cancelada pelo usuário
        End If
    End With
    
    browseFolderPath = folderPath
    
err:
    Exit Function
End Function
