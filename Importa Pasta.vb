Option Private Module

'Esta subrotino é a responsável por percorrer a pasta e executar a subrotina "Importar Formulários" para cada arquivo
'contido dentro daquela pasta alvo
Public Sub ImportaPasta(path As String)

    Dim wb As Workbook
    Dim myPath As String
    Dim myFile As String
    Dim myExtension As String
    Dim FldrPicker As FileDialog

    'Aumento da performance da macro
    Application.ScreenUpdating = False


    myPath = path & Application.PathSeparator

    If myPath = "" Then GoTo ResetSettings

    'Target File Extension (must include wildcard "*")
    myExtension = "*.xls*"

    'Target Path with Ending Extention
    myFile = Dir(myPath & myExtension)

    'Loop through each Excel file in folder
    Do While myFile <> ""

        'Ensure Workbook has opened before moving on to next line of code
        'DoEvents
        
        Call ImportaFormulario(myPath & myFile)
    
        'Get next file name
        myFile = Dir
    Loop
    
    'Message Box when tasks are completed
    MsgBox "Importação finalizada com sucesso!"
    
ResetSettings:
    'Reset Macro Optimization Setting
    Application.ScreenUpdating = True
        
End Sub

'Função que busca o nome e a quantidade de arquivos dentro de uma pasta
Public Function listararquivos(myPath As String)
    Dim xFSO As Object
    Dim xFolder As Object
    Dim xFile As Object
    Dim xFiDialog As FileDialog
    Dim Text As String
    Dim qtdArquivos As Integer
    
    
    Text = ""
    qtdArquivos = 0
    
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    Set xFolder = xFSO.GetFolder(myPath)
    
    For Each xFile In xFolder.Files
        Text = Text & xFile.Name & vbLf
        qtdArquivos = qtdArquivos + 1
    Next
    
    listararquivos = " (" & CStr(qtdArquivos) & " arquivos encontrados" & ")" & vbLf & Text
End Function
