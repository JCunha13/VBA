Attribute VB_Name = "Módulo1"
Sub ImportCSVToExcel()
    Dim ws As Worksheet
    Dim csvFilePath As String
    Dim csvLine As String
    Dim csvData() As String
    Dim rowNum As Long
    Dim colNum As Long
    Dim fNum As Integer
    
    ' Definir a planilha de destino
    Set ws = ThisWorkbook.Sheets("Planilha1")
    
    ' Caminho para o arquivo CSV
    csvFilePath = "C:\caminho\para\seu\arquivo.csv"
    
    ' Abrir o arquivo CSV para leitura
    fNum = FreeFile
    Open csvFilePath For Input As #fNum
    
    ' Inicializar a linha na planilha do Excel
    rowNum = 1
    
    ' Ler o arquivo CSV linha por linha
    Do While Not EOF(fNum)
        Line Input #fNum, csvLine
        csvData = Split(csvLine, ",")
        
        ' Escrever os dados na planilha do Excel
        For colNum = LBound(csvData) To UBound(csvData)
            ws.Cells(rowNum, colNum + 1).Value = csvData(colNum)
        Next colNum
        
        rowNum = rowNum + 1
    Loop
    
    ' Fechar o arquivo CSV
    Close #fNum
    
    MsgBox "Importação do CSV concluída!"
End Sub

