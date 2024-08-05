Attribute VB_Name = "Módulo1"
Sub ExportarCSVParaPDF()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim csvPath As String
    Dim pdfPath As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerRange As Range
    Dim dataRange As Range

    ' Defina o caminho do arquivo CSV e o caminho do arquivo PDF de saída
    csvPath = "C:\Users\Juarez Cunha\Documents\Data\amazon_reviews.csv"
    pdfPath = "C:\Users\Juarez Cunha\Documents\Data\amazon_reviews.pdf"
    
    ' Abrir o arquivo CSV
    Set wb = Workbooks.Open(Filename:=csvPath)
    Set ws = wb.Sheets(1)
    
    ' Identificar a última linha e coluna com dados
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Definir o intervalo dos cabeçalhos e dos dados
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    
    ' Aplicar formatações aos cabeçalhos
    With headerRange
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255) ' Cor de fundo azul claro
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
    ' Aplicar formatações aos dados
    With dataRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
    ' Autoajustar a largura das colunas
    ws.Columns.AutoFit
    
    ' Salvar a planilha como PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath
    
    ' Fechar o arquivo CSV sem salvar as alterações
    wb.Close SaveChanges:=False
    
    ' Informar o usuário que o processo foi concluído
    MsgBox "Arquivo CSV formatado e exportado para PDF com sucesso!", vbInformation
End Sub

