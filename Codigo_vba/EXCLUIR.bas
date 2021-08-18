Attribute VB_Name = "Módulo1"
'codigo para excluir planilhas atravez de uma dada determinada ao abrir a planilha
Private Sub Workbook_Open()
'definindo a data para a exclusão
If DateDiff("d", Date, "01/08/2021") < 0 Then
'desabilitando o display
Application.DisplayAlerts = False
Dim BK As Integer, A As Integer
For BK = Sheets.Count To 0 Step -1
If Sheets.Count = 1 Then
ActiveWorkbook.Protect ("soldado_985")
MsgBox "total de planilhas excluidas: " & A & " expiraram a: " & DateDiff("d", "01/08/2021", Date) & " dias", vbInformation, "EXCLUSÃO DE PLANILHAS"
Exit Sub
End If
A = A + 1
Sheets(BK).Delete
Next BK
Application.DisplayAlerts = True
Else
Exit Sub
End If
End Sub

'Escrito por wdeybsonjunho@gmail.com 16/08/2021
 'contatos (62) 9 9836-3956
