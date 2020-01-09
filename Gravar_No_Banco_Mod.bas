Option Explicit

Sub GravaNobanco()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim WB              As Workbook
    Dim WS1             As Worksheet
    
    Dim QtLn                As Integer
    Dim LnAtiva             As Integer
    Dim A                   As Integer
    
    Dim ProtAnatel          As String
    
    ProtAnatel = Cad_Form.ID_Anatel_ComboBox.Value
    Set WB = ThisWorkbook
    Set WS1 = WB.Sheets("Finalizado")

    QtLn = 0
    QtLn = Application.WorksheetFunction.Match(Cad_Form.ID_Anatel_ComboBox.Value, WS1.Range("P:P"), False)
        
        LnAtiva = 1
        'WS1.Range("A" & A).Select
    
        Dim AltPergunta1        As String
        Dim AltPergunta2        As String
        Dim AltPergunta3        As String
        Dim AltPergunta4        As String
        Dim AltPergunta5        As String
        Dim AltPergunta6        As String
        Dim AltPergunta7        As String
        Dim Feito               As String

        Dim SQL                 As String
        Dim Col                 As Integer
'        Dim cn                  As New ADODB.Connection
        Dim rs                  As New ADODB.Recordset
        
        WS1.Select
        
        'Alimentar variáveis
        AltPergunta1 = WS1.Range("Q" & QtLn).Value
        AltPergunta2 = WS1.Range("R" & QtLn).Value
        AltPergunta3 = WS1.Range("S" & QtLn).Value
        AltPergunta4 = WS1.Range("T" & QtLn).Value
        AltPergunta5 = WS1.Range("U" & QtLn).Value
        AltPergunta6 = WS1.Range("V" & QtLn).Value
        AltPergunta7 = WS1.Range("W" & QtLn).Value
        Feito = WS1.Range("X" & QtLn).Value
        
'        'Atualiza o banco
'        cn.Open ConexaoDB
        
        SQL = "UPDATE Transbordo_Anatel SET"
        SQL = SQL & " Pergunta1 = '" & AltPergunta1 & "',"
        SQL = SQL & " Pergunta2 = '" & AltPergunta2 & "',"
        SQL = SQL & " Pergunta3 = '" & AltPergunta3 & "',"
        SQL = SQL & " Pergunta4 = '" & AltPergunta4 & "',"
        SQL = SQL & " Pergunta5 = '" & AltPergunta5 & "',"
        SQL = SQL & " Pergunta6 = '" & AltPergunta6 & "',"
        SQL = SQL & " Pergunta7 = '" & AltPergunta7 & "',"
        SQL = SQL & " Feito = '" & Feito & "'"
        SQL = SQL & "WHERE "
        SQL = SQL & " FOCUS_NUM_CHAMADO = '" & ProtAnatel & "'"
            
        cn.Execute SQL
    

'
'        On Error GoTo exit_point
'            cn.Execute SQL
'
'            cn.Close
'
'        On Error GoTo 0
'
'    End If
'
'Sair:
'    On Error Resume Next
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'
'    Exit Sub
'
'exit_point:
'    On Error Resume Next
'    cn.Close
'        MsgBox "Erro na abertura do banco"
'        Application.ScreenUpdating = True
'        Resume Sair
        

    cn.Execute SQL
'    cn.Close
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


