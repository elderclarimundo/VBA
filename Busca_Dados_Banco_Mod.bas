
Dim vUltCelSel      As Range
Public Sub AtualizaDadosBanco()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.Sheets("Finalizado").Visible = xlSheetVisible
    ThisWorkbook.Sheets("SUPERVISORES").Visible = xlSheetVisible
    ThisWorkbook.Sheets("BASE").Visible = xlSheetVisible
    
    Set vUltCelSel = ActiveCell
    
    Dim SQL         As String
    Dim cn          As New ADODB.Connection
    Dim rs          As New ADODB.Recordset
    Dim i           As Integer
    Dim Col         As Integer
    Dim arq         As String
    Dim vcol        As Range
    Dim vRng        As Range
    
    Dim WB          As Workbook
    Dim WS1         As Worksheet
    
    Set WB = ThisWorkbook
    Set WS1 = WB.Sheets("BASE")
    
    WS1.Select
    WS1.Range("A1").Select
    
    'Apagar as células utilizadas anteriormente
    WS1.Range("A6").CurrentRegion.ClearContents

    'Iniciar a inserção dos dados
    Col = 1
    
'    'Criar conexão com o banco
'    Set cn = New ADODB.Connection
    
'    'Abrir conexão
'    cn.Open ConexaoDB
    
    'Criar um recordset
    Set rs = New ADODB.Recordset
    
    SQL = RetornaSQL(1)
    
    'Realiza a consulta
    rs.Open SQL, cn
    
    'Verifica se há dados no recordset
    If rs.EOF = True Then
       MsgBox "Você não possui IDs Anatel para tabular."
    End If
    
    'INserir dados na planilha
    WS1.Cells(1, 1).CopyFromRecordset rs
    
    'Fechar Recordset
    rs.Close
    
'    'Fechar conexão com o banco
'    cn.Close
    
    ThisWorkbook.Sheets("Finalizado").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("SUPERVISORES").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("BASE").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    
End Sub

'Function ConexaoDB()
'
'    Dim arq         As String
'
'    arq = "\\dtrj56312\AJUSTE-SUPER\TABULADOR\Transbordo_Anatel.accdb"
'
'    ConexaoDB = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & arq & ";"
'
'End Function


Public Function RetornaSQL(vqQuery As String)

    Dim WB1                 As Workbook
    Dim WS1                 As Worksheet
    Dim URede               As String
    Dim WS As Object
    

    Set WB1 = ThisWorkbook
    Set WS1 = WB1.Sheets("SUPERVISORES")
    Set WS = CreateObject("Wscript.network")
    
    URede = Format(WS.UserName, ">")

    Dim Gestor              As String
    Gestor = WS1.Range("B" & Application.WorksheetFunction.Match(URede, WS1.Range("A:A"), False)).Value

    Select Case vqQuery
        Case 1
            'SQL = "Select * from Transbordo_Anatel WHERE SUPERVISOR = '" & Gestor & "'"
            SQL = "Select * from Transbordo_Anatel WHERE Feito = '0' AND SUPERVISOR = '" & Gestor & "'" & "ORDER BY DATA ASC"
            'SQL = "Select * from Transbordo_Anatel WHERE Feito = '0' And YEAR([Data]) = YEAR(NOW()) AND MONTH([Data]) = MONTH(NOW()) AND SUPERVISOR = '" & Gestor & "'"
            
    End Select
            
        RetornaSQL = SQL

End Function
