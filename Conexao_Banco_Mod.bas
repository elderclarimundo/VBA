Option Explicit

    Public cn          As New ADODB.Connection
    Public rs          As New ADODB.Recordset
    Public arq         As String

Public Sub CriaConexaoBanco()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Criar conexão com o banco
    Set cn = New ADODB.Connection
    
    'Abrir conexão
    cn.Open ConexaoDB
    
    'Busca dados no banco
    BuscaDadosBanco_Mod.AtualizaDadosBanco
    
    'Adiciona os IDs na lista
    ID_Anatel_Mod.ID_Anatel
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

Public Sub EncerraConexao()

    On Error Resume Next
    'Fechar conexão com o banco
    cn.Close

End Sub

Function ConexaoDB()

    Dim arq         As String

    arq = "\\dtrj56312\AJUSTE-SUPER\TABULADOR\Transbordo_Anatel.accdb"

    ConexaoDB = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & arq & ";"

End Function
