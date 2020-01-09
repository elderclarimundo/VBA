Sub EnviarEmail()

    Dim OutProg As Object
    Dim OutMail As Object
    
    'Criando objetos do Microsoft Office Outlook (chamando recursos do Outlook)
    Set OutProg = CreateObject("Outlook.Application")
    Set OutMail = OutProg.CreateItem(0)
    
    With OutMail
        .to = "vcodsantos@timbrasil.com.br"
        .cc = ""
        .BCC = ""
        .Subject = "TABULADOR ANATEL | " & Format(Format(Date, "mmmm"), ">") & " - " & Format(Date, "YYYY")
        .HTMLBody = "HTML Content" 'Padrão Formatação HTML
        .Body = "IDs Anatel tabuladas com sucesso!"
        .Send
    End With
    
    'Esvaziar da memória os objetos criados
    Set OutMail = Nothing
    Set OutProg = Nothing

End Sub