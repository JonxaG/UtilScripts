Dim outobj, mailobj
Dim strFileText
Dim objFileToRead

Set outobj = CreateObject("Outlook.Application")
Set mailobj = outobj.CreateItem(0)

With mailobj
    .To = "mailer@mailer.com"
    .CC = "copia@mailer.com"
    .Subject = "Server down"
    .Body = "text"
    .Attachments.Add "RUTA FICHERO ADJUNTO"
    .Display
End With

mailobj.SaveAs "RUTA BORRADOR CORREO .msg"

mailobj.Close olDiscard

Set outobj = Nothing
Set mailobj = Nothing
