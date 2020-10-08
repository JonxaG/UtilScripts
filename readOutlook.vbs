Set objOutlook = CreateObject("Outlook.Application")
 
Set objNamespace = objOutlook.GetNamespace("MAPI")
 
Set objFolder = objNamespace.GetDefaultFolder(6) 'Inbox
 
Set colItems = objFolder.Items
 
Set colFilteredItems = colItems.Restrict("[Unread]=true")
 
For k = colFilteredItems.Count to 1 step -1 
    set objMessage  = colFilteredItems.Item(k)
    objMessage.Unread = False
    'Accion por cada correo
next
