Const wdReplaceAll = 2
Set objWord =  GetObject(,"Word.Application")
objWord.Visible = True

Set objDoc = objWord.ActiveDocument
Set objSelection = objWord.Selection

objSelection.Find.Text = WScript.Arguments(0)
objSelection.Find.Forward = TRUE
objSelection.Find.MatchWholeWord = TRUE
objSelection.Find.Replacement.Text = WScript.Arguments(1)
objSelection.Find.ClearFormatting()
objSelection.Find.Execute ,,,,,,,,,,wdReplaceAll