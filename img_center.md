# Guide On How To Center All Images In Word

Press Alt + F11 to open the VBA editor.

Go to Insert > Module and paste the following code:
```vba
Sub CenterAllImages()
    Dim img As InlineShape
    For Each img In ActiveDocument.InlineShapes
        img.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next img
End Sub
```
Close the VBA editor and return to Word.

Press Alt + F8, select CenterAllImages, and click Run.

This macro will loop through all inline images in the document and center them.
