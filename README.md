<div align="center">

## Add only unique items to a listbox \(no duplicates added\)\. Fast and short without any loops\.


</div>

### Description

Use this method to avoid adding an item to a ListBox that already exists.

It's a lot faster and shorter than submissions that uses loops etc.
 
### More Info
 
StringToAdd = the string to add (if not already exists)

lst = your ListBox


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Fredrik Schultz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fredrik-schultz.md)
**Level**          |Intermediate
**User Rating**    |4.8 (53 globes from 11 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fredrik-schultz-add-only-unique-items-to-a-listbox-no-duplicates-added-fast-and-short-with__1-11678/archive/master.zip)





### Source Code

```
Private Sub AddUnique(StringToAdd As String, lst As ListBox)
  lst.Text = StringToAdd
  If lst.ListIndex = -1 Then
    'it does not exist, so add it..
    lst.AddItem StringToAdd
  End If
End Sub
```

