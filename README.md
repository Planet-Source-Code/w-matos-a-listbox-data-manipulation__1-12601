<div align="center">

## A ListBox Data Manipulation


</div>

### Description

It allows you to move items from one listbox control to another, remove items, or move item positioning within the control. (FIELD LISTBOX DATA MOVE MANAGE ENTRY ITEM INDEX)
 
### More Info
 
one or two listbox controls (or cbo boxes with some tweaking)

When the commands are called it looks at what is currently selected in the control. If nothing is selected... it does nothing.

Nothing, simply acts upon the controls


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[W\. Matos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/w-matos.md)
**Level**          |Intermediate
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/w-matos-a-listbox-data-manipulation__1-12601/archive/master.zip)





### Source Code

```
Option Explicit
'-----------------
' Mod Name FieldProcessing
' Author: W. Matos
' Date: November 07, 2000
' Description: This module provides a series of commands that acts upon
' a set of list boxes (you can change the code to act upon both
' list boxes and combo boxes by declaring the objects as
' 'object' and not ListBox)
' This module lets you:
' 1) Add a field from a source object to a destination object
' 2) Add all fields from a source object to a destination object
' 3) Remove a field from a source object to a destination object
' 4) Move a field up in the object.
' 5) Move a field down in the object.
'
'
' Comment: I understand the simplicity of this set of procedures. However,
' I had never taken the time to actually create this. Since creating
' this module, creating the forms has been greatly simplified.
'
'
' Use: Here is a sample set of code on how to use:
'
' To add a field
'Private Sub cmdAddSummaryField_Click()
' AddField Me.lstAvailFlds, Me.lstSummaryFields
'End Sub
'
' To Move a field down
'Private Sub cmdMoveDownSummary_Click()
' MoveFldDown lstSummaryFields
'End Sub
'
' To move field up:
'Private Sub cmdMoveUpSummary_Click()
' MoveFldUp lstSummaryFields
'End Sub
'
' to Remove a field
'Private Sub cmdRemoveSummary_Click()
' RemoveField lstSummaryFields
'End Sub
'
' To add all fields:
' Private Sub cmdRemoveAllSummary_Click()
' AddAllFields lstAvailFlds, lstSummaryFields
' End Sub
'
' To remove all fields:
' Just call lstsummaryfields.clear
'--------------------------
Public Sub AddAllFields(lstSource As Object, lstDest As Object)
  Dim x As Integer
  lstDest.Clear
  For x = 0 To lstSource.ListCount - 1
    lstDest.AddItem lstSource.List(x)
  Next x
End Sub
Public Sub AddField(Src As Object, Dest As Object)
  Dim x As Integer
  If Src.ListIndex < 0 Then Exit Sub
  If Src.SelCount > 1 Then
    For x = 0 To Src.ListCount - 1
      If Src.Selected(x) Then Dest.AddItem Src.List(x)
    Next x
  Else
    Dest.AddItem Src.List(Src.ListIndex)
  End If
End Sub
Public Sub RemoveField(Src As Object)
  Dim x As Integer
  If Src.ListIndex < 0 Then Exit Sub
  If Src.ListCount < 1 Then Exit Sub
  If Src.SelCount > 1 Then
restart:
    For x = 0 To Src.ListCount - 1
      If Src.Selected(x) Then
        Src.RemoveItem x
        GoTo restart
      End If
    Next x
  Else
    Src.RemoveItem Src.ListIndex
  End If
End Sub
Public Sub MoveFldUp(lb As Object)
  Dim tmpField As String
  Dim i As Integer
  i = lb.ListIndex
  If lb.ListCount < 1 Then Exit Sub
  If i > 0 And i < lb.ListCount Then
    tmpField = lb.List(i - 1)
    lb.List(i - 1) = lb.List(i)
    lb.List(i) = tmpField
    lb.ListIndex = i - 1
    lb.Selected(i - 1) = True
    lb.Selected(i) = False
  End If
End Sub
Public Sub MoveFldDown(lb As Object)
  Dim tmpField As String
  Dim i As Integer
  i = lb.ListIndex
  If lb.ListCount < 1 Then Exit Sub
  If i > -1 And i < lb.ListCount - 1 Then
    tmpField = lb.List(i + 1)
    lb.List(i + 1) = lb.List(i)
    lb.List(i) = tmpField
    lb.ListIndex = i + 1
    lb.Selected(i + 1) = True
    lb.Selected(i) = False
  End If
End Sub
```

