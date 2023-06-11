Attribute VB_Name = "NewMacros"
Sub spoiler_whole_line()
Attribute spoiler_whole_line.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.spoiler_whole_line"
'
' Adds a spoiler on the line the cursor is on and moves the cursor to the start of the line
'
'
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub spoiler_after_cursor()
Attribute spoiler_after_cursor.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.spoiler_after_cursor"
'
' Adds a spoiler to the text after the cursor on the line. 
'
'
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub spoiler_multiple_lines()
'
' Adds a spoiler on 3 lines (including the current one). 
' Change Count:=3 to Count:=N to make it work for N lines.
'   
    Selection.HomeKey Unit:=wdLine
    Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub