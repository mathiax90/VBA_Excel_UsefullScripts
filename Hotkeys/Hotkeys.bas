Attribute VB_Name = "Hotkeys"

'cp 1251

Sub auto_open()
    Call SetOnKeys
End Sub


Sub SetOnKeys()
    Application.OnKey "^+{E}", "MergeUnmergeCells"
    Application.OnKey "^+{C}", "UniqCount"
    Application.OnKey "^+{A}", "AddRow"
    Application.OnKey "^+{Z}", "AddColumn"
    Application.OnKey "^+{V}", "VerticalStrip"
    Application.OnKey "^+{H}", "HorizontalStrip"
    Application.OnKey "^+{G}", "AllBorders"
    Application.OnKey "^+{I}", "Info"
End Sub


Sub Info()
    ' russian
    MsgBox "CTRL+R - Объединить с сохранением текста" & Chr(13) _
    & "CTRL+G - Все граница (Grid)" & Chr(13) _
    & "CTRL+С - Подсчитать уникальные" & Chr(13) _
    & "CTRL+A - Добавить строку" & Chr(13) _
    & "CTRL+Z - Добавить столбец" & Chr(13) _
    & "CTRL+V - Вертикальная зебра" & Chr(13) _
    & "CTRL+H - Горизонтальная зебра" & Chr(13) _
    & "CTRL+I - Данное окно" & Chr(13)

    ' english    '

    ' MsgBox "CTRL+R - Merge Unmerge" & Chr(13) _
    ' & "CTRL+G - All borders" & Chr(13) _
    ' & "CTRL+С - Count unique" & Chr(13) _
    ' & "CTRL+A - Add row" & Chr(13) _
    ' & "CTRL+Z - Add column" & Chr(13) _
    ' & "CTRL+V - Vertical strip" & Chr(13) _
    ' & "CTRL+H - Horizontal strip" & Chr(13) _
    ' & "CTRL+I - This window" & Chr(13)
End Sub

Sub MergeUnmergeCells()
    Dim rng As Range
    Set rng = Selection
    Dim rsltVal As String

    If rng.MergeCells Then
        rng.UnMerge
    Else
        For Each cl In rng.Cells
            rsltVal = Trim(rsltVal) & " " & Trim(cl.Value)
        Next cl
        
        rng.ClearContents
        rng.Merge
        rng.Value = Trim(rsltVal)
        rng.HorizontalAlignment = xlCenter
        rng.VerticalAlignment = xlCenter
    End If
End Sub

Sub AddRow()
    Dim rng As Range
    Set rng = Selection
    rng.Rows(1).EntireRow.Insert
End Sub

Sub AddColumn()
    Dim rng As Range
    Set rng = Selection
    rng.Columns(1).EntireColumn.Insert
End Sub

Sub UniqCount()

    Dim myRange As Range, myCell As Range, myCollection As New Collection, _
    myElement As Variant, i As Long
         
    Set myRange = Selection
         
    On Error Resume Next
    For Each myCell In myRange
        If Trim(myCell.Value) <> "" Then
        myCollection.Add CStr(myCell.Value), CStr(myCell.Value)
        End If
    Next myCell
    On Error GoTo 0
    MsgBox myCollection.Count

End Sub

Sub VerticalStrip()
    Set rng = Selection
    For i = 1 To rng.Columns.Count Step 2
        rng.Columns(i).Interior.ColorIndex = 15
    Next i
    With Selection.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

Sub HorizontalStrip()
    Set rng = Selection
    For i = 2 To rng.Rows.Count Step 2
        rng.Rows(i).Interior.ColorIndex = 15
    Next i
    
    With Selection.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    
End Sub
    
Sub AllBorders()
    With Selection.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub
   
    


