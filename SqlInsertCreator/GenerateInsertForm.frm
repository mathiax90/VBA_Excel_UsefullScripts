VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GenerateInsertForm 
   Caption         =   "SqlInsertCreator"
   ClientHeight    =   5790
   ClientLeft      =   12045
   ClientTop       =   4380
   ClientWidth     =   9930
   OleObjectBlob   =   "GenerateInsertForm.frx":0000
End
Attribute VB_Name = "GenerateInsertForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ColumnNameRowSelectButton_Click()
Dim rng As Range


'Temporarily Hide Userform
  Me.Hide
start:
'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Выбор строки с названиями", _
      Prompt:="Выберите строку с названиями колонок таблицы.", _
      Type:=8)
  On Error GoTo 0

'Test to ensure User Did not cancel
  If rng Is Nothing Then
    Me.Show 'unhide userform
    Exit Sub
  End If

If rng.Rows.Count > 1 Then
    MsgBox "Разрешено выбирать только одну строку"
    GoTo start
End If

ColumnNameRowTextBox = rng.Address
ColumnTypeRowTextBox = rng.Offset(1, 0).Address
ValueRowTextBox = rng.Offset(2, 0).Address
InsertCellTextBox = rng.Cells(1, rng.Columns.Count + 1).Offset(2, 0).Address


SaveSetting "GenerateInsert", "Main", "ColumnNameRow", ColumnNameRowTextBox.Text

'Unhide Userform
  Me.Show
End Sub

Private Sub ColumnTypeRowSelectButton_Click()
Dim rng As Range


'Temporarily Hide Userform
  Me.Hide

'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Выбор строки с типами", _
      Prompt:="Выберите строку с типами колонок таблицы.", _
      Type:=8)
  On Error GoTo 0

'Test to ensure User Did not cancel
  If rng Is Nothing Then
    Me.Show 'unhide userform
    Exit Sub
  End If
  
If rng.Rows.Count > 1 Then
    MsgBox "Разрешено выбирать только одну строку"
    GoTo start
End If

  
ColumnTypeRowTextBox = rng.Address
SaveSetting "GenerateInsert", "Main", "ColumnTypeRow", ColumnTypeRowTextBox.Text
'Unhide Userform
  Me.Show
End Sub



Private Sub ColumnTypeRowTextBox_Change()

End Sub

Private Sub CommandButton1_Click()
'Temporarily Hide Userform
  'Me.Hide

GenerateInsertHelpForm.Show (vbModal)
'Unhide Userform
  'Me.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me

End Sub

Private Sub GenerateInsertButton_Click()




If ColumnNameRowTextBox.Text = "" Then
    MsgBox "Не выбраны наименования колонок таблицы"
    Exit Sub
End If

If ColumnTypeRowTextBox.Text = "" Then
    MsgBox "Не выбраны типы колонок таблицы"
    Exit Sub
End If

If ValueRowTextBox.Text = "" Then
    MsgBox "Не выбраны данные для Insert"
    Exit Sub
End If


If InsertCellTextBox.Text = "" Then
    MsgBox "Не выбран адрес для вставки Insert"
    Exit Sub
End If

If IdGenCheckBox.Value And IdGenTextBox.Text = "" Then
    MsgBox "Не ввёдено название поля для newid()"
    Exit Sub
End If

Dim ColumnNameRng As Range
Dim ColumnTypeRng As Range
Dim ValueRowRng As Range
Dim InsertCell As Range

Set ColumnNameRng = Range(ColumnNameRowTextBox.Text)
Set ColumnTypeRng = Range(ColumnTypeRowTextBox.Text)
Set ValueRowRng = Range(ValueRowTextBox.Text)
Set InsertCell = Range(InsertCellTextBox.Text)

If ColumnNameRng.Columns.Count <> ColumnTypeRng.Columns.Count Then
     MsgBox "Количество наименований не равно количеству типов"
    Exit Sub
End If

If ValueRowRng.Columns.Count <> ColumnTypeRng.Columns.Count Then
     MsgBox "Количество типов не равно количеству данных"
    Exit Sub
End If


Dim InsertCellRow As Long
Dim InsertCellCol As Long

InsertCellRow = InsertCell.Row
InsertCellCol = InsertCell.Column

Dim TableName As String

If TableNameTextBox.Text = "" Then
    TableName = "table_name"
Else
    TableName = TableNameTextBox.Text
End If

Dim ColumnList As String

ColumnList = "INSERT INTO " & TableName & " ("

If IdGenCheckBox Then
    ColumnList = ColumnList & IdGenTextBox.Text & ", "
End If


For i = 1 To ColumnNameRng.Columns.Count
    ColumnList = ColumnList & ColumnNameRng(1, i) & ", "
Next i
ColumnList = Left(ColumnList, Len(ColumnList) - 2) & ") values ("

If IdGenCheckBox Then
    ColumnList = ColumnList & "newid(), "
End If

InsertCell.Value = ColumnList

For i = 1 To ColumnTypeRng.Columns.Count
    cType = ColumnTypeRng.Cells(1, i).Value
    vAddr = ValueRowRng.Cells(1, i).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Set TargetCell = Cells(InsertCellRow, InsertCellCol + i)
        
    If i = ColumnTypeRng.Columns.Count Then
        'если последняя строка то в формулах не нужна запятая
        If LCase(cType) Like "*varchar*" Then
            If NullStringCheckBox Then
                formula_ = "=ЕСЛИ(" & vAddr & "="""";""Null"";""'"" & " & vAddr & " & ""'"")"
            Else
                formula_ = "=""'"" & " & vAddr & " & ""'"""
            End If
            GoTo nextrow
        End If
        
        If LCase(cType) = "guid" Then
            formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")<>"""";""'"" & " & vAddr & " & ""'"";""null"")"
            GoTo nextrow
        End If
        
        If LCase(cType) Like "*decimal*" Then
            formula_ = "=ЕСЛИ(" & vAddr & "="""";""NULL"";ПОДСТАВИТЬ(" & vAddr & ";"","";""."") & """")"
            GoTo nextrow
        End If
        
        If LCase(cType) = "date" Then
            If IsDate(ValueRowRng.Cells(1, i)) Then
                 formula_ = "=ЕСЛИ(" & vAddr & "="""";""NULL""; ""'"" & ГОД(" & vAddr & ") & ""-"" & ПРАВСИМВ(""0"" & МЕСЯЦ(" & vAddr & ");2) & ""-"" & ПРАВСИМВ(""0"" & ДЕНЬ(" & vAddr & ");2) & ""'"")"
            Else
                If ParseDateCheckBox Then
                    formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")="""";""NULL"";""'"" & ПРАВСИМВ(" & vAddr & ";4) & ""-"" & ПСТР(" & vAddr & ";4;2) & ""-"" & ЛЕВСИМВ(" & vAddr & ";2) & ""'"")"
                    GoTo nextrow
                Else
                    formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")<>"""";""'"" & " & vAddr & " & ""'"";""null"")"
                    GoTo nextrow
                End If
            End If
            GoTo nextrow
        End If
        
        formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")="""";""NULL"";" & vAddr & "&"""")"
    
    Else
        If LCase(cType) Like "*varchar*" Then
            If NullStringCheckBox Then
                formula_ = "=ЕСЛИ(" & vAddr & "="""";""Null,"";""'"" & " & vAddr & " & ""',"")"
            Else
                formula_ = "=""'"" & " & vAddr & " & ""',"""
            End If
            GoTo nextrow
        End If
        
        If LCase(cType) = "guid" Then
            formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")<>"""";""'"" & " & vAddr & " & ""',"";""null,"")"
            GoTo nextrow
        End If
        
        If LCase(cType) Like "*decimal*" Then
            formula_ = "=ЕСЛИ(" & vAddr & "="""";""NULL,"";ПОДСТАВИТЬ(" & vAddr & ";"","";""."") & "","")"
            GoTo nextrow
        End If
        
        If LCase(cType) = "date" Then
            If IsDate(ValueRowRng.Cells(1, i)) Then
                 formula_ = "=ЕСЛИ(" & vAddr & "="""";""NULL,""; ""'"" & ГОД(" & vAddr & ") & ""-"" & ПРАВСИМВ(""0"" & МЕСЯЦ(" & vAddr & ");2) & ""-"" & ПРАВСИМВ(""0"" & ДЕНЬ(" & vAddr & ");2) & ""',"")"
            Else
                If ParseDateCheckBox Then
                    formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")="""";""NULL,"";""'"" & ПРАВСИМВ(" & vAddr & ";4) & ""-"" & ПСТР(" & vAddr & ";4;2) & ""-"" & ЛЕВСИМВ(" & vAddr & ";2) & ""',"")"
                    GoTo nextrow
                Else
                    formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")<>"""";""'"" & " & vAddr & " & ""',"";""null,"")"
                    GoTo nextrow
                End If
            End If
            GoTo nextrow
        End If
        
        formula_ = "=ЕСЛИ(СЖПРОБЕЛЫ(" & vAddr & ")="""";""NULL,"";" & vAddr & "&"","")"
    End If
nextrow:
    TargetCell.FormulaLocal = formula_
Next i

    Set TargetCell = Cells(InsertCellRow, InsertCellCol + ColumnTypeRng.Columns.Count + 1)

    TargetCell.Value = ")"
    
Set tableCreateCell = Range(InsertCellTextBox.Text).Offset(-1, 0)
tableCreateCell.Value = "create table " & TableName & "(" & Chr(10)

For Each cl In ColumnNameRng
    tableCreateCell.Value = tableCreateCell.Value & cl.Value & " " & ColumnTypeRng(1, cl.Column).Value & "," & Chr(10)
Next

tableCreateCell.Value = Left(tableCreateCell.Value, Len(tableCreateCell.Value) - 2) & Chr(10)
tableCreateCell.Value = tableCreateCell.Value & ")"



MsgBox "Done!"
End Sub

Private Sub IdGenCheckBox_Click()
'MsgBox IdGenCheckBox.Value

If IdGenCheckBox.Value Then
    IdGenTextBox.Enabled = True
Else
    IdGenTextBox.Enabled = False
End If


End Sub






Private Sub IdGenTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveSetting "GenerateInsert", "Main", "IdGen", IdGenTextBox.Text
End Sub

Private Sub InsertCellSelectButton_Click()
Dim rng As Range


'Temporarily Hide Userform
  Me.Hide

'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Выбор строкe с данными", _
      Prompt:="Выберите строку с данными.", _
      Type:=8)
  On Error GoTo 0

'Test to ensure User Did not cancel
  If rng Is Nothing Then
    Me.Show 'unhide userform
    Exit Sub
  End If
  
InsertCellTextBox = rng.Cells(1, 1).Address
SaveSetting "GenerateInsert", "Main", "InsertCell", InsertCellTextBox.Text

'Unhide Userform
  Me.Show
End Sub




Private Sub TableNameTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SaveSetting "GenerateInsert", "Main", "TableName", TableNameTextBox.Text
End Sub

Private Sub UserForm_Initialize()

ColumnNameRowTextBox.Text = GetSetting("GenerateInsert", "Main", "ColumnNameRow", "")
ColumnTypeRowTextBox.Text = GetSetting("GenerateInsert", "Main", "ColumnTypeRow", "")
ValueRowTextBox.Text = GetSetting("GenerateInsert", "Main", "ValueRow", "")
IdGenTextBox.Text = GetSetting("GenerateInsert", "Main", "IdGen", "")
InsertCellTextBox.Text = GetSetting("GenerateInsert", "Main", "InsertCell", "")
TableNameTextBox.Text = GetSetting("GenerateInsert", "Main", "TableName", "")
'SaveSetting AppName, Section, Key, Setting '// AppName - название вашей программы,
 
'Чтение данных из реестра:

   
End Sub

Private Sub ValueRowSelectButton_Click()
Dim rng As Range


'Temporarily Hide Userform
  Me.Hide

'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Выбор строкe с данными", _
      Prompt:="Выберите строку с данными.", _
      Type:=8)
  On Error GoTo 0

'Test to ensure User Did not cancel
  If rng Is Nothing Then
    Me.Show 'unhide userform
    Exit Sub
  End If
  
If rng.Rows.Count > 1 Then
    MsgBox "Разрешено выбирать только одну строку"
    GoTo start
End If
  
ValueRowTextBox = rng.Address
SaveSetting "GenerateInsert", "Main", "ValueRow", ValueRowTextBox.Text


'Unhide Userform
  Me.Show
End Sub
