Attribute VB_Name = "useful_functions"

 Public Sub CreateListBoxHeader(body As MSForms.ListBox, header As MSForms.ListBox, arrHeaders)
            ' make column count match
            header.ColumnCount = body.ColumnCount
            header.ColumnWidths = body.ColumnWidths

        ' add header elements
        header.Clear
        header.AddItem
        Dim i As Integer
        For i = 0 To UBound(arrHeaders)
            header.List(0, i) = arrHeaders(i)
        Next i

        ' make it pretty
        body.ZOrder (1)
        header.ZOrder (0)
        header.SpecialEffect = fmSpecialEffectFlat
        header.BackColor = RGB(200, 200, 200)
        header.Height = 10

        ' align header to body (should be done last!)
        header.Width = body.Width
        header.Left = body.Left
        header.Top = body.Top - (header.Height - 1)
End Sub

Public Function VIGENERE_ENCRYPT(pt As String, key As String)
'Encrypt plaintext pt using given key using Vigenere cipher.
'Assumes the key consists of a-z characters (either upper or lower case), no special characters
'including spaces. Plaintext may include spaces but otherwise only letters.

Dim ct As String, ptChar As String, keyChar As String, i As Integer

pt = Replace(pt, " ", "") ' Strip spaces from plaintext first

For i = 0 To Len(pt) - 1
    ptChar = Mid(pt, i + 1, 1)
    keyChar = Mid(key, (i Mod Len(key)) + 1, 1)
    ct = ct & Chr(((Asc(ptChar) + Asc(keyChar)) Mod 256))
Next i

VIGENERE_ENCRYPT = ct

End Function


Public Function VIGENERE_DECRYPT(ct As String, key As String)
'Encrypt plaintext pt using given key using Vigenere cipher.
'Assumes the key consists of a-z characters (either upper or lower case), no special characters
'including spaces. Plaintext may include spaces but otherwise only letters.

Dim pt As String, ptChar As String, keyChar As String, i As Integer

ct = Replace(ct, " ", "") ' Strip spaces from plaintext first

For i = 0 To Len(ct) - 1
    ctChar = Mid(ct, i + 1, 1)
    keyChar = Mid(key, (i Mod Len(key)) + 1, 1)
    pt = pt & Chr(((Asc(ctChar) + 256 - Asc(keyChar)) Mod 256))
Next i

VIGENERE_DECRYPT = pt

End Function

Public Sub protect()
    For Each sheet In ActiveWorkbook.Sheets
        sheet.protect ("1234")
    Next
End Sub

Public Sub unprotect()

    For Each sheet In ActiveWorkbook.Sheets
        sheet.unprotect ("1234")
    Next
End Sub

Sub ExpandAll()

         Dim Current As Worksheet

         For Each Current In Worksheets
            Current.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
         Next

End Sub

Sub CollapseAll()

         Dim Current As Worksheet

         For Each Current In Worksheets
            Current.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
         Next

End Sub

Public Function booli(ByVal income As Variant) As Integer
'returns valid boolean, solved problems with english deutsch
If (income) Then
    booli = 1
Else
    booli = 0
End If

End Function

Public Function sql_to_array(ByVal sql As String) As Variant
    Dim rs
    On Error GoTo Problem:
    Dim array1() As Variant
    Set rs = CurrentDb().OpenRecordset(sql)   'opens a recordset with the sqlstatment which was passed
    If rs.RecordCount = 0 Then        'abort if no entrie is found
        Exit Function
    End If

    With rs
        .MoveLast
        .MoveFirst
        sql_to_array = .GetRows(.RecordCount)   'pastes recordset to array
        .Close
    End With

    Exit Function
    
Problem:
    Resume Next

    End Function
    
    Public Function insert(ByVal table As String, ParamArray params() As Variant)
'ParamArray allows to provide X parameters
    On Error GoTo Err
    Dim fields As String
    Dim values As String

    For i = 0 To UBound(params)
    'loop all parameter pairs and use the first one as fieldname and the second one as value
        If i Mod 2 = 0 Then
            fields = komma(fields, params(i))
        Else
            values = SQLkomma(values, params(i))
        End If
    Next


    Dim sql As String
    sql = "Insert into " & table & "(" & fields & ")" & " values(" & values & ")"
    'Create sql string and run
Debug.Print (sql)
    CurrentDb.execute (sql)
    insert = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
    Exit Function

Err:
    MsgBox (Error)



End Function


Public Sub Delete(ByVal table As String, whereField As String, whereValue)
'ParamArray allows to provide X parameters
    On Error GoTo Err

    Dim sql As String

    'read the parameter Array and chain them for update     Delete *from table where WhereFiled='whereValue'
    sql = "Delete * from " & table & " where " & whereField & "=" & SQLkomma("", whereValue)

    'run sql
    Debug.Print (sql)
    CurrentDb.execute (sql)
    Exit Sub

Err:
    MsgBox (Error)
   

End Sub

Public Sub Update(ByVal table As String, ByVal whereField As String, ByVal whereValue As String, ParamArray params() As Variant)
'ParamArray allows to provide X parameters

    On Error GoTo Err
    Dim fields As String
    Dim i As Integer
    i = 0

    'read the parameter Array and chain them for update     Field="Value"
    Do Until i > UBound(params)
        fields = komma(fields, params(i)) & "=" & SQLkomma("", params(i + 1))
        
        i = i + 2
    Loop

    'create sql string with variables
    Dim sql As String
    sql = "Update " & table & " set " & fields & " where " & whereField & "=" & SQLkomma("", whereValue)
    'run sql
    Debug.Print (sql)
    CurrentDb.execute (sql)
    Exit Sub

Err:
    MsgBox (Error)


End Sub

Public Function checkInTable(ByVal table As String, ByVal column As String, ByVal value As String) As Boolean
    arr = sql_to_array("Select count(*) from " & table & " where " & column & "='" & value & "'")
    If arr(0, 0) >= 1 Then checkInTable = True Else: checkTable = False
End Function

Public Function map(ByVal field As String, Optional ByVal table As String = "AssetGroups") As Integer
'matched the wanted field with the given table. Returns an index

Dim found As Boolean
found = False

For i = 0 To CurrentDb.TableDefs(table).fields.Count - 1
    If CurrentDb.TableDefs(table).fields(i).Name = field Then
        map = i
        found = True
    End If
   
Next
If found = False Then
    'if no column is found raise error
     Err.Raise Number:=vbObjectError + 513, Description:="No colum found"
End If

Set fld = Nothing

End Function

Public Function sqlMap(ByVal sql As String, ByVal field As String) As Integer
'matched the wanted field with the given sqlStatement. Returns an index
    Dim rs
    Dim found As Boolean
    
    Set rs = CurrentDb().OpenRecordset(sql)   'opens a recordset with the sqlstatment which was passed
    If rs.RecordCount = 0 Then        'abort if no entrie is found
        Exit Function
    End If

    For i = 0 To rs.fields.Count - 1
    If rs.fields(i).Name = field Then
        sqlMap = i
        found = True
    End If
    Next
    
    If found = False Then
    'matched the wanted field with the given table. Returns an index
        rs.Close
         Err.Raise Number:=vbObjectError + 513, Description:="No colum found"
    End If

    rs.Close
    Exit Function
    


End Function

Public Function simicolon(ByVal old As String, ByVal neu As String, byval seperator as string) as string
    Dim buffer As String

    'first item gets no komma
    If old = "" Then
        buffer = neu                'first item is added to string without komma
    Else
        buffer = old & seperator & neu    'other elements are added to the string and are seperated by a komma
    End If
    
    simicolon = buffer
End Function

Private Function checkIntegerInString(ByVal str As String) As Boolean

 Dim i As Integer

    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Then
            HasNumber = True
            Exit Function
        End If
    Next i
     
     
End Function

Public Function inarr(ByVal arr As Variant, ByVal wert As String) As Boolean

For i = 0 To UBound(arr)        'loop array
    If wert = arr(i) Then       'if value is in array
        inarr = True            'return true
        Exit Function           'and exit function
    End If
Next

inarr = False                   'otherwise return false

End Function

Public Sub append(ByRef arr As Variant, ByVal value As Variant)
On Error GoTo hier
    ReDim Preserve arr(0 To UBound(arr) + 1) As Variant
    arr(UBound(arr)) = value
    Exit Sub
hier:
    arr = Array(value)
    
End Sub

Public Function IsArrayAllocated(Arr As Variant) As Boolean
        On Error Resume Next
        IsArrayAllocated = IsArray(Arr) And _
                           Not IsError(LBound(Arr, 1)) And _
                           LBound(Arr, 1) <= UBound(Arr, 1)
End Function

Public Function inlistbox(ByVal box As String, ByVal val As String) As Boolean
'check if value is in listox or dropbox
Dim boo As Boolean
boo = False
For i = 0 To Me.Controls(box).ListCount - 1
        If CStr(Me.Controls(box).List(i, 0)) = CStr(val) Then
            boo = True
        End If
Next
inlistbox = boo

End Function

Public Sub sortListbox(ByRef box As msforms.ListBox, ByVal columns As Integer)
   'Sort listbox
   ' !!!!!!!!!! listboxes with a rowsource cannot be sorted. Only listboxes with are used with "addItem" can be sorted"!!!!!!
   '
   Dim buffer() As String
   ReDim buffer(columns)
 
  For a = 0 To box.ListCount - 1
      For b = a To box.ListCount - 1
            If box.List(b, 0) < box.List(a, 0) Then
            
                For i = 0 To box.ColumnCount - 1
                    buffer(i) = box.List(a, i)
                Next
                For i = 0 To box.ColumnCount - 1
                    box.List(a, i) = box.List(b, i)
                    box.List(b, i) = buffer(i)
                Next
                
           End If
      Next
  Next

End Sub

Public Function GetFilenameFromPath(ByVal strPath As String) As String

    'returns the file name from a path. Recursive function
    If Right(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left(strPath, Len(strPath) - 1)) + Right(strPath, 1)
    End If
    
End Function

Public Function checkFile(ByVal path As String) As Boolean
    'checks file of existing
    checkFile = Dir(path) <> ""
End Function

Public Function checkFolder(ByVal path As String) As Boolean
    'checks file of existing
    checkFolder = Dir(path, vbDirectory) <> ""
End Function

Public Function generateFolder(byval path as string, byval foldername as string) As String
'call checkFolder first
            MkDir  path & folderName 
            
End Function

Public Function getFiles(ByVal path As String) As Variant

    Dim file As String
    Dim Arr() As Variant
                        ' add latest backslash if not existing
    file = Dir(path & IIf(Mid(path, Len(path)) = "\", "", "\") & "*")
    Do While Len(file) > 0
        Call append(Arr, file)
        file = Dir()
    Loop
    getFiles = Arr

End Function

Public Sub appointment(ByVal subject As String, ByVal datum As String, ByVal body As String)
'created an appointment in the outlook calender

Debug.Print ("Appointment was created")

 Dim App As Object
  Set App = CreateObject("Outlook.Application")
  Set myItem = App.CreateItem(1)
    myItem.MeetingStatus = olNonMeeting
    
    myItem.subject = subject
    myItem.body = body
    'set start to 12 pm
    myItem.Start = "12:00 " & datum
    myItem.Save

MsgBox ("Appointment created")
End Sub

Public Sub Sendmail(ByVal subject As String, ByVal body As String, ByVal who As String, cc As String, attachments As String)
    'sends a mail with the given parameters
    'reference in code extras references musst be added
    On Error GoTo Ende
        Dim oApp As Object
    
        Dim oMail As Object
        Set oApp = CreateObject("Outlook.application")  'create an outlook object

        Set oMail = oApp.CreateItem(olMailItem) 'create email and fill with data
        oMail.HTMLBody = body
    
        oMail.subject = subject
        oMail.to = who
        oMail.cc = cc
        
        
        Attachment = Split(attachments, ";")
        For i = 1 To UBound(Attachment)
            If Attachment(i) <> "" Then
                oMail.attachments.Add (Attachment(i))
            End If
        Next
        
        oMail.Save                              'send mail and clears the objects
        Set oMail = Nothing
        Set oApp = Nothing
        Exit Sub
Ende:

    MsgBox (Error)


End Sub

Public Function getUsername()
    getUsername = Environ$("Username")                     'use the windows system name
End Function

Public Function findColumn(ByVal rang As String, ByVal val As String, ByVal sh As Worksheet, Optional ByVal searchtype As Integer = 0)
    'match whole cell or part of
    
    On Error GoTo hier
    If searchtype = 0 Then
        findColumn = sh.Range(rang).find(val, LookIn:=xlValues, LookAt:=xlWhole).Column
    Else
        findColumn = sh.Range(rang).find(val, LookIn:=xlValues, LookAt:=xlPart).Column
    End If
        Exit Function
    hier:

    findColumn = -1
    
End Function


Public Function findRow(ByVal rang As String, ByVal val As String, ByVal sh As Worksheet, Optional ByVal searchtype As Integer = 0)
    'match whole cell or part of
    On Error GoTo hier
    If searchtype = 0 Then
        findRow = sh.Range(rang).Find(val, LookIn:=xlValues, LookAt:=xlWhole).Row
    Else
        findRow = sh.Range(rang).Find(val, LookIn:=xlValues, LookAt:=xlPart).Row
    End If
    Exit Function
hier:

    findRow = -1
End Function

Public Function lastrow(Optional ByVal sheet As String = "") As Integer
    'returns the last row of the given worksheet. If not given returns last row of active worksheet
    On Error GoTo err
    
    If sheet = "" Then
        lastrow = (ActiveWorkbook.ActiveSheet.Cells.find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row)
    Else
        lastrow = (ActiveWorkbook.Sheets(sheet).Cells.find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row)
    
    End If
    Exit Function
    
    err:
    MsgBox ("Tabelle ist nicht vorhanden")

End Function



Public Function lastCol(Optional ByVal sheet As String = "") As Integer
    'returns the last column of the given worksheet. If not given returns last column of active worksheet
    On Error GoTo err
    
    If sheet = "" Then
        lastCol = (ActiveWorkbook.ActiveSheet.Cells(1, ActiveWorkbook.ActiveSheet.Columns.Count).End(xlToLeft).Column)
    Else
        lastCol = (ActiveWorkbook.Sheets(sheet).Cells(1, ActiveWorkbook.ActiveSheet.Columns.Count).End(xlToLeft).Column)
    
    End If
    
    Exit Function
    
    err:
    MsgBox ("Tabelle ist nicht vorhanden")
    'given table is not existing

End Function

Public Function openFile(Optional excelFilter As Boolean = False, Optional ByVal startPath As String = "")

    Dim FILE_DIA As FileDialog
    
    Set FILE_DIA = Application.FileDialog(msoFileDialogFilePicker)
    
    With FILE_DIA
    
        .AllowMultiSelect = False
        .InitialFileName = "C:\users\" & Environ$("Username") & "\" & startPath
        
        If excelFilter Then
        .Filters.Add "BOXI Spreadsheets", "*.xls; *.xlsx; *.xlsm; *.csv", 1
        End If
        If .Show = -1 Then
                
            openFile = .SelectedItems(1)
        Else
            openFile = Null
        End If

    End With
    
    Set FILE_DIA = Nothing

End Function
