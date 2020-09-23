Attribute VB_Name = "Functions"
'Use to manage records buttons (First, Previous, Next, Last)
Function Manage_Records_Buttons(Frst As CommandButton, Prvs As CommandButton, Nxt As CommandButton, Lst As CommandButton, DB As Data)
    Dim Pos As Integer
    Dim NbRec As Integer
           
    'Get position and number of records
    Pos = Se(DB.Recordset.AbsolutePosition + 1, 1)
    NbRec = Se(Get_Number_Of_Record("Customer"), 1)
    
    'If there is many records
    If NbRec > 1 Or Pos > 1 Then
        'If the current record is the last of the recordset
        If Pos >= NbRec Then
            If Frst.Enabled = False Then Frst.Enabled = True
            If Prvs.Enabled = False Then Prvs.Enabled = True
            If Nxt.Enabled = True Then Nxt.Enabled = False
            If Lst.Enabled = True Then Lst.Enabled = False
        End If
        'If the current record is the first of the recordset
        If Pos = 1 Then
            If Nxt.Enabled = False Then Nxt.Enabled = True
            If Lst.Enabled = False Then Lst.Enabled = True
            If Frst.Enabled = True Then Frst.Enabled = False
            If Prvs.Enabled = True Then Prvs.Enabled = False
        End If
        'If the current record is not the first and not the last
        If Pos > 1 And Pos < NbRec Then
            If Frst.Enabled = False Then Frst.Enabled = True
            If Prvs.Enabled = False Then Prvs.Enabled = True
            If Nxt.Enabled = False Then Nxt.Enabled = True
            If Lst.Enabled = False Then Lst.Enabled = True
        End If
    Else
        'If there is less than 1 record
        If Frst.Enabled = True Then Frst.Enabled = False
        If Prvs.Enabled = True Then Prvs.Enabled = False
        If Nxt.Enabled = True Then Nxt.Enabled = False
        If Lst.Enabled = True Then Lst.Enabled = False
    End If
    
End Function

'Use to get the number of record of a table
Function Get_Number_Of_Record(Table As String) As Integer
    Dim dbs As Database
    Dim rst As Recordset
   
    Set dbs = OpenDatabase(App.Path + "\BDD.MDB")
    Set rst = dbs.OpenRecordset("SELECT Count(*)AS Nombre FROM " + Table)
    If Not rst.EOF Then
        Get_Number_Of_Record = rst![Nombre]
        rst.Close
    Else
        Get_Number_Of_Record = 0
    End If
    dbs.Close
End Function

'Use to verify if a record is existing
Function Exists_Record(Table As String, Field As String, Value As String) As Boolean
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = OpenDatabase(App.Path + "\BDD.MDB")
    Set rst = dbs.OpenRecordset("SELECT * FROM " + Table + " WHERE " + Field + "=""" + Value + """")
    If Not rst.EOF Then
        Exists_Record = True
        rst.Close
    Else
        Exists_Record = False
    End If
    dbs.Close
End Function

'Use to return as value if a string is empty or null
Function Se(Str As Variant, Optional ReturnValueIfEmpty As Variant = "") As Variant
    If Str <> "" Then
        If IsNumeric(ReturnValueIfEmpty) Then
            If IsNumeric(Str) Then
                Se = Str
            Else
                Se = 0
            End If
        Else
            Se = Str
        End If
    Else
        Se = ReturnValueIfEmpty
    End If
End Function

'Use in Event KeyPress to limit number of character in textbox
Function Max_Length(Ctl As Control, KeyAscii As Integer, MaxLength As Integer)
    
    'If key is <Return>
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        'Go to next control.
        SendKeys "{TAB}"
    Else
        'If the max length is not zero
        If MaxLength <> 0 Then
            'If the maximum length has been reached
            If KeyAscii <> vbKeyBack And (Len(Ctl.Text) >= MaxLength) Then
                'If no text is selected
                If Ctl.SelLength = 0 Then
                    'Don't accept more characters
                    KeyAscii = 0
                    Beep
                End If
            End If
        End If
    End If
End Function

'Select all the text in a control
Function Select_All(Ctl As Control)
    If Ctl <> "" Then
        Ctl.SelStart = 0
        Ctl.SelLength = Len(Ctl)
    End If
End Function

'Use in Event KeyPress to set all characters in uppercase
Function UCaseMask(KeyAscii As Integer) As Integer
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Function

'Verify if a file exists
Function FileExists(FileName As String) As Boolean
    Dim l As Long
    
    On Error Resume Next
    
    l = FileLen(FileName)
    
    FileExists = Not (Err.Number > 0)
    
    On Error GoTo 0
End Function
