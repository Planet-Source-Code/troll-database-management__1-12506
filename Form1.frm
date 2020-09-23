VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data BDD 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3195
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   15
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton BtnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4365
      TabIndex        =   10
      Top             =   2745
      Width           =   1200
   End
   Begin VB.CommandButton BtnLast 
      Height          =   375
      Left            =   1230
      Picture         =   "Form1.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2745
      Width           =   395
   End
   Begin VB.CommandButton BtnNext 
      Height          =   375
      Left            =   855
      Picture         =   "Form1.frx":02EC
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2745
      Width           =   375
   End
   Begin VB.CommandButton BtnPrevious 
      Height          =   375
      Left            =   480
      Picture         =   "Form1.frx":048E
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2745
      Width           =   375
   End
   Begin VB.CommandButton BtnFirst 
      Height          =   375
      Left            =   105
      Picture         =   "Form1.frx":0630
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2745
      Width           =   375
   End
   Begin VB.Frame FrmCustomer 
      Caption         =   "Current Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   13
      Top             =   1020
      Width           =   5445
      Begin VB.TextBox TBCode 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Code"
         DataSource      =   "BDD"
         Height          =   315
         Left            =   3480
         TabIndex        =   17
         Top             =   270
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox TBSociety 
         DataField       =   "Society"
         DataSource      =   "BDD"
         Height          =   315
         Left            =   1275
         TabIndex        =   3
         Top             =   1050
         Width           =   3960
      End
      Begin VB.TextBox TBName 
         DataField       =   "Name"
         DataSource      =   "BDD"
         Height          =   315
         Left            =   1275
         TabIndex        =   2
         Top             =   660
         Width           =   3960
      End
      Begin VB.TextBox TBCodeTemp 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   270
         Width           =   1740
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Society"
         Height          =   255
         Left            =   210
         TabIndex        =   16
         Top             =   1065
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   675
         Width           =   930
      End
   End
   Begin VB.CommandButton BtnNew 
      Caption         =   "New (F2)"
      Height          =   375
      Left            =   1710
      TabIndex        =   8
      Top             =   2745
      Width           =   1290
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "Delete (F5)"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2745
      Width           =   1290
   End
   Begin VB.Frame Frame2 
      Caption         =   "Find Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   105
      TabIndex        =   11
      Top             =   120
      Width           =   5445
      Begin VB.TextBox TBValue 
         BackColor       =   &H00E4FFFF&
         Height          =   315
         Left            =   1275
         TabIndex        =   0
         Top             =   285
         Width           =   3960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Find String"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   225
      Left            =   3795
      TabIndex        =   18
      Top             =   1005
      Width           =   1710
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldCode As String

Private Sub BDD_Reposition()
    
    'Reminder current code
    OldCode = TBCode
    
    TBCodeTemp = TBCode
    
    'Refresh position in database
    Main.Caption = "Database" + " [" + CStr(BDD.Recordset.AbsolutePosition + 1) + " of " + CStr(Get_Number_Of_Record("Customer")) + "]"
End Sub

Private Sub BtnClose_Click()
    Unload Me
    End
End Sub

Private Sub BtnDelete_Click()
    Dim Code As String
    Dim NotDeleted As Boolean
    
    'Keep current Code value
    Code = TBCode
    NotDeleted = False
    
    'If Code is not null
    If TBCode <> "" Then
    
        'Save database to prevent error
        BDD.Refresh
        
        'Find record to be delete
        Find_Item "Code", Code
        
        'Confirm deletion
        If vbYes = MsgBox("Are you sure to delete this record?", vbQuestion + vbYesNo, "Attention") Then
            'Delete record
            BDD.Recordset.Delete
            BDD.Refresh
        Else
            NotDeleted = True
        End If
    End If
    
    'If deleting has not been aborted
    If NotDeleted = False Then
    
        'If the database if not empty
        If Not BDD.Recordset.EOF And Not BDD.Recordset.BOF Then
            'Go to first record
            BDD.Recordset.MoveFirst
        Else
            'Add a new record
            BDD.Recordset.AddNew
        End If

        Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
    End If
End Sub

Private Sub BtnFirst_Click()
    BDD.Recordset.MoveFirst
    Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
End Sub

Private Sub BtnLast_Click()
    BDD.Recordset.MoveLast
    Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
End Sub

Private Sub BtnNew_Click()
    
    'Save current database
    BDD.Refresh
    
    'Add the new record
    BDD.Recordset.AddNew

    'Disable all records buttons
    BtnFirst.Enabled = False
    BtnPrevious.Enabled = False
    BtnNext.Enabled = False
    BtnLast.Enabled = False

    'Focus on the database primary key
    TBCodeTemp.SetFocus
End Sub

Private Sub BtnNext_Click()
    BDD.Recordset.MoveNext
    Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
End Sub

Private Sub BtnPrevious_Click()
    BDD.Recordset.MovePrevious
    Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyPageUp
            KeyCode = 0
            If BtnPrevious.Enabled = True Then Call BtnPrevious_Click
        Case vbKeyPageDown
            KeyCode = 0
            If BtnNext.Enabled = True Then Call BtnNext_Click
        Case vbKeyF2
            KeyCode = 0
            Call BtnNew_Click
        Case vbKeyF5
            KeyCode = 0
            Call BtnDelete_Click
    End Select
End Sub

Private Sub Form_Load()
    If FileExists(App.Path + "\BDD.MDB") Then
        BDD.DatabaseName = App.Path + "\BDD.MDB"
        BDD.EOFAction = vbEOFActionAddNew
        BDD.Refresh
        If BDD.Recordset.EOF Or BDD.Recordset.BOF Then
            BDD.Recordset.AddNew
        End If
        CBCriteria = "Contains"
        Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
    Else
        MsgBox "Can't find BDD.MDB in folder.", vbCritical, "Error"
        Call BtnClose_Click
    End If
End Sub

Private Sub TBCodeTemp_GotFocus()
    Select_All TBCodeTemp
End Sub

Private Sub TBCodeTemp_KeyPress(KeyAscii As Integer)
    UCaseMask KeyAscii
    Max_Length TBCodeTemp, KeyAscii, 6
End Sub

Private Sub TBCodeTemp_Validate(Cancel As Boolean)
    Dim Code As String
    
    Code = TBCodeTemp
    
    'If Code already exist
    If Exists_Record("Customer", "Code", TBCodeTemp) And Code <> OldCode Then
        If All_Empty Then
            BDD.Recordset.CancelUpdate
        End If
        Find_Item "Code", Code
    Else
        If Code <> OldCode Then
            TBCode = TBCodeTemp
            BDD.Refresh
            Find_Item "Code", Code
        End If
    End If
    OldCode = TBCode
End Sub

Private Sub TBName_GotFocus()
    Select_All TBName
End Sub

Private Sub TBName_KeyPress(KeyAscii As Integer)
    Validate_Key_Not_Null KeyAscii
    Max_Length TBName, KeyAscii, 50
End Sub

Private Sub TBSociety_GotFocus()
    Select_All TBSociety
End Sub

Private Sub TBSociety_KeyPress(KeyAscii As Integer)
    Validate_Key_Not_Null KeyAscii
    Max_Length TBSociety, KeyAscii, 50
End Sub

Private Sub TBValue_GotFocus()
    Select_All TBValue
End Sub

Private Sub TBValue_KeyPress(KeyAscii As Integer)
    Max_Length TBValue, KeyAscii, 100
End Sub

Private Sub TBValue_Validate(Cancel As Boolean)

    If TBValue <> "" Then
        
        'Find string in Code field
        BDD.Recordset.FindFirst "Code" + " LIKE ""*" + TBValue + "*"""
        'If no match
        If BDD.Recordset.NoMatch Then
            'Find string in Name field
            BDD.Recordset.FindFirst "Name" + " LIKE ""*" + TBValue + "*"""
        End If
        'If no match
        If BDD.Recordset.NoMatch Then
            'Find string in Society field
            BDD.Recordset.FindFirst "Society" + " LIKE ""*" + TBValue + "*"""
        End If
        'If no match
        If BDD.Recordset.NoMatch Then
            MsgBox "No record were found.", vbInformation, "Searching"
        End If
        
        Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
        
    End If
End Sub

'If all fields are empty
Function All_Empty() As Boolean
    If TBCode <> "" Or TBName <> "" Or TBSociety <> "" Then
        All_Empty = False
    Else
        All_Empty = True
    End If
End Function

'Find a record
Function Find_Item(Field As String, Code As String, Optional Criteria As String = "=")
    If Not BDD.Recordset.EOF And Not BDD.Recordset.BOF Then
        BDD.Recordset.FindFirst Field + Criteria + " """ + Code + """"
        Manage_Records_Buttons BtnFirst, BtnPrevious, BtnNext, BtnLast, BDD
    End If
End Function

'Prevent enter informations before primary key
Function Validate_Key_Not_Null(KeyAscii As Integer)
    If All_Empty Then
        KeyAscii = 0
        MsgBox "You have to enter Code first.", vbInformation, "Error"
        TBCodeTemp.SetFocus
    End If
End Function
