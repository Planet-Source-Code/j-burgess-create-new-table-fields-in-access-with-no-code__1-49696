VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAO Create Table Component"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   Icon            =   "createTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Field"
      Height          =   465
      Left            =   7290
      TabIndex        =   13
      Top             =   2115
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Build Code"
      Height          =   465
      Left            =   7305
      TabIndex        =   12
      Top             =   4260
      Width           =   1500
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto Incrementing Field"
      Height          =   270
      Left            =   4350
      TabIndex        =   11
      Top             =   5355
      Value           =   1  'Checked
      Width           =   2160
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4305
      TabIndex        =   10
      Text            =   "ID"
      Top             =   5025
      Width           =   2820
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   210
      TabIndex        =   8
      Text            =   "ID"
      Top             =   5025
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   165
      TabIndex        =   0
      Text            =   "TableName"
      Top             =   90
      Width           =   4005
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4350
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   750
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Field"
      Height          =   465
      Left            =   7275
      TabIndex        =   3
      Top             =   1560
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   735
      Width           =   4005
   End
   Begin MSComctlLib.ListView List1 
      Height          =   3180
      Left            =   195
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1545
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5609
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Field Name"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type            "
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Primary Index"
      Height          =   240
      Index           =   3
      Left            =   255
      TabIndex        =   9
      Top             =   5385
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "Table Name"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   435
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   240
      Index           =   1
      Left            =   4410
      TabIndex        =   6
      Top             =   1140
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "Field Name"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1095
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer

Private Sub Combo1_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub Command1_Click()

With List1.ListItems.Add
 .Text = Text1.Text
End With

With List1.ListItems(X).ListSubItems.Add
  .Text = Combo1.Text
End With

X = X + 1

List1.Visible = False
Call LVSetAllColWidths(List1, LVSCW_AUTOSIZE_USEHEADER)
List1.Visible = True

Text1.Text = ""
Text1.SetFocus

End Sub

Private Sub Command4_Click()
CreateTable
End Sub

Private Sub Command5_Click()
Dim RemoveMe As String

If List1.ListItems.Count = 0 Then Exit Sub

List1.ListItems.Remove (List1.SelectedItem.Index)

X = X - 1

End Sub

Private Sub Form_Load()

X = 1

Combo1.AddItem "dbText"
Combo1.AddItem "dbMemo"
Combo1.AddItem "dbLong"
Combo1.AddItem "dbDate"
Combo1.AddItem "dbSingle"
Combo1.AddItem "dbCurrency"
Combo1.AddItem "dbBoolean"
Combo1.ListIndex = 0

List1.Visible = False
Call LVSetAllColWidths(List1, LVSCW_AUTOSIZE_USEHEADER)
List1.Visible = True

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If Trim(Text1.Text) = "" Then Exit Sub

If KeyAscii = 13 Then Call Command1_Click

End Sub

Sub CreateTable()
Dim TableName As String
Dim FieldName As String
Dim FieldType As String
Dim CT As String
Dim X As Integer

TableName = Trim(Text2.Text)
TableName = Replace(TableName, " ", "_")

CT = "Sub Create" & TableName & "Table()" & vbCrLf
CT = CT & "Dim tblObjectTracking As TableDef" & vbCrLf
CT = CT & "Dim AddField As Field" & vbCrLf
CT = CT & "Dim X As Integer" & vbCrLf
CT = CT & "Dim Fld As Field" & vbCrLf & vbCrLf

CT = CT & "On Error Resume Next" & vbCrLf & vbCrLf

CT = CT & "Set db = OpenDatabase(APP_LOC, False, False, " & Chr(34) & ";pwd=" & Chr(34) & " & APP_PWD)" & vbCrLf & vbCrLf

CT = CT & "'create new table" & vbCrLf
CT = CT & "Set tblObjectTracking = db.CreateTableDef(" & Chr(34) & TableName & Chr(34) & ")" & vbCrLf & vbCrLf

For X = 1 To List1.ListItems.Count
  FieldName = List1.ListItems.Item(X).Text
  FieldName = Replace(FieldName, " ", "_")
  FieldType = List1.ListItems.Item(X).ListSubItems(1).Text
  
  CT = CT & "'define new field: " & FieldName & vbCrLf
  
    If FieldType = "dbText" Then
       CT = CT & "Set Fld = tblObjectTracking.CreateField(" & Chr(34) & FieldName & Chr(34) & ", dbText, 250)" & vbCrLf & vbCrLf
       CT = CT & "With Fld" & vbCrLf
       CT = CT & " .AllowZeroLength = True" & vbCrLf
       CT = CT & "End With" & vbCrLf & vbCrLf
    ElseIf FieldType = "dbMemo" Then
       CT = CT & "Set Fld = tblObjectTracking.CreateField(" & Chr(34) & FieldName & Chr(34) & ", dbMemo)" & vbCrLf & vbCrLf
       CT = CT & "With Fld" & vbCrLf
       CT = CT & " .AllowZeroLength = True" & vbCrLf
       CT = CT & "End With" & vbCrLf & vbCrLf
    Else
       CT = CT & "Set Fld = tblObjectTracking.CreateField(" & Chr(34) & FieldName & Chr(34) & ", " & FieldType & ")" & vbCrLf & vbCrLf
    End If
    
    CT = CT & "'append to table" & vbCrLf
    CT = CT & "With tblObjectTracking.Fields" & vbCrLf
    CT = CT & " .Append Fld" & vbCrLf
    CT = CT & " .Refresh" & vbCrLf
    CT = CT & "End With" & vbCrLf & vbCrLf
    
Next X

If Check1.Value <> 0 Then
  FieldName = Trim(Text4.Text)
  FieldName = Replace(FieldName, " ", "_")

  CT = CT & "'add an auto incrementing field for ID" & vbCrLf
  CT = CT & "Set Fld = tblObjectTracking.CreateField(" & Chr(34) & FieldName & Chr(34) & ", dbLong)" & vbCrLf & vbCrLf

  CT = CT & "With Fld" & vbCrLf
  CT = CT & " .Attributes = .Attributes Or dbAutoIncrField" & vbCrLf
  CT = CT & "End With" & vbCrLf & vbCrLf
  
  CT = CT & "'append to table" & vbCrLf
  CT = CT & "With tblObjectTracking.Fields" & vbCrLf
  CT = CT & ".Append Fld" & vbCrLf
  CT = CT & ".Refresh" & vbCrLf
  CT = CT & "End With" & vbCrLf & vbCrLf
End If

CT = CT & "'append all fields to newly created table" & vbCrLf
CT = CT & "db.TableDefs.Append tblObjectTracking" & vbCrLf & vbCrLf

If Trim(Text3.Text) <> "" Then
  FieldName = Trim(Text3.Text)
  FieldName = Replace(FieldName, " ", "_")

  CT = CT & "'sets primary index on ID field" & vbCrLf
  CT = CT & "db.Execute " & Chr(34) & "CREATE UNIQUE INDEX PrimaryKey ON " & TableName & " (" & FieldName & " ASC) WITH PRIMARY DISALLOW NULL" & Chr(34) & "" & vbCrLf & vbCrLf
End If

CT = CT & "db.Close: Set db = Nothing" & vbCrLf & vbCrLf

CT = CT & "Set tblObjectTracking = Nothing" & vbCrLf & vbCrLf

CT = CT & "End Sub" & vbCrLf

Form2.Text1.Text = CT
Form2.Show 1

End Sub
