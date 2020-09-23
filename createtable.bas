Attribute VB_Name = "Module1"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30

'--- ListView Set Column Width Messages ---'
Public Enum LVSCW_Styles
   LVSCW_AUTOSIZE = -1
   LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

Public Sub LVSetAllColWidths(lv As ListView, ByVal Style As LVSCW_Styles)
   Dim ColumnIndex As Long
   '--- loop through all of the columns in the listview and size each
   With lv
      For ColumnIndex = 1 To .ColumnHeaders.Count
         LVSetColWidth lv, ColumnIndex, Style
      Next ColumnIndex
   End With
End Sub

Public Sub LVSetColWidth(lv As ListView, ByVal ColumnIndex As Long, ByVal Style As LVSCW_Styles)
   '------------------------------------------------------------------------------
   '--- If you include the header in the sizing then the last column will
   '--- automatically size to fill the remaining listview width.
   '------------------------------------------------------------------------------
   With lv
      ' verify that the listview is in report view and that the column exists
      If .View = lvwReport Then
         If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
            Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
         End If
      End If
   End With
End Sub

Sub Listtxt()

'List1.Visible = False
'Call LVSetAllColWidths(List1, LVSCW_AUTOSIZE_USEHEADER)
'List1.Visible = True

End Sub
