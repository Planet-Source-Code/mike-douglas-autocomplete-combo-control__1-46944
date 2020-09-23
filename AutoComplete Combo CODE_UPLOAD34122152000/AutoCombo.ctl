VERSION 5.00
Begin VB.UserControl AutoCombo 
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ScaleHeight     =   1050
   ScaleWidth      =   3015
   Begin VB.ComboBox Combo 
      Height          =   315
      ItemData        =   "AutoCombo.ctx":0000
      Left            =   0
      List            =   "AutoCombo.ctx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   2475
   End
End
Attribute VB_Name = "AutoCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim blnAuto As Boolean 'Keeps the autocomplete functions from
                        'triggering the Change event

Private Sub UserControl_Resize()
    Combo.Width = UserControl.Width
    UserControl.Height = Combo.Height
End Sub

Private Sub combo_Change()
    Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    'don't do if no text or if change was made by autocomplete coding
    If Not blnAuto And Combo.Text <> "" Then
        'save the selection start point (cursor position)
        iStart = Combo.SelStart
        'get the part the user has typed (not selected)
        strPart = Left$(Combo.Text, iStart)
        For iLoop = 0 To Combo.ListCount - 1
            'compare each item to the part the user has typed,
            '"complete" with the first good match
            strItem = UCase$(Combo.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(Combo.Text) Then
                'partial match but not the whole thing.
                '(if whole thing, nothing to complete!)
                blnAuto = True
                Combo.SelText = Mid$(Combo.List(iLoop), iStart + 1) 'add on the new ending
                Combo.SelStart = iStart   'reset the selection
                Combo.SelLength = Len(Combo.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub combo_KeyDown(KeyCode As Integer, Shift As Integer)
    'Unless we watch out for it, backspace or delete will just delete
    'the selected text (the autocomplete part), so we delete it here
    'first so it doesn't interfere with what the user expects
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        Combo.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        combo_LostFocus
        'the following causes the item to be selected and
        'the cursor placed at the end:
        Combo.SelStart = Len(Combo.Text)
        
        'This would select the whole thing instead:
        'combo.SelLength = Len(combo.Text)
    
        'alternatively, you could move the focus to the next control here
    End If
End Sub

Private Sub combo_LostFocus()
    Dim iLoop As Integer
    'Match capitalization if item entered is one on the list
    If Combo.Text <> "" Then
        For iLoop = 0 To Combo.ListCount - 1
            If UCase$(Combo.List(iLoop)) = UCase$(Combo.Text) Then
                blnAuto = True
                Combo.Text = Combo.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Public Sub AddItem(ByVal Item As String, Optional Index As Integer)
    Combo.AddItem Item, Index
End Sub

Public Function List(ByVal Index As Integer) As String
    List = Combo.List(Index)
End Function

Public Function ItemData(ByVal Index As Integer) As Long
    ItemData = Combo.ItemData(Index)
End Function

Public Function ListIndex() As Long
    ListIndex = Combo.ListIndex
End Function

Public Function ListCount() As Long
    ListCount = Combo.ListCount
End Function

Public Sub Clear()
    Combo.Clear
End Sub


