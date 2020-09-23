VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "show value"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   510
      Width           =   1785
   End
   Begin prjAutoComplete.AutoCombo AutoCombo1 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox AutoCombo1.List(AutoCombo1.ListIndex)
End Sub

Private Sub Form_Load()
    'add a bunch of items.  autocombo1.combo's Sorted property is
    'True so they will end up in order
    AutoCombo1.AddItem "Apples"
    AutoCombo1.AddItem "Oranges"
    AutoCombo1.AddItem "Bananas"
    AutoCombo1.AddItem "Pears"
    AutoCombo1.AddItem "Peaches"
    AutoCombo1.AddItem "Pineapples"
    AutoCombo1.AddItem "Grapes"
    AutoCombo1.AddItem "Blueberries"
    AutoCombo1.AddItem "Raspberries"
    AutoCombo1.AddItem "Blackberries"
    AutoCombo1.AddItem "Papaya"
    AutoCombo1.AddItem "Kiwi"
    AutoCombo1.AddItem "Watermelon"
    AutoCombo1.AddItem "Guava"
End Sub



