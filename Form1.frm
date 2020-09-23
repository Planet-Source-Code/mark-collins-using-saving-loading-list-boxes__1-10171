VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "List Box Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Load...."
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Data"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add item"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save...."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -====================-
'Created by: Mark Collins    Friday, July 28, 2000
'For the purpose of educating
'others or something.
' -====================-
Private Sub Timer1_Timer()

End Sub

Private Sub Command1_Click()
Dim SaveList As Long
    'make sure that if stuff F*cks up
    'that the prog doesnt crash.
    On Error Resume Next
    'makes directory$ equal to the path and name of our file
    directory$ = App.Path & "\DumbData.txt"
    'opens the file to write to it
    Open directory$ For Output As #1
    'this makes savelist& equal to every number between
    '0 and the number of items in the list minus 1
    For SaveList& = 0 To List1.ListCount - 1
        'prints the current listindex text into the file
        Print #1, List1.List(SaveList&)
    'goes to the next savelist number
    Next SaveList&
    'slaps the lid on that ole' file of ours!
    Close #1
    MsgBox "The List box was saved to " & directory$
    
End Sub

Private Sub Command2_Click()

'get the info we wanna  add to the listbox
crap$ = InputBox("Please enter the data you would like to add to the list box", "Add data...", "I saw a saw in a scene i saw")
List1.AddItem crap$
End Sub

Private Sub Command3_Click()
'check to make sure the index is atleast 0
If List1.ListIndex < 0 Then Exit Sub
'remove the selected item
List1.RemoveItem (List1.ListIndex)

End Sub

Private Sub Command4_Click()
Dim MyString As String
    'make sure that if stuff F*cks up
    'that the prog doesnt crash.
    On Error Resume Next
    'clears the listbox
    List1.Clear
    'make directory$ equal the dir of our saved list.
    directory$ = App.Path & "\DumbData.txt"
    'open the file
    Open directory$ For Input As #1
    'loop through all of the lines in the
    'file and add each separate line to the listbox
    While Not EOF(1)
        'puts the line out of the text file into mystring$
        Input #1, MyString$
        DoEvents
        List1.AddItem MyString$
    Wend
    'close the file!
    Close #1

MsgBox "File loaded"
End Sub

Private Sub Form_Load()

End Sub
