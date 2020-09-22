VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ß's Word-Counter :P"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count Words"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Count these words."
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblWords 
      AutoSize        =   -1  'True
      Caption         =   "Click Count Words."
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1365
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCount_Click()
Dim strInput As String
Dim strWords() As String
Dim x As Integer
Dim y As Integer
Dim FoundWords As Integer

FoundWords = 0
strInput = Trim(txtInput.Text) 'strip any spaces from the beginning and end to speed up the search a bit

If Len(strInput) > 0 Then 'if there's anything left after we stripped spaces...
    strWords = Split(strInput, " ") 'split the "words" into an array
    For x = 0 To UBound(strWords)   'for every one of the "words" we found...
        If Len(strWords(x)) > 0 Then    'if there's actually something here then...
            'for y=97 to 122            'uncomment this block to only count "words" that contain the letters a-z
                'if instr(1,lcase(strwords(x)),chr(y))>0 then
                    FoundWords = FoundWords + 1 'update the number of words found
                    'exit for   'so we don't count the same word more than once :P
                'end if
            'next y
        End If
    Next x
End If
lblWords.Caption = FoundWords & " words found in text." 'update the GUI.
End Sub
