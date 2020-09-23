VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Thirteen Out"
   ClientHeight    =   4500
   ClientLeft      =   3045
   ClientTop       =   2025
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8775
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Picture1 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6960
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblLeft 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Cards Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   28
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1065
   End
   Begin VB.Image NewCard 
      Enabled         =   0   'False
      Height          =   1440
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   120
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   27
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   26
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   25
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   24
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   23
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   22
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   21
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   20
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   19
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   18
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   17
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   16
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   15
      Left            =   840
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   14
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   13
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   12
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   11
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   10
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   9
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   8
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   7
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   6
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   5
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   4
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   3
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   2
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   1
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1065
   End
   Begin VB.Image Pic 
      Height          =   1425
      Index           =   0
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1065
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuGame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuRules 
         Caption         =   "&Rules"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cardVal(1 To 52) As Integer 'Holds the value of the card
Dim NoClick As Integer          'Number of click to compare if 13
Dim ClickVal(1 To 2) As Integer 'Holds the "index" you have clicked
Dim LastCard As Integer         'Holds the last "index" to check for value
Dim CardCurPos As Integer       'Holds the current "index" to check for value
Dim CardsGone As Integer        'Holds how many cards are gone from the pyramide

'*** Shuffle and initialze a new game***'
Private Sub Shuffle()
Dim i, x, tmpVal As Integer
Dim bolkoll As Boolean
'***Some Cleaning up***'
LastCard = 52
CardCurPos = 29
    For i = 1 To 52
        cardVal(i) = Empty
    Next
    
    For i = 0 To 27
    Pic(i).Enabled = True
    Pic(i).Visible = True
    PlacePicHolders (i)
    Next
    Pic(28).Picture = LoadPicture("")
'***End Cleaning up***'

'***Randomly fill the cardval array with cardvalues***'
For i = 0 To 51
    bolkoll = False
    Randomize
    tmpVal = Int((Rnd * 52) + 1) 'Genrates a random number between 1-52
    For x = 0 To i 'Check if the random number (tmpval) already exitst)
        If tmpVal = cardVal(x + 1) Then 'Loop the same loop again
            i = i - 1
            bolkoll = True
            Exit For 'Not necesarry to loop any further
        End If
    Next
    If bolkoll = False Then 'Set the random number (tmpVal) to the current pos in cardval
        cardVal(i + 1) = tmpVal
    End If
Next
End Sub
'***Place the cards there they should be***'
Private Sub PlacePicHolders(Index As Integer)
Select Case Index 'Set Top Pos
    Case 0
    Pic(0).Top = 600
    Case 1, 2
        Pic(Index).Top = 960
    Case 3 To 5
        Pic(Index).Top = 1320
    Case 6 To 9
        Pic(Index).Top = 1680
    Case 10 To 14
        Pic(Index).Top = 2040
    Case 15 To 20
        Pic(Index).Top = 2400
    Case 21 To 27
        Pic(Index).Top = 2760
End Select

Select Case Index ' Set Left Pos
    Case 0, 4, 12, 24
    Pic(Index).Left = 3840
    Case 1, 7, 17
    Pic(Index).Left = 3240
    Case 2, 8, 18
    Pic(Index).Left = 4440
    Case 3, 11, 23
    Pic(Index).Left = 2640
    Case 5, 13, 25
    Pic(Index).Left = 5040
    Case 6, 16
    Pic(Index).Left = 2040
    Case 9, 19
    Pic(Index).Left = 5640
    Case 10, 22
    Pic(Index).Left = 1440
    Case 14, 26
    Pic(Index).Left = 6240
    Case 15
    Pic(Index).Left = 840
    Case 21
    Pic(Index).Left = 240
    Case 26
    Pic(Index).Left = 6240
    Case 27
    Pic(Index).Left = 7440
End Select
End Sub

'***Get the real value (0-12 there 0 is king)***'
Private Function RealVal(ByVal CheckVal As Integer) As Integer
'Divide the value checkval (cardval(?)) with 13, the rest is the real value
RealVal = Int(CheckVal Mod 13)
End Function

'***When one card is gone Rearrange the values in cardval(?)***'
Private Sub ReArr(BeginArr As Integer) 'BeginArr is where to start moving the values
Dim i As Integer
For i = BeginArr To LastCard - 1
cardVal(i) = cardVal(i + 1)
Next
LastCard = LastCard - 1
CardCurPos = CardCurPos - 1
End Sub

'*** Flip the cards over and move the last cards (less than three) in the beginning***'
Private Function Checkflip(intRest As Integer) As Integer 'Intrest = how many cards left
Dim i, Last, NextLast As Integer
Last = cardVal(LastCard) 'Temporary store the last value
NextLast = cardVal(LastCard - 1) 'Temporary store the second last value
For i = LastCard To 28 Step -1
cardVal(i) = cardVal(i - intRest) 'Move the value no of steps in intrest
Next

Select Case intRest
    Case 1
        cardVal(29) = Last 'Move the old last card first in deck
    Case 2
        cardVal(29) = Last 'Move the old last card first in deck
        cardVal(30) = NextLast 'Move the old second last card second in deck
End Select

End Function

Private Sub mnuExit_Click()
Unload Me
End Sub

'***Start the shuffle and get the cards***'
Private Sub mnuGame_Click()
Dim i As Integer
Shuffle
For i = 1 To 28
    Pic(i - 1).Picture = LoadPicture(App.Path & "\card" & cardVal(i) & ".gif")
Next
NewCard.Enabled = True
End Sub
'***Show the rules***'
Private Sub mnuRules_Click()
MsgBox "The rules are simple, Click two cards," & vbCrLf & _
"if the value is 13 togheter they are gone" & vbCrLf & _
"You can click cards in both the pyramide and outside it." & vbCrLf & _
"But only if they are free." & vbCrLf & _
"You get new cards by clicking the deck with the backside up." & vbCrLf & _
"You throw three cards at the time, the last cards if they" & vbCrLf & _
"are less then 3 goes to beginning of the deck." & vbCrLf & _
"You may 'Flip' the deck as long as you want and can get" & vbCrLf & _
"some cards out from the pyramide." & vbCrLf & _
"Try to get as many cards as possible out from the pyramide.", , "Thirteen out Rules"

End Sub

'***Throw the cards***'
Private Sub NewCard_Click()
If CardCurPos = LastCard Then CardCurPos = 32 'If it is the lastcard set the pos to first card
CardCurPos = CardCurPos + 3 'Throw three cards at a time
If CardCurPos > LastCard Then ' if cardcurpos is bigger than lastcard
    Checkflip (CardCurPos - LastCard) 'Divied cardcurpos with lastcard and get a rest
    CardCurPos = 32 'set cardcurpos to the third card in the newly flipped deck
End If
PlaySound ("card")
Pic(28).Picture = LoadPicture(App.Path & "\card" & cardVal(CardCurPos) & ".gif")
NoClick = 0 'Take away possible clicks
End Sub

'*** compare the values of the 2 clicks if it is 13***'
Private Sub Pic_Click(Index As Integer)
Dim i As Integer

If FreeCard(Index) = True Then 'if the clicked card is free then
    NoClick = NoClick + 1 'Increase Number of clicks with one
    If Index = 28 Then 'if the card is "besides the pyramide
        ClickVal(NoClick) = CardCurPos 'Set clickval to the current pos in the deck
    Else 'if the card is in the pyramide
        ClickVal(NoClick) = Index + 1 'set clickval to the value in the valarr
    End If
    
        If RealVal(cardVal(ClickVal(1))) = 0 Then 'if it is a king on the first click
            PlaySound ("card")
            NoClick = 0
            If Index = 28 Then 'Outside the pyramide
                ReArr (CardCurPos) 'Rearrange the values in the array
                If CardCurPos < 29 Then 'if it is the last card in the showed pics from deck
                    CardCurPos = 29
                    Pic(28).Picture = LoadPicture("") 'show nothing
                Else 'show the card under the one you throw away
                    Pic(28).Picture = LoadPicture(App.Path & _
                                        "\card" & cardVal(CardCurPos) & ".gif")
                End If
            Else 'If the card is on the pyramide
                Pic(Index).Move Picture1.Left, Picture1.Top 'move it away from the pyramide
                Pic(Index).Visible = False 'Hide it
                Pic(Index).Enabled = False 'No clicking possible
                CardsGone = CardsGone + 1 'increase cardsgone with one
            End If
        End If

    If NoClick = 2 Then 'Compare NoClick(1) and NoClick(2) if it is 13
        NoClick = 0
        If RealVal(cardVal(ClickVal(1))) + RealVal(cardVal(ClickVal(2))) = 13 Then
            PlaySound ("card")
        For i = 1 To 2
            If ClickVal(i) < 29 Then 'In the pyramide
                Pic(ClickVal(i) - 1).Move Picture1.Left, Picture1.Top 'Move it from the pyramide
                Pic(ClickVal(i) - 1).Enabled = False 'No clicks possible
                Pic(ClickVal(i) - 1).Visible = False 'Hide it
            Else 'Not in the pyramide
                ReArr (ClickVal(i)) 'Rearrange the values in the array
                If CardCurPos < 29 Then 'if the current pos in the deck is lower than 29
                    CardCurPos = 29 'Set the current pos in the deck to 29
                    Pic(28).Picture = LoadPicture("") 'Show no picture
                Else 'In other case show the correct card under the thrown away card
                    Pic(28).Picture = LoadPicture(App.Path & _
                                        "\card" & cardVal(CardCurPos) & ".gif")
                End If
            End If
        Next
            'Increase the number of cards gone
            If ClickVal(1) <= 28 Then CardsGone = CardsGone + 1
            If ClickVal(2) <= 28 Then CardsGone = CardsGone + 1

        End If
    End If
Else
NoClick = 0
End If
lblLeft.Caption = "Cards left " & 28 - CardsGone 'Show how many cards there are left in the pyramide
End Sub

'***Check if the clicked card is free***'
Private Function FreeCard(Index As Integer) As Boolean
Dim CardRow As Integer 'Position in the pyramide
Select Case Index
    Case 0 'First row in the pyramide (Topcard)
        CardRow = 1
    Case 1, 2 'second row in the pyramide
        CardRow = 2
    Case 3 To 5 'third row in the pyramide
        CardRow = 3
    Case 6 To 9 'Fourth row in the pyramide
        CardRow = 4
    Case 10 To 14 'Fifth row in the pyramide
        CardRow = 5
    Case 15 To 20 'Sixth row in the pyramide
        CardRow = 6
End Select

Select Case Index ' if the cards next under is gone...
    Case 0 To 20
        If Pic(Index + CardRow).Left = Picture1.Left And _
            Pic(Index + CardRow + 1).Left = Picture1.Left Then
            FreeCard = True
        End If
    Case Else ' all other cards are free
        FreeCard = True
End Select
End Function


