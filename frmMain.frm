VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic Blackjack"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar barProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ListBox lstResults 
      Height          =   2205
      Left            =   1800
      TabIndex        =   11
      Top             =   1080
      Width           =   5415
   End
   Begin VB.CommandButton cmdStartProgram 
      Caption         =   "Start Program"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtDeals 
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Text            =   "1000"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtStake 
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Text            =   "400"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtBetUnit 
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Text            =   "10"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtShufflePercent 
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Text            =   "60"
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.Slider sliDecks 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   6
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label lblPushes 
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblRemain 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblPlayed 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Hands Remain:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Hands Played:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblStake 
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblDealerWins 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblPlayerWins 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Stake:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Dealer Wins:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Player Wins:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label5 
      Caption         =   "Number of Deals"
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Stake"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Bet Unit $"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Shuffle %"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Decks"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Pushes:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intDecks(6, 13) As Integer
Dim intCardsRemainingCount As Integer
Dim intBetUnit As Integer
Dim intStake As Integer
Dim intCount As Integer
Dim intAceCount As Integer
Dim intFiveCount As Integer
Dim intDealerDownCard As Integer
Dim intDealerHand(10) As Integer
Dim intPlayerHand(10) As Integer
Dim intDealerPOS As Integer
Dim intPlayerPOS As Integer
Dim intDealerValue As Integer
Dim intPlayerValue As Integer
Dim intPlayerWins As Integer
Dim intDealerWins As Integer
Dim intHandsRemain As Integer
Dim bolPlayerStand As Boolean
Dim bolDealerStand As Boolean
Dim bolDoubleDown As Boolean
Dim bolPlayerBust As Boolean
Dim bolPlayerBlackJack As Boolean
Dim intPush As Integer

Private Sub YuccaCount(ByVal intCurrentCard As Integer)
'This routine calculates the count based on the Yucca Count. this sub receives
'a variable, when a call to this sub is made the byVal intCurrentCard requires
'that an integer be passed into this sub. If an integer is NOT passed the program
'will crash with an error.

If intCurrentCard >= 3 And intCurrentCard <= 7 Then
intCount = intCount + 1
    If intCurrentCard = 5 Then
    intFiveCount = intFiveCount + 1
    End If
ElseIf intCurrentCard >= 9 And intCurrentCard <= 13 Then
intCount = intCount - 1
ElseIf intCurrentCard = 1 Then
intAceCount = intAceCount + 1
End If

'lstResults.AddItem ("Count is: " & intCount)
'lstResults.AddItem ("Ace Count is: " & intAceCount)
'lstResults.AddItem ("Five Count is: " & intFiveCount)

End Sub

Private Sub Shuffle()
'This Subroutine Shuffles six decks of cards (
'
'In Blackjack suit is irrelivant as long as there are four of each card in each deck
'The array intDecks(6,13) means 6 decks, with 13 cards. when we set each card
'to 4 we are telling the computer that each of the 13 cards can appear a max
'of four times in each deck. Later we can program how many decks to chose from

'the intCardsRemainingCount tells us how many cards we can play before the
'dealer shuffles. It is based on a percentage that is set on the main form

lstResults.AddItem ("Shuffling Cards...")

intCardsRemainingCount = Int((sliDecks.Value * 52) * (txtShufflePercent / 100))
intCount = 0
intAceCount = 0
intFiveCount = 0

For I = 1 To 6
For J = 1 To 13
intDecks(I, J) = 4
Next J
Next I
End Sub

Private Sub cmdStartProgram_Click()
lstResults.Clear
lstResults.AddItem ("!!Starting New Simulation!!")
lstResults.AddItem ("Number of Decks: " & sliDecks.Value)
lstResults.AddItem ("Suffle Percentage: " & txtShufflePercent.Text)
lstResults.AddItem ("Bet Unit: " & txtBetUnit.Text)
lstResults.AddItem ("Stake: " & txtStake.Text)
lstResults.AddItem ("Number of Deals: " & txtDeals.Text)
lstResults.AddItem ("===================================")

intBetUnit = txtBetUnit.Text
intStake = txtStake.Text
intDealerWins = 0
intPlayerWins = 0
intPush = 0

Shuffle

barProgress.Value = 0
barProgress.Max = txtDeals.Text

For I = 1 To txtDeals.Text

'Shuffle Deck if out of cards
If intCardsRemainingCount <= 0 Then
Shuffle
End If

'New Game!! Set Variables to 0
intDealerPOS = 0
intPlayerPOS = 0
intDealerValue = 0
intPlayerValue = 0
intCount = 0
intAceCount = 0
intFiveCount = 0

For J = 0 To 10
intPlayerHand(J) = 0
intDealerHand(J) = 0
Next J

'Deal Innitial Hands, All other code handled in the Deal To Subs

lstResults.AddItem ("Begin new Hand")
bolPlayerStand = False
bolDealerStand = False
bolDoubleDown = False
bolPlayerBust = False
bolPlayerBlackJack = False

DealToPlayer
DealToDealer (True)
DealToPlayer
DealToDealer (False)

Do Until (bolPlayerStand = True)
DoEvents
DoPlayer
Loop

Do Until (bolDealerStand = True)
DoEvents
DoDealer
Loop

FindWinner

lblPlayerWins.Caption = intPlayerWins
lblDealerWins.Caption = intDealerWins
lblStake.Caption = intStake
lblPlayed.Caption = I
lblRemain.Caption = Int(txtDeals.Text - I)
lblPushes.Caption = intPush

DoEvents

'End Program if Player Runs out of money
If intStake < intBetUnit Then
lstResults.AddItem ("Player Out of money! Game Over!")
I = txtDeals.Text
End If
barProgress.Value = I
Next I

lstResults.AddItem ("Simulation Complete!")
barProgress.Value = 0
End Sub

Private Sub FindWinner()
'Please check logic on this sub, I think somewhere I have made a mistake...

If intPlayerValue <= 21 And (intPlayerValue > intDealerValue) And intDealerValue < 21 Then
PlayerWin
ElseIf bolPlayerBust = True Then
PlayerLose
ElseIf intPlayerValue <= 21 And intDealerValue > 21 Then
PlayerWin
ElseIf intPlayerValue < intDealerValue Then
PlayerLose
ElseIf intPlayerValue = intDealerValue Then
PlayerPush
ElseIf bolPlayerBlackJack = True Then
PlayerWin
End If

End Sub

Private Sub DealToDealer(ByVal bolDown As Boolean)
'bolDown will tell the computer to store the dealers card in the Down Variable
'So that it is not calculated in with the count or score until it is flipped over
Dim intChoseCard As Integer
Dim intChoseDeck As Integer
Dim bolPass As Boolean

bolPass = False

Do Until (bolPass = True)
intChoseDeck = Int((sliDecks.Value - 1 + 1) * Rnd + 1)
intChoseCard = Int((13 - 1 + 1) * Rnd + 1)

If intDecks(intChoseDeck, intChoseCard) > 0 Then
intDecks(intChoseDeck, intChoseCard) = intDecks(intChoseDeck, intChoseCard) - 1
YuccaCount (intChoseCard)
intCardsRemainingCount = intCardsRemainingCount - 1
bolPass = True
Else
bolPass = False
End If

Loop

'Set Number of Dealer Cards on the Table (Position in Array)
intDealerPOS = intDealerPOS + 1

'Actualy Deal the card, Set J-K to 10

If bolDown = True Then
'This will prevent the "unknown" card from being included in the count or the
'dealers value until it is revealed. The Player will not take Hole into consideration
'when making hit\stand descisions
intDealerDownCard = intChoseCard
intDealerHand(1) = 0
lstResults.AddItem ("Dealer Delt: Hole")
Else

If intChoseCard > 10 Then
lstResults.AddItem ("Dealer Delt: 10")
intDealerHand(intDealerPOS) = intDealerHand(intDealerPOS) + 10
ElseIf intChoseCard = 1 Then
lstResults.AddItem ("Dealer Delt: A")
'Gives Aces an initial Value of 11, we will check for a bust and set them
'to 1 if the ace being 11 is causing a bust later...
intDealerHand(intDealerPOS) = intDealerHand(intDealerPOS) + 11
Else
lstResults.AddItem ("Dealer Delt: " & intChoseCard)
intDealerHand(intDealerPOS) = intDealerHand(intDealerPOS) + intChoseCard
End If

End If
'Calculate the Value of the Hand
CheckDealerValue


End Sub

Private Sub CheckDealerValue()
intDealerValue = 0

'Calculates the Dealers value
For I = 1 To intDealerPOS
intDealerValue = intDealerValue + intDealerHand(I)
Next I

'Checks to make sure that Dealers hand is not over 21, if it is then
'it will check to see if there are any 11's (Aces) in the hand
'If there are then it will make them 1's instead and then it will check
'again
If intDealerValue > 21 Then
'Check for Aces
For I = 1 To intDealerPOS

    If intDealerValue > 21 And intDealerHand(I) = 11 Then
    intDealerHand(I) = 1
    intDealerValue = 0
        'if an 11 ace is found then we need to recalculate the score after 11 = 1
        For J = 1 To intDealerPOS
        intDealerValue = intDealerValue + intDealerHand(J)
        Next J
        
    End If
    
Next I

End If


If intDealerDownCard = 11 And intDealerHand(2) = 10 And intDealerPOS = 2 Then
intDealerHand(1) = 11
intDealerDownCard = 0
YuccaCount (11)
DealerBlackjack
ElseIf intDealerValue > 21 Then
DealerBust
End If
lstResults.AddItem ("Dealer Has: " & intDealerValue)
End Sub

Private Sub DealerBust()
lstResults.AddItem ("Dealer Busts...")
bolDealerStand = True
End Sub

Private Sub DealerBlackjack()
lstResults.AddItem ("Dealer Has Blackjack...")
bolDealerStand = True
End Sub

Private Sub DealToPlayer()
Dim intChoseCard As Integer
Dim intChoseDeck As Integer
Dim bolPass As Boolean

bolPass = False

Do Until (bolPass = True)
intChoseDeck = Int((sliDecks.Value - 1 + 1) * Rnd + 1)
intChoseCard = Int((13 - 1 + 1) * Rnd + 1)

If intDecks(intChoseDeck, intChoseCard) > 0 Then
intDecks(intChoseDeck, intChoseCard) = intDecks(intChoseDeck, intChoseCard) - 1
YuccaCount (intChoseCard)
intCardsRemainingCount = intCardsRemainingCount - 1
bolPass = True
Else
bolPass = False
End If

Loop

'Set Number of Player Cards on the Table (Position in Array)
intPlayerPOS = intPlayerPOS + 1

'Actualy Deal the card, Set J-K to 10
If intChoseCard > 10 Then
lstResults.AddItem ("Player Delt: 10")
intPlayerHand(intPlayerPOS) = intPlayerHand(intPlayerPOS) + 10
ElseIf intChoseCard = 1 Then
lstResults.AddItem ("Player Delt: A")
'Gives Aces an initial Value of 11, we will check for a bust and set them
'to 1 if the ace being 11 is causing a bust later...
intPlayerHand(intPlayerPOS) = intPlayerHand(intPlayerPOS) + 11
Else
lstResults.AddItem ("Player Delt: " & intChoseCard)
intPlayerHand(intPlayerPOS) = intPlayerHand(intPlayerPOS) + intChoseCard
End If

'Calculate the Value of the Hand
CheckPlayerValue



End Sub

Private Sub CheckPlayerValue()

intPlayerValue = 0

'Calculates the players value
For I = 1 To intPlayerPOS
intPlayerValue = intPlayerValue + intPlayerHand(I)
Next I

'Checks to make sure that players hand is not over 21, if it is then
'it will check to see if there are any 11's (Aces) in the hand
'If there are then it will make them 1's instead and then it will check
'again
If intPlayerValue > 21 Then
'Check for Aces
For I = 1 To intPlayerPOS

    If intPlayerValue > 21 And intPlayerHand(I) = 11 Then
    intPlayerHand(I) = 1
    intPlayerValue = 0
        'if an 11 ace is found then we need to recalculate the score after 11 = 1
        For J = 1 To intPlayerPOS
        intPlayerValue = intPlayerValue + intPlayerHand(J)
        Next J
        
    End If
    
Next I

End If

'intPlayerPOS = 2 indicates 2nd card drawn, so a 21 MUST be a Blackjack
If intPlayerValue = 21 And intPlayerPOS = 2 Then
PlayerBlackjack
End If

lstResults.AddItem ("Player Has: " & intPlayerValue)


End Sub

Private Sub DoPlayer()
'Play will continue as if the player is following BASIC STRATEGY. The Count
'Strategy can be added in later. (Also needs basic for multideck)
'Also note that program was written a bit in longhand so as to make it
'a bit easier to read. It is a bit sloppy but will not affect performance.

'**TODO: ADD ACE AND SPLIT STRATEGIES HERE ALSO ADD MULTIDECK BASIC STRATEGY**
CheckPlayerValue
If bolPlayerStand = False Then
If intPlayerValue > 21 Then
PlayerBust
ElseIf intPlayerValue >= 17 Then
PlayerStand
ElseIf (intPlayerValue > 12 And intPlayerValue < 17) And (intDealerValue > 1 Or intDealerValue < 7) Then
PlayerStand
ElseIf (intPlayerValue > 12 And intPlayerValue < 17) And (intDealerValue > 6 And intDealerValue < 12) Then
PlayerHit
ElseIf intPlayerValue = 12 And (intDealerValue = 2 Or intDealerValue = 3) Then
PlayerHit
ElseIf intPlayerValue = 12 And (intDealerValue > 3 And intDealerValue < 7) Then
PlayerStand
ElseIf intPlayerValue = 12 And (intDealerValue > 7 And intDealerValue < 12) Then
PlayerHit
'<> means NOT equal to
ElseIf intPlayerValue = 11 And (intDealerValue <> 10 And intDealerValue <> 11) Then
PlayerDouble
ElseIf intPlayerValue = 11 And (intDealerValue = 10 And intDealerValue = 11) Then
PlayerHit
ElseIf intPlayerValue = 10 And (intDealerValue <> 10 And intDealerValue <> 11) Then
PlayerDouble
ElseIf intPlayerValue = 10 And (intDealerValue = 10 And intDealerValue = 11) Then
PlayerHit
ElseIf intPlayerValue = 9 And (intDealerValue > 1 And intDealerValue < 7) Then
PlayerDouble
ElseIf intPlayerValue = 9 And (intDealerValue > 6 And intDealerValue < 12) Then
PlayerHit
Else
PlayerHit
End If
End If
'Never Take insurance! Also note that the Splitting and Soft hand Has NOT YET BEEN ADDED
CheckPlayerValue
End Sub

Private Sub DoDealer()
'Flip Hole Card
intDealerHand(1) = intDealerDownCard
CheckDealerValue
YuccaCount (intDealerDownCard)

If bolPlayerBlackJack = False And bolPlayerBust = False Then
If intDealerValue < 16 Then
DealerHit
Else
DealerStand
End If
End If

If bolPlayerBlackJack = True Then
DealerStand
End If

If bolPlayerBust = True Then
DealerStand
End If


End Sub

Private Sub DealerHit()
lstResults.AddItem ("Dealer Hits...")
DealToDealer (False)
End Sub

Private Sub DealerStand()
lstResults.AddItem ("Dealer Stands...")
bolDealerStand = True
End Sub

Private Sub PlayerHit()
'Basicaly return to the deal to player sub
lstResults.AddItem ("Player Hitting...")
DealToPlayer
End Sub

Private Sub PlayerBust()
lstResults.AddItem ("Player Busts...")
bolPlayerBust = True
bolPlayerStand = True
End Sub

Private Sub PlayerLose()
If bolDoubleDown = True Then
intStake = intStake - (intBetUnit * 2)
Else
intStake = intStake - intBetUnit
End If
bolPlayerStand = True
intDealerWins = intDealerWins + 1
End Sub

Private Sub PlayerPush()
lstResults.AddItem ("Player Push...")
intPush = intPush + 1
End Sub

Private Sub PlayerWin()
If bolDoubleDown = True Then
intStake = intStake + (intBetUnit * 2)
ElseIf bolPlayerBlackJack = True Then
'Assume 3 to 5 blackjack
intStake = intStake + (intBetUnit * 2.5)
Else
intStake = intStake + intBetUnit
End If
intPlayerWins = intPlayerWins + 1
End Sub

Private Sub PlayerBlackjack()
lstResults.AddItem ("Player Has Blackjack...")
bolPlayerBlackJack = True
bolPlayerStand = True
End Sub

Private Sub PlayerDouble()
lstResults.AddItem ("Player Doubles")
PlayerHit
bolDoubleDown = True
bolPlayerStand = True
End Sub

Private Sub PlayerStand()
lstResults.AddItem ("Player Stands...")
bolPlayerStand = True
End Sub

Private Sub Form_Load()
Randomize
'Always use Randomize, Randomize will generate new seed, otherwise
'random numbers will always be the same.
End Sub

