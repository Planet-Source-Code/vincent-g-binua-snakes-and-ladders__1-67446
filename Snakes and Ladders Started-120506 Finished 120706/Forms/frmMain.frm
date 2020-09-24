VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snake and Ladders"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8565
      TabIndex        =   0
      Top             =   6525
      Width           =   915
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7575
      Top             =   2925
   End
   Begin MSComctlLib.ImageList imglstDice 
      Left            =   7140
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9872
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E4F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13276
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1811B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrDice 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   7635
      Top             =   2355
   End
   Begin MSComctlLib.ImageList imglstBoard 
      Left            =   7680
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   540
      ImageHeight     =   450
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CFFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B19D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B941
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBoard 
      Height          =   6750
      Left            =   135
      ScaleHeight     =   11.8
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   14.261
      TabIndex        =   1
      Top             =   120
      Width           =   8145
      Begin VB.PictureBox PlayerTwo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   270
         ScaleHeight     =   0.318
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   0.318
         TabIndex        =   5
         Top             =   6420
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.PictureBox PlayerOne 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   270
         ScaleHeight     =   0.318
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   0.318
         TabIndex        =   2
         Top             =   6135
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgBoard 
         Height          =   6750
         Left            =   -15
         Picture         =   "frmMain.frx":524FF
         Top             =   -15
         Width           =   8100
      End
   End
   Begin VB.Image imgNotification 
      Height          =   3585
      Left            =   8385
      Picture         =   "frmMain.frx":65A50
      Top             =   1830
      Width           =   1245
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   8415
      X2              =   9660
      Y1              =   5550
      Y2              =   5550
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   8415
      X2              =   9645
      Y1              =   5565
      Y2              =   5565
   End
   Begin VB.Image imgPointer 
      Height          =   480
      Left            =   8310
      Picture         =   "frmMain.frx":6B9AC
      Top             =   330
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   8400
      X2              =   9630
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   8400
      X2              =   9645
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   390
      Left            =   8820
      Top             =   1185
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   390
      Left            =   8820
      Top             =   360
      Width           =   465
   End
   Begin VB.Label lblSecondPlayer 
      Caption         =   "Second Player"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      TabIndex        =   4
      Top             =   930
      Width           =   1245
   End
   Begin VB.Label lblFirstPlayer 
      Caption         =   "First Player"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8550
      TabIndex        =   3
      Top             =   90
      Width           =   990
   End
   Begin VB.Image imgDice 
      Height          =   585
      Left            =   8760
      Top             =   5760
      Width           =   555
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim index As Integer
Dim ActivePlayer As String
Dim FirstPlayerTopDistances(1 To 10) As Double
Dim SecondPlayerTopDistances(1 To 10) As Double
Dim PlayersLeftDistances(1 To 10) As Double

Private Sub RollDie()
    tmrDice.Enabled = True
End Sub

Private Sub MovePlayer()
    tmrMove.Enabled = True
End Sub

Private Sub cmdRoll_Click()
    RollDie
End Sub

Private Sub SetActivePlayer(Player As String)
    ActivePlayer = Player
    
    If imgPointer.Visible = True Then
        If Player = "PlayerOne" Then
            imgPointer.Top = Constants.TOP_POINTER_FIRST
            'lblFirstPlayer.Enabled = True
            'lblSecondPlayer.Enabled = False
        Else
            imgPointer.Top = Constants.TOP_POINTER_SECOND
            'lblFirstPlayer.Enabled = False
            'lblSecondPlayer.Enabled = True
        End If
    End If
End Sub

Private Function GetActivePlayer() As String
    GetActivePlayer = ActivePlayer
End Function

Private Sub tmrDelay_Timer()
    tmrDelay.Interval = tmrDelay.Interval + 1

    If tmrDelay.Interval = 100 Then
        tmrDelay.Interval = 1
        tmrDelay.Enabled = False
    End If
End Sub

Private Sub ShowPlayers()
    PlayerOne.Visible = True
    PlayerTwo.Visible = True
End Sub

Private Sub ShowPointers()
    imgPointer.Visible = True
End Sub

Private Sub EnableDie()
    cmdRoll.Enabled = True
End Sub

Private Sub EnablePlayers()
    lblFirstPlayer.Enabled = True
    lblSecondPlayer.Enabled = True
End Sub

Private Sub EnableBoard()
    imgBoard.Picture = imglstBoard.ListImages(2).Picture
    imgNotification.Picture = imglstBoard.ListImages(3).Picture
End Sub

Private Sub SetPlayersInitialPosition()
    PlayerOne.Left = Constants.PLAYERS_INITIAL_LEFT_POSITION
    PlayerTwo.Left = Constants.PLAYERS_INITIAL_LEFT_POSITION
    PlayerOne.Top = Constants.FIRSTPLAYER_1LDISTANCE_FROM_TOP
    PlayerTwo.Top = Constants.SECONDPLAYER_1LDISTANCE_FROM_TOP
End Sub

Private Sub PopulateArraysForDistances()
    Dim FirstPlayerTopDistance As Double
    Dim SecondPlayerTopDistance As Double
    Dim PlayersLeftDistance As Double
    Dim ClimbRatio As Double
    Dim StepRatio As Double
    
    FirstPlayerTopDistance = PlayerOne.Top
    SecondPlayerTopDistance = PlayerTwo.Top
    
    'since both yields the same values, we may choose any of the players
    PlayersLeftDistance = PlayerOne.Left
    
    ClimbRatio = CLIMB_LENGTH + PlayerOne.Height / 2
    StepRatio = STEP_LENGTH + PlayerOne.Width / 2
    
    For i = 1 To 10
        FirstPlayerTopDistances(i) = FirstPlayerTopDistance
        SecondPlayerTopDistances(i) = SecondPlayerTopDistance
        
        PlayersLeftDistances(i) = PlayersLeftDistance
        
        FirstPlayerTopDistance = FirstPlayerTopDistance - ClimbRatio
        SecondPlayerTopDistance = SecondPlayerTopDistance - ClimbRatio
               
        PlayersLeftDistance = PlayersLeftDistance + (2 * StepRatio)
    Next i
End Sub

Private Sub Initialize()
    Randomize
    imgBoard.Picture = imglstBoard.ListImages(1).Picture
    SetActivePlayer "PlayerOne"
    PopulateArraysForDistances
    PopulateActiveSquares
    SetPlayersInitialPosition
End Sub

Private Function IsSquareActive(Player As PictureBox) As Boolean
    For i = 1 To UBound(ActiveSquares)
        If ActiveSquares(i) = GetSquare(Player) Then
            IsSquareActive = True
            Exit Function
        Else
            IsSquareActive = False
        End If
    Next i
End Function

Private Sub Command1_Click()
    GoToSquare 100, PlayerOne
End Sub

Private Sub Form_Load()
    Initialize
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNewGame_Click()
    SetActivePlayer "PlayerOne"
    SetPlayersInitialPosition
    
    ShowPlayers
    ShowPointers
    EnableDie
    EnablePlayers
    EnableBoard
    
    mnuNewGame.Enabled = False
End Sub

Private Sub tmrDice_Timer()
    Dim i As Double

    index = Rand.Generate(1, 5)
    imgDice.Picture = imglstDice.ListImages(index).Picture
    
    '============================================================
    'This gives the effect of image transition from fast to slow
    'the greater the interval, the slower the transition
    If tmrDice.Interval < 50 Then
        tmrDice.Interval = tmrDice.Interval + 5
    Else
        If tmrDice.Interval < 100 Then
            tmrDice.Interval = tmrDice.Interval + 10
        Else
            
            'Restore default interval of tmrDice
            tmrDice.Interval = 5
            tmrDice.Enabled = False
            
            MovePlayer
        End If
    End If
    '============================================================
End Sub

Public Sub GoToSquare(SquareNumber As Integer, _
                       Player As PictureBox, Optional index As Integer)
    
    Dim SquareNumberPlusDice As Integer
       
    If (SquareNumber + index) > 100 Then
        SquareNumberPlusDice = 100 - (SquareNumber + index) Mod 100
    Else
        SquareNumberPlusDice = SquareNumber + index
    End If
    
    If GetActivePlayer = "PlayerTwo" Then
        PlayerTwo.Move PlayersLeftDistances(GetCol(SquareNumberPlusDice)), _
                       SecondPlayerTopDistances(GetRow(SquareNumberPlusDice))
    Else
        PlayerOne.Move PlayersLeftDistances(GetCol(SquareNumberPlusDice)), _
                       FirstPlayerTopDistances(GetRow(SquareNumberPlusDice))
    End If

    
End Sub

Public Sub RaiseEventsOnActiveSquares(SquareNumber As Integer, Player As PictureBox)
    Select Case SquareNumber
        Case 5
            GoToSquare 15, Player
            PlaySoundX "ladder2.wav"
        Case 9
            GoToSquare 12, Player
            PlaySoundX "ladder2.wav"
        Case 18
            GoToSquare 39, Player
            PlaySoundX "ladder2.wav"
        Case 27
            GoToSquare 48, Player
            PlaySoundX "ladder2.wav"
        Case 44
            GoToSquare 84, Player
            PlaySoundX "ladder2.wav"
        Case 67
            GoToSquare 74, Player
            PlaySoundX "ladder2.wav"
        Case 83
            GoToSquare 99, Player
            PlaySoundX "ladder2.wav"
        Case 25
            GoToSquare 4, Player
            PlaySoundX "snakehiss.wav"
        Case 13
            GoToSquare 7, Player
            PlaySoundX "snakehiss.wav"
        Case 69
            GoToSquare 48, Player
            PlaySoundX "snakehiss.wav"
        Case 76
            GoToSquare 37, Player
            PlaySoundX "snakehiss.wav"
        Case 79
            GoToSquare 61, Player
            PlaySoundX "snakehiss.wav"
        Case 91
            GoToSquare 72, Player
            PlaySoundX "snakehiss.wav"
        Case 94
            GoToSquare 75, Player
            PlaySoundX "snakehiss.wav"
    End Select
    
End Sub

Private Sub tmrMove_Timer()
    tmrMove.Interval = tmrMove.Interval + 1
    
    '============================================================
    'After the Die has been rolled,
    'Show a little delay before moving the players on the board
    If tmrMove.Interval > 30 Then
        'Restore default interval of tmrMove
        tmrMove.Interval = 1
        tmrMove.Enabled = False
        Debug.Print GetSquare(PlayerOne)
        If GetActivePlayer = "PlayerTwo" Then
            GoToSquare GetSquare(PlayerTwo), PlayerTwo, index
            
            If IsSquareActive(PlayerTwo) = True Then
                RaiseEventsOnActiveSquares GetSquare(PlayerTwo), PlayerTwo
            End If
            
            SetActivePlayer "PlayerOne"
        Else
            GoToSquare GetSquare(PlayerOne), PlayerOne, index
            
            If IsSquareActive(PlayerOne) = True Then
                RaiseEventsOnActiveSquares GetSquare(PlayerOne), PlayerOne
            End If
            
            SetActivePlayer "PlayerTwo"
        End If
        
        If GetSquare(PlayerOne) = 100 Then
            PlaySoundX "finish.wav"
            imgNotification.Picture = imglstBoard.ListImages(4).Picture
            cmdRoll.Enabled = False
            mnuNewGame.Enabled = True
        Else
            If GetSquare(PlayerTwo) = 100 Then
                PlaySoundX "finish.wav"
                imgNotification.Picture = imglstBoard.ListImages(5).Picture
                cmdRoll.Enabled = False
                mnuNewGame.Enabled = False
            End If
        End If
    End If
    '=======================================================================
End Sub

'============================================================================
'Returns the level or row where the player (used as a parameter) stays
Private Function GetLevel(Player As PictureBox) As Integer
    Dim i As Integer
    
    If Player.Name = "PlayerOne" Then
        For i = 1 To 10
            If Round(PlayerOne.Top) = Round(FirstPlayerTopDistances(i)) Then
                GetLevel = i
                Exit Function
            End If
        Next i
    Else
        For i = 1 To 10
            If Round(PlayerTwo.Top) = Round(SecondPlayerTopDistances(i)) Then
                GetLevel = i
                Exit Function
            End If
        Next i
    End If
End Function
'============================================================================

'============================================================================
'Returns the column of a player used as a parameter
Private Function GetColumn(Player As PictureBox) As Integer
    Dim i As Integer
    
    For i = 1 To 10
        If Round(Player.Left) = Round(PlayersLeftDistances(i)) Then
            GetColumn = i
            Exit Function
        End If
    Next i
End Function
'============================================================================

'============================================================================
'Determines the row where a SquareNumber resides
Private Function GetRow(SquareNumber As Integer) As Integer
    If SquareNumber Mod 10 = 0 Then
        GetRow = SquareNumber / 10
    Else
        GetRow = ((SquareNumber - (SquareNumber Mod 10)) / 10) + 1
        Debug.Print SquareNumber / 10 + 1
    End If
End Function
'============================================================================

'============================================================================
'Determines the column where a SquareNumber resides
Private Function GetCol(SquareNumber As Integer) As Integer
    Dim column As Integer
    
    If GetRow(SquareNumber) Mod 2 = 0 Then
        column = (GetRow(SquareNumber) * 10) - SquareNumber + 1
    Else
        column = (GetRow(SquareNumber) * 10) - (10 + SquareNumber)
    End If
    
    GetCol = Abs(column)
'============================================================================
End Function

'===========================================================================
'Returns the number of square where the player stays
Private Function GetSquare(Player As PictureBox) As Integer
    Dim row As Integer
    Dim column As Integer
    
    row = GetLevel(Player)
    column = GetColumn(Player)
    
    If row Mod 2 = 0 Then
        GetSquare = (row * 10) - (column - 1)
    Else
        GetSquare = (row * 10) - (10 - column)
    End If
End Function
'============================================================================

Public Function FileExists(FullFileName) As Boolean

    ' Passed a filename (with path) returns
    ' True if the file exists, False if not.

    Dim s As String
    
    s = Dir(FullFileName)
    
    If s = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Private Sub PlaySoundX(filename As String)

' If sound is enabled and filename exists,
' play the specified sound.

    filename = App.Path & "\.." & "\Sounds\" & filename
    
    If FileExists(filename) Then
        PlaySound filename, SND_ASYNC Or SND_FILENAME 'SND_ASYNC Or SND_FILENAME
    End If

End Sub


