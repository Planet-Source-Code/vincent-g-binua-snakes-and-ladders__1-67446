Attribute VB_Name = "SnakesAndLadders"
'===================================================================================
'This module consists of functions, procedures, events and logic of the game
'This depends entirely on the layout of the board

'----BOARDS LAYOUT------
'OBJECT     |      QTY
'Ladder     |       7
'Snake      |       7
'Squares    |       100

'Number of Players = 2
'Boards Dimension = 10x10

'----LADDERS--------------------------------------------------
'NO.|   Begins at Square | Climb at Square | *Level |  Column
'1  |           5        |      15         |    1   |    5
'2  |           9        |      12         |    1   |    9
'3  |           18       |      39         |    2   |    3
'4  |           27       |      48         |    3   |    8
'5  |           44       |      74         |    5   |    4
'6  |           67       |      84         |    7   |    7
'7  |           83       |      99         |    9   |    3
'-------------------------------------------------------------

'----SNAKES---------------------------------------------------
'NO.|   Begins at Square | Slides at Square | *Level |  Column
'1  |           25       |      4           |   3    |   5
'2  |           13       |      7           |   2    |   8
'3  |           69       |      48          |   7    |   9
'4  |           76       |      37          |   8    |   5
'5  |           79       |      61          |   8    |   2
'6  |           91       |      72          |   10   |   10
'7  |           94       |      75          |   10   |   7
'-------------------------------------------------------------


'------PLAYERS DIRECTION OF STEP AT A GIVEN LEVEL---------------------------
'LEVEL  |   DIRECTION
'  1    |      L->R
'  2    |      R->L
'  3    |      L->R
'  4    |      R->L
'  5    |      L->R
'  6    |      R->L
'  7    |      L->R
'  8    |      R->L
'  9    |      L->R
'  10   |      R->L
'
'Odd Numbered Level shows a movement from Left to Right
'Even Numbered Level shows a movement from Right to Left
'----------------------------------------------------------------------------

Public ActiveSquares() As Integer

'======================================================================
'This Module populates an array of square numbers that are active
'- active squares are squares that raises an event when a player
'stops into (e.g. Climbing the ladders, Sliding the snakes)
'This depends again on the layout of the board

Public Sub PopulateActiveSquares()
    
    ReDim ActiveSquares(1 To NUMBER_OF_ACTIVE_SQUARES)
    
    ActiveSquares(1) = 5
    ActiveSquares(2) = 9
    ActiveSquares(3) = 18
    ActiveSquares(4) = 27
    ActiveSquares(5) = 44
    ActiveSquares(6) = 67
    ActiveSquares(7) = 83
    ActiveSquares(8) = 25
    ActiveSquares(9) = 13
    ActiveSquares(10) = 69
    ActiveSquares(11) = 76
    ActiveSquares(12) = 79
    ActiveSquares(13) = 91
    ActiveSquares(14) = 94
    
End Sub
'======================================================================

