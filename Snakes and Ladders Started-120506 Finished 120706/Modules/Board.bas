Attribute VB_Name = "Board"
Public Sub MoveLeft(Player As PictureBox, Steps As Integer)
    Player.Left = Player.Left + ((2 * Steps) * (STEP_LENGTH + Player.Width / 2))
End Sub

Public Sub Climb(Player As PictureBox, Climbs As Integer)
    Player.Top = Player.Top - (Climbs * (CLIMB_LENGTH + Player.Height / 2))
End Sub
