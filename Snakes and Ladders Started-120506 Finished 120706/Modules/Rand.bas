Attribute VB_Name = "Rand"

Public Function Generate(Low As Integer, High As Integer) As Integer
    Generate = ((High - Low) * Rnd + Low)
End Function
