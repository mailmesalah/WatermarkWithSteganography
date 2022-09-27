Attribute VB_Name = "MSharedCodes"

' Translate a password into an offset value.
Public Function NumericPassword(ByVal password As String) As Long
Dim Value As Long
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
Dim strLen As Integer

    ' Initialize the shift values to different
    ' non-zero values.
    shift1 = 3
    shift2 = 17

    ' Process the message.
    strLen = Len(password)
    For i = 1 To strLen
        ' Add the next letter.
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function

' Pick an unused (r, c, pixel) combination.
' Will try to add random generated row and column with red or green or blue pixel, if already added it will generate an error
' The error is ignored and will try the random check again untill an unused combination of row colum and pixel is added to the collection
Public Sub PickPosition(ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer, ByRef r As Integer, ByRef c As Integer, ByRef pixel As Integer)
Dim position_code As String

    On Error Resume Next
    Do
        ' Pick a position.
        r = Int(Rnd * wid)
        c = Int(Rnd * hgt)
        pixel = Int(Rnd * 3)

        ' See if the position is unused.
        position_code = "(" & r & "," & c & "," & pixel & ")"
        used_positions.Add position_code, position_code
        If Err.Number = 0 Then Exit Do
        Err.Clear
    Loop
End Sub

' Return the color's components.
' Helps gets back the Red,Green and Blue components of the pixel
Public Sub getRGBFromPixel(ByVal color As OLE_COLOR, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    'FF =1111 1111 (8 one bits)
    r = color And &HFF& 'First 8 bit mask which is Red
    g = (color And &HFF00&) \ &H100& 'Second 8 bit mask which Green
    b = (color And &HFF0000) \ &H10000 'Third 8 bit mask which is Blue
End Sub
