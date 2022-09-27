VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FDecodeSteganography 
   Caption         =   "Decode"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TPassword 
      Height          =   540
      Left            =   105
      TabIndex        =   5
      Top             =   1095
      Width           =   8325
   End
   Begin VB.CommandButton CDecode 
      Caption         =   "Decode"
      Height          =   555
      Left            =   8745
      TabIndex        =   4
      Top             =   1680
      Width           =   1770
   End
   Begin VB.TextBox TMessage 
      Height          =   540
      Left            =   105
      TabIndex        =   3
      Top             =   1710
      Width           =   8325
   End
   Begin VB.CommandButton CBrowseImage 
      Caption         =   "Browse Image"
      Height          =   555
      Left            =   8745
      TabIndex        =   2
      Top             =   450
      Width           =   1770
   End
   Begin VB.TextBox TImage 
      Height          =   540
      Left            =   105
      TabIndex        =   1
      Top             =   480
      Width           =   8325
   End
   Begin VB.PictureBox PImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4455
      Left            =   105
      ScaleHeight     =   4395
      ScaleWidth      =   10230
      TabIndex        =   0
      Top             =   2625
      Width           =   10290
   End
   Begin MSComDlg.CommonDialog DialogBox 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Left            =   8775
      TabIndex        =   6
      Top             =   1245
      Width           =   1695
   End
End
Attribute VB_Name = "FDecodeSteganography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBrowseImage_Click()
'This Sub routine is used to browse background Image
'file string variable is used to store the file name
Dim file As String
    'DialogBox opens the browser for browsing a file
    DialogBox.ShowOpen
    'Once browsing is over the filenames is obtained in the file variable
    file = DialogBox.FileName
    'The image text box is shown with the current image
    TImage.Text = file
    'The image file is loaded to the PImage PictureBox control
    PImage.Picture = LoadPicture(file)
End Sub

Private Sub CSelectOutputFolder_Click()
'The subroutine is used to select the custom output filename
'file is the variable used to store the output file name
Dim file As String
    'DialogBox control opens the browser for finding the path and file name on with the file need to be saved
    DialogBox.ShowSave
    'the filename is obtained in the variable file
    file = DialogBox.FileName
    'the obtained name is showed in the ToutputFolder text box
    TOutputFolder.Text = file
End Sub

Private Sub CDecode_Click()
Dim msg_length As Byte
Dim msg As String
Dim ch As Byte
Dim i As Integer
Dim used_positions As Collection
Dim wid As Integer
Dim hgt As Integer
Dim show_pixels As Boolean
 
    ' Initialize the random number generator.
    Rnd -1
    'Randomizes with the offset value of the password given
    Randomize NumericPassword(TPassword.Text)
    'Width and height is set from the Image
    wid = PImage.ScaleWidth
    hgt = PImage.ScaleHeight
    'Collection for storing used rows,columns and pixel combinations where the each bits of the character from message are stored, are saved
    Set used_positions = New Collection

    ' Decode the message length. First character of the encoded message is the length of the message which is retrieved
    ' Length of the message determines how many charatcers has to be decoded.
    msg_length = DecodeByte(used_positions, wid, hgt)

    ' Decode the message.
    For i = 1 To msg_length
        ' Each Characters are decoded
        ch = DecodeByte(used_positions, wid, hgt)
        ' Asci values are converted back to characters
        msg = msg & Chr$(ch)
    Next i
    ' The message Text box is updated with the obtained full message
    TMessage.Text = msg
    
    'Prints Out the Successful message
    MsgBox "Successfully Decoded !"

End Sub

' Decode this byte's data.
Private Function DecodeByte(ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer) As Byte
Dim Value As Integer
Dim i As Integer
Dim byte_mask As Integer
Dim r As Integer
Dim c As Integer
Dim pixel As Integer
Dim clrr As Byte
Dim clrg As Byte
Dim clrb As Byte
Dim color_mask As Integer

    'This variable decides which bit of the character is being masked at current iteration
    'For example :- 1= 00000001,2= 00000010,4=00000100 etc
    byte_mask = 1
    
    For i = 1 To 8
        ' Pick a random pixel and RGB component.
        PickPosition used_positions, wid, hgt, r, c, pixel

        ' Get the pixel's color components.
        ' From PImage.Point(r, c) pixel, its Red, Green and Blue components(values) are retrieved
        getRGBFromPixel PImage.Point(r, c), clrr, clrg, clrb

        ' Get the stored bit value by masking the first character of the selected pixel
        Select Case pixel
            Case 0
                'If the Red pixel combonent is selected, its value is Anded with hexadecimal value 1 to  obtained the bit value in the first postition from the right of the color combonent
                color_mask = (clrr And &H1)
            Case 1
                'If the Green pixel combonent is selected, its value is Anded with hexadecimal value 1 to  obtained the bit value in the first postition from the right of the color combonent
                color_mask = (clrg And &H1)
            Case 2
                'If the Blue pixel combonent is selected, its value is Anded with hexadecimal value 1 to  obtained the bit value in the first postition from the right of the color combonent
                color_mask = (clrb And &H1)
        End Select

        If color_mask Then
            ' The bit value is set to the character byte by Oring the current Character byte with the byte_mask byte which will be 1 in that particular bit position
            Value = Value Or byte_mask
        End If
        
        'bytemask updates to next bit position of the character byte selected from the message.
        'For example :- 1= 00000001,2= 00000010,4=00000100 etc
        byte_mask = byte_mask * 2
    Next i

    'Value variable is basically Integer variable, so has to be converted to Byte.
    DecodeByte = CByte(Value)
End Function

