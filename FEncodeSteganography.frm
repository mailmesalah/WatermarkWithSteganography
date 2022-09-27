VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FEncodeSteganography 
   Appearance      =   0  'Flat
   Caption         =   "Encode"
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
      TabIndex        =   7
      Top             =   2310
      Width           =   8325
   End
   Begin VB.CommandButton CSelectOutputFolder 
      Caption         =   "Select Output Folder"
      Height          =   555
      Left            =   8745
      TabIndex        =   6
      Top             =   1050
      Width           =   1770
   End
   Begin VB.TextBox TOutputFolder 
      Height          =   540
      Left            =   105
      TabIndex        =   5
      Top             =   1095
      Width           =   8325
   End
   Begin VB.CommandButton CEncode 
      Caption         =   "Encode"
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
      Top             =   2925
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
      TabIndex        =   8
      Top             =   2460
      Width           =   1695
   End
End
Attribute VB_Name = "FEncodeSteganography"
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

Private Sub CEncode_Click()
Dim msg As String
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
    '255 characters from the message text box is taken for encoding. Characters after 255 length is ignored
    msg = Left$(TMessage.Text, 255)
    'Collection for storing used rows,columns and pixel combinations where the each bits of the character from message are stored, are saved
    Set used_positions = New Collection

    ' Encode the message length.
    EncodeByte CByte(Len(msg)), used_positions, wid, hgt

    ' Encode the message.
    For i = 1 To Len(msg)
        'Each Characters from message is iterated and is encoded each bits of the character at 8 different row, column and pixel combinations
        EncodeByte Asc(Mid$(msg, i, 1)), used_positions, wid, hgt
    Next i
    'The processed Image is updated on the visible Image
    PImage.Picture = PImage.Image
    
    'Saves the processed Image to the output file name provided by the TOutputFolder text box
    SavePicture PImage.Image, Trim(TOutputFolder.Text)
    'Prints Out the Successful message
    MsgBox "Successfully Encoded !"

End Sub

' Encode this byte's data.
Private Sub EncodeByte(ByVal Value As Byte, ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer)
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
        
        ' Get the value we must store.
        ' The character asci value is Bitwise Anded by the byte_mask variable to retrive the current bit value to be stored in the pixel
        ' It will be either 0 or 1, for eg:- Value = 01101001 and byte_mask is 00000100 then we get 00000000 , ie 0(zero), Thus 3rd bit from right of the character is obtained
        If Value And byte_mask Then
            color_mask = 1
        Else
            color_mask = 0
        End If

        ' Update the color.
        ' Adds the obtained bit value to the last bit of the selected pixel byte.
        Select Case pixel
            Case 0
                ' If the pixel color selected is Red, then Red component of the pixel is changed
                clrr = (clrr And &HFE) Or color_mask
            Case 1
                ' If the selected pixel color is Green, its last bit is changed with the obtained value
                clrg = (clrg And &HFE) Or color_mask
            Case 2
                ' Or if the Blue color of the pixel is selected by PickPostiton() procedure the character bit obtained is stored in the first bit from the right of the Blue byte.
                clrb = (clrb And &HFE) Or color_mask
        End Select

        ' Set the pixel's color.
        ' The processed pixel is stored back in the Image.
        PImage.PSet (r, c), RGB(clrr, clrg, clrb)
        
        'bytemask updates to next bit position of the character byte selected from the message.
        'For example :- 1= 00000001,2= 00000010,4=00000100 etc
        byte_mask = byte_mask * 2
    Next i
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

Private Sub Form_Load()
'On loading the Form , the TOutputFolder control is set as current application path with OutputImage.jpg as name of the file
    TOutputFolder.Text = App.Path & "\OutputImage.jpg"
End Sub
