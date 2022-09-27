VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FWatermark 
   Caption         =   "Watermark"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   825
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox TextY 
      Height          =   420
      Left            =   3075
      TabIndex        =   16
      Text            =   "0"
      Top             =   3990
      Width           =   750
   End
   Begin VB.TextBox TextX 
      Height          =   420
      Left            =   1470
      TabIndex        =   15
      Text            =   "0"
      Top             =   3990
      Width           =   750
   End
   Begin VB.CommandButton CCreateImage 
      Caption         =   "Create Image"
      Height          =   555
      Left            =   9060
      TabIndex        =   8
      Top             =   2580
      Width           =   1770
   End
   Begin VB.TextBox TOutputFolder 
      Height          =   540
      Left            =   420
      TabIndex        =   7
      Top             =   1920
      Width           =   8325
   End
   Begin VB.CommandButton CSelectOutputFolder 
      Caption         =   "Select Output Folder"
      Height          =   555
      Left            =   9060
      TabIndex        =   6
      Top             =   1875
      Width           =   1770
   End
   Begin VB.TextBox TWatermark 
      Height          =   540
      Left            =   420
      TabIndex        =   5
      Top             =   1200
      Width           =   8325
   End
   Begin VB.TextBox TImage 
      Height          =   540
      Left            =   420
      TabIndex        =   4
      Top             =   525
      Width           =   8325
   End
   Begin VB.CommandButton CBrowseWatermark 
      Caption         =   "Browse Watermark"
      Height          =   555
      Left            =   9060
      TabIndex        =   3
      Top             =   1170
      Width           =   1770
   End
   Begin VB.CommandButton CBrowseImage 
      Caption         =   "Browse Image"
      Height          =   555
      Left            =   9060
      TabIndex        =   2
      Top             =   495
      Width           =   1770
   End
   Begin VB.PictureBox PWatermark 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   885
      Left            =   9465
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   1
      Top             =   3420
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox PImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1950
      Left            =   6015
      ScaleHeight     =   1890
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   3600
      Width           =   3195
   End
   Begin MSComDlg.CommonDialog DialogBox 
      Left            =   630
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Y"
      Height          =   330
      Left            =   2595
      TabIndex        =   18
      Top             =   4065
      Width           =   420
   End
   Begin VB.Label TX 
      Caption         =   "X"
      Height          =   330
      Left            =   1005
      TabIndex        =   17
      Top             =   4065
      Width           =   420
   End
   Begin VB.Label LWatermarkHeight 
      Caption         =   "0"
      Height          =   225
      Left            =   4935
      TabIndex        =   14
      Top             =   3330
      Width           =   3210
   End
   Begin VB.Label LWatermarkWidth 
      Caption         =   "0"
      Height          =   225
      Left            =   4935
      TabIndex        =   13
      Top             =   3090
      Width           =   3210
   End
   Begin VB.Label LImageHeight 
      Caption         =   "0"
      Height          =   225
      Left            =   480
      TabIndex        =   12
      Top             =   3300
      Width           =   3210
   End
   Begin VB.Label LImageWidth 
      Caption         =   "0"
      Height          =   225
      Left            =   480
      TabIndex        =   11
      Top             =   3060
      Width           =   3210
   End
   Begin VB.Label Label2 
      Caption         =   "Watermark Details"
      Height          =   225
      Left            =   4935
      TabIndex        =   10
      Top             =   2730
      Width           =   3210
   End
   Begin VB.Label Label1 
      Caption         =   "Image Details"
      Height          =   225
      Left            =   480
      TabIndex        =   9
      Top             =   2715
      Width           =   3210
   End
End
Attribute VB_Name = "FWatermark"
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
    'The width and Height of the Image is displayed in the following labels.
    LImageWidth.Caption = "Width " & PImage.ScaleWidth
    LImageHeight.Caption = "Height " & PImage.ScaleHeight
End Sub

Private Sub CBrowseWatermark_Click()
'This Sub routine is used to browse Watermark Image
'file string variable is used to store the file name
Dim file As String
    'DialogBox opens the browser for browsing a file
    DialogBox.ShowOpen
    'Once browsing is over the filenames is obtained in the file variable
    file = DialogBox.FileName
    'The watermark text box is shown with the current image
    TWatermark.Text = file
    'The image file is loaded to the PWarermark PictureBox control
    PWatermark.Picture = LoadPicture(file)
    'The width and Height of the Image is displayed in the following labels.
    LWatermarkWidth.Caption = "Width " & PWatermark.ScaleWidth
    LWatermarkHeight.Caption = "Height " & PWatermark.ScaleHeight
End Sub

Private Sub CCreateImage_Click()
'This subroutine creates the water mark and then saves to the output file
    'Creates the Watermark on the Background Image at given X and Y coordinates
    DrawWatermark PWatermark, PImage, Val(TextX.Text), Val(TextY.Text)
    'Saves the processed Image to the output file name provided by the TOutputFolder text box
    SavePicture PImage.Image, Trim(TOutputFolder.Text)
    'Prints Out the Successful message
    MsgBox "Successfully Done !"
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

Private Sub DrawWatermark(ByRef WaterMarkImage As PictureBox, ByRef BackGroundImage As PictureBox, ByVal x As Integer, ByVal y As Integer)
'Draws the watermark on the background image according the colors transperancy at the points where the water mark is draw pixels of watermark image
Const ALPHA As Byte = 128
Dim transparent As OLE_COLOR
Dim wmColor As OLE_COLOR
Dim bgColor As OLE_COLOR
Dim newColor As OLE_COLOR
Dim px As Integer
Dim py As Integer

    ' Get the transparent color.
    transparent = WaterMarkImage.Point(0, 0)
    
    WaterMarkImage.ScaleMode = vbPixels
    BackGroundImage.ScaleMode = vbPixels
    'Iterates through the Watermark Image by each pixels
    For py = 0 To WaterMarkImage.ScaleHeight - 1
        For px = 0 To WaterMarkImage.ScaleWidth - 1
            'Takes each pixels from water mark
            wmColor = WaterMarkImage.Point(px, py)
            'Checks the water mark pixel obtained is transparent or not
            If wmColor <> transparent Then
                'If the pixel is not transparent pixel of Background Image at that particular point is taken
                bgColor = BackGroundImage.Point(x + px, y + py)
                'New pixel is obtained by combining Background and Watermark pixels
                newColor = CombineColors(wmColor, bgColor, ALPHA)
                'Changes the pixel of Background with thr new pixel
                BackGroundImage.PSet (x + px, y + py), newColor
            End If
        Next px
    Next py
    
End Sub

Private Function CombineColors(ByVal clr1 As OLE_COLOR, ByVal clr2 As OLE_COLOR, ByVal A As Byte) As OLE_COLOR
'This function returns a color that is the combination of watermark and baground pixel consodering the Alpha for transperancy
Dim r1 As Long
Dim g1 As Long
Dim b1 As Long
Dim r2 As Long
Dim g2 As Long
Dim b2 As Long

    b1 = Int(clr1 / 65536)
    g1 = Int((clr1 - (65536 * b1)) / 256)
    r1 = clr1 - ((b1 * 65536) + (g1 * 256))

    b2 = Int(clr2 / 65536)
    g2 = Int((clr2 - (65536 * b2)) / 256)
    r2 = clr2 - ((b2 * 65536) + (g2 * 256))

    r1 = (A * r1 + (255 - A) * r2) \ 256
    g1 = (A * g1 + (255 - A) * g2) \ 256
    b1 = (A * b1 + (255 - A) * b2) \ 256

    CombineColors = r1 + 256 * g1 + 65536 * b1
End Function

