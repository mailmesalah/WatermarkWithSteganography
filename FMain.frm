VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Watermark and Steganography"
   ClientHeight    =   7140
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu MWatermark 
      Caption         =   "Watermark"
   End
   Begin VB.Menu MSteganography 
      Caption         =   "Steganography"
      Begin VB.Menu MSEncode 
         Caption         =   "Encode"
      End
      Begin VB.Menu MSDecode 
         Caption         =   "Decode"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MSDecode_Click()
    FDecodeSteganography.Show
End Sub

Private Sub MSEncode_Click()
    FEncodeSteganography.Show
End Sub

Private Sub MWatermark_Click()
    FWatermark.Show
End Sub
