VERSION 5.00
Begin VB.Form frmTransBlt 
   Caption         =   "Transparent Blitting"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Form"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrawSprite 
      Caption         =   "Draw Sprite"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrawMask 
      Caption         =   "Draw Masks"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   1680
      Picture         =   "frmTransBlt.frx":0000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   3120
      Width           =   1980
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   3840
      Picture         =   "frmTransBlt.frx":C042
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   3120
      Width           =   1980
   End
End
Attribute VB_Name = "frmTransBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Transparent Blts
'
Option Explicit


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long



Private Sub cmdClear_Click()

'Clear the form
Me.Cls

End Sub

Private Sub cmdDrawMask_Click()

'Draws the mask with vbSrcAnd raster operation
BitBlt Me.hDC, 0, 0, picMask.ScaleWidth, picMask.ScaleHeight, picMask.hDC, 0, 0, vbSrcAnd

End Sub

Private Sub cmdDrawSprite_Click()

'Draws the sprite witht the vbSrcPaint raster operation
BitBlt Me.hDC, 0, 0, picSprite.ScaleWidth, picSprite.ScaleHeight, picSprite.hDC, 0, 0, vbSrcPaint

End Sub
