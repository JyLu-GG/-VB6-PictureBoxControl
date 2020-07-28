VERSION 5.00
Begin VB.Form frmPictureBoxControlPractice 
   Caption         =   "PicturebBox Control Practice"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17445
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   17445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cndSaveImage 
      Caption         =   "Save Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   11
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdSavePicture 
      Caption         =   "Save Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdDrawPicture 
      Caption         =   "Draw Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   9
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoadPicture 
      Caption         =   "Load Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetPictureBoxdotImage 
      Caption         =   "Get PictureBox.Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12480
      TabIndex        =   7
      Top             =   6240
      Width           =   4455
   End
   Begin VB.CommandButton cmdGetPictureBoxdotPicture 
      Caption         =   "Get PictureBox.Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   6
      Top             =   6240
      Width           =   4455
   End
   Begin VB.PictureBox pbImage 
      Height          =   4935
      Left            =   11760
      ScaleHeight     =   4875
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
   End
   Begin VB.PictureBox pbPicture 
      Height          =   4935
      Left            =   6000
      ScaleHeight     =   4875
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
   End
   Begin VB.PictureBox pbOrignal 
      Height          =   4935
      Left            =   240
      ScaleHeight     =   4875
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label lblPictureBoxdotImage 
      Caption         =   "PictureBox.Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   5
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label lblPictureBoxdotPicture 
      Caption         =   "PictureBox.Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label lblOrginalPictureBox 
      Caption         =   "Orignal PictureBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "frmPictureBoxControlPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' ============================================================
'' This is pictrue box control form.
'' Have some simply function for picturebox control practice.
''
'' Writer is J.Y.L by 2020/07/24
'' ============================================================

Option Explicit

Dim ret As Boolean

Private Sub cmdDrawPicture_Click()
    ret = DrawRectangle(pbOrignal, 2500, 2500, 4500, 4500, vbGreen)
    If ret = False Then MsgBox ("The DrawPicture is fail")
End Sub

Private Sub cmdGetPictureBoxdotImage_Click()
    ret = GetImageIntoPictureBox(pbImage, pbOrignal.Image)
    If ret = False Then MsgBox ("The GetImageIntoPictureBox is fail")
End Sub

Private Sub cmdGetPictureBoxdotPicture_Click()
    ret = GetPictureIntoPictureBox(pbPicture, pbOrignal.Picture)
    If ret = False Then MsgBox ("The GetPictureIntoPictureBox is fail")
End Sub

Private Sub cmdLoadPicture_Click()
    ret = LoadImage(pbOrignal, App.Path & "\" & "TestPicture.jpg")
    If ret = False Then MsgBox ("The LoadPicture is fail")
End Sub

Private Sub cmdSavePicture_Click()
    ret = SaveImage(pbOrignal.Picture, App.Path, "IsPicture.jpg")
    If ret = False Then MsgBox ("The SaveImage is fail")
End Sub

Private Sub cndSaveImage_Click()
    ret = SaveImage(pbOrignal.Image, App.Path, "IsImage.jpg")
    If ret = False Then MsgBox ("The SaveImage is fail")
End Sub

Private Sub Form_Load()
    '' down below will that draw image auto rewrok, is very important
    pbOrignal.AutoRedraw = True
    pbPicture.AutoRedraw = True
    pbImage.AutoRedraw = True
End Sub
