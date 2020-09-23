VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shaped Controls Example"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   2055
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3720
      TabIndex        =   3
      Top             =   60
      Width           =   2895
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3615
      LargeChange     =   5
      Left            =   2280
      Max             =   90
      TabIndex        =   2
      Top             =   60
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      Height          =   1815
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1860
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   60
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   315
      Left            =   60
      Top             =   3720
      Width           =   6555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UpdateShapes()

    ' In this subroutine, we shape all the objects on the form
    
    ShapeObject Me, 1  ' We have to use twips for the form because we cant use its scaled dimensions
    ShapeObject Picture1
    ShapeObject Command1
    ShapeObject VScroll1
    ShapeObject List1
    ShapeObject Option1
    
End Sub
Private Sub Form_Load()
    DoEvents
    Me.Show      ' Show the form
    UpdateShapes ' Shape the Objects

    ' Display the greeting message

    Form1.CurrentX = (Me.ScaleWidth / 4)
    Form1.CurrentY = Shape1.Top + 3
    Form1.Print "Move the slider bar to see the interactive demonstration"
End Sub

Private Sub VScroll1_Change()

    ' Resize the controls on the fly. cSize is a global variable defined in module 1.
    
    cSize = VScroll1.Value  ' Change cSize
    UpdateShapes            ' Update the controls
End Sub

Private Sub VScroll1_Scroll()

    VScroll1_Change         ' Notifiy the scrollbar that it has changed
    
End Sub
