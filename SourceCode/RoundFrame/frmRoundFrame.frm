VERSION 5.00
Begin VB.Form frmRoundFrame 
   Caption         =   "Rounded Frame Demo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   300
      Left            =   3780
      TabIndex        =   9
      Top             =   2790
      Width           =   675
   End
   Begin VB.CommandButton cmdChangeCaption 
      Caption         =   "Change Caption"
      Height          =   300
      Left            =   1485
      TabIndex        =   8
      Top             =   2790
      Width           =   1320
   End
   Begin VB.PictureBox picRoundedFrame 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2670
      Left            =   60
      ScaleHeight     =   178
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   1
      Top             =   75
      Width           =   4515
      Begin VB.PictureBox picCaptionContainer 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   315
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   6
         Top             =   0
         Width           =   1200
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Test Frame"
            Height          =   195
            Left            =   45
            TabIndex        =   7
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.PictureBox picBotRight 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   4005
         ScaleHeight     =   360
         ScaleWidth      =   405
         TabIndex        =   5
         Top             =   2190
         Width           =   405
         Begin VB.Shape Shape10 
            BorderColor     =   &H80000010&
            Height          =   375
            Left            =   -210
            Top             =   -15
            Width           =   615
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H80000014&
            Height          =   375
            Left            =   -30
            Top             =   -30
            Width           =   420
         End
      End
      Begin VB.PictureBox picTopRight 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   4005
         ScaleHeight     =   360
         ScaleWidth      =   405
         TabIndex        =   4
         Top             =   0
         Width           =   405
         Begin VB.Shape Shape8 
            BorderColor     =   &H80000010&
            Height          =   300
            Left            =   -15
            Top             =   75
            Width           =   420
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H80000014&
            Height          =   300
            Left            =   -15
            Top             =   90
            Width           =   405
         End
      End
      Begin VB.PictureBox picBotLeft 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   405
         TabIndex        =   3
         Top             =   2190
         Width           =   405
         Begin VB.Shape Shape6 
            BorderColor     =   &H80000010&
            Height          =   375
            Left            =   60
            Top             =   -15
            Width           =   360
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H80000014&
            Height          =   360
            Left            =   75
            Top             =   -15
            Width           =   345
         End
      End
      Begin VB.PictureBox picTopLeft 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   405
         TabIndex        =   2
         Top             =   0
         Width           =   405
         Begin VB.Shape Shape4 
            BorderColor     =   &H80000014&
            Height          =   300
            Left            =   75
            Top             =   90
            Width           =   345
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H80000010&
            Height          =   300
            Left            =   60
            Top             =   75
            Width           =   360
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000014&
         Height          =   2445
         Left            =   75
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   4320
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000010&
         Height          =   2475
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   4350
      End
   End
   Begin VB.CommandButton cmdRandomCorners 
      Caption         =   "Round Corners"
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   2790
      Width           =   1320
   End
End
Attribute VB_Name = "frmRoundFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()

frmAbout.Show vbModal

End Sub

Private Sub Form_Load()

If lblCaption.Caption = "" Then                     ' If the user has set
    picCaptionContainer.Visible = False             ' the label to nothing,
    Exit Sub                                        ' hide the container.
End If

picCaptionContainer.Width = lblCaption.Width + 10   ' Set the label's container
                                                    ' to be ten pixels larger
                                                    ' than the Autosize label
                                                    ' inside of it.

End Sub

Private Sub cmdChangeCaption_Click()

Dim NewCaption As String                ' Create a string to hold our new caption

NewCaption = InputBox("What would you like to set the caption to?", "Caption Change")

lblCaption.Caption = NewCaption         ' Set new caption to label control

End Sub

Private Sub cmdRandomCorners_Click()

Dim bolCornerArray(3) As Boolean        ' Array to determine if corners are
                                        ' to be rounded:
                                        '
                                        ' 0 = False  : No Rounding
                                        ' 1 = True   : Rounded
                                        
Dim intLoopConstruct As Integer         ' Integer for the For...Next loop

For intLoopConstruct = 0 To 3
    
    bolCornerArray(intLoopConstruct) = Int(Rnd(8) * 2)  ' Assign value to a corner

Select Case intLoopConstruct            ' Determine which corner is active
                                        ' by being the intLoopConstruct variable

Case 0      ' Top Left Corner

    If bolCornerArray(intLoopConstruct) = 0 Then
        picTopLeft.Visible = False      ' Hide Rounded Corner
    Else
        picTopLeft.Visible = True       ' Show Rounded Corner
    End If

Case 1      ' Bottom Left Corner

    If bolCornerArray(intLoopConstruct) = 0 Then
        picBotLeft.Visible = False      ' Hide Rounded Corner
    Else
        picBotLeft.Visible = True       ' Show Rounded Corner
    End If

Case 2      ' Top Right Corner

    If bolCornerArray(intLoopConstruct) = 0 Then
        picTopRight.Visible = False     ' Hide Rounded Corner
    Else
        picTopRight.Visible = True      ' Show Rounded Corner
    End If

Case 3      ' Bottom Right Corner

    If bolCornerArray(intLoopConstruct) = 0 Then
        picBotRight.Visible = False     ' Hide Rounded Corner
    Else
        picBotRight.Visible = True      ' Show Rounded Corner
    End If

End Select

Next intLoopConstruct

End Sub

Private Sub lblCaption_Change()

If lblCaption.Caption = "" Then                     ' If the user has set
    picCaptionContainer.Visible = False             ' the label to nothing,
    Exit Sub                                        ' hide the container.
End If

picCaptionContainer.Width = lblCaption.Width + 10   ' Set the label's container
                                                    ' to be ten pixels larger
                                                    ' than the Autosize label
                                                    ' inside of it.

End Sub

