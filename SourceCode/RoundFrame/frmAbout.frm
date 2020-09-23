VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Rounded Frame Demo"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   330
      Left            =   4665
      TabIndex        =   4
      Top             =   3030
      Width           =   960
   End
   Begin VB.TextBox txtDisclaimer 
      Height          =   2055
      Left            =   1695
      TabIndex        =   3
      Top             =   825
      Width           =   3930
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   3570
      Left            =   0
      ScaleHeight     =   3570
      ScaleWidth      =   1545
      TabIndex        =   0
      Top             =   -15
      Width           =   1545
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Creation Date:"
         ForeColor       =   &H80000014&
         Height          =   240
         Left            =   75
         TabIndex        =   6
         Top             =   3060
         Width           =   1350
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "  July 25, 2000"
         ForeColor       =   &H80000014&
         Height          =   240
         Left            =   75
         TabIndex        =   5
         Top             =   3270
         Width           =   1350
      End
   End
   Begin VB.Label Label2 
      Caption         =   "by Joseph M. Ferris (joseph.ferris@cdicorp.com)"
      Height          =   225
      Left            =   1710
      TabIndex        =   2
      Top             =   450
      Width           =   3840
   End
   Begin VB.Label Label1 
      Caption         =   "Rounded Frame Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1695
      TabIndex        =   1
      Top             =   105
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()

Unload Me

End Sub

Private Sub Form_Load()

txtDisclaimer.Text = "The source code included is for demonstation purposes " & _
                     "only.  Neither the author, nor the respective party " & _
                     "that possesses this file for download, are to be held " & _
                     "responsible for any damages that may occur from the " & _
                     "execution of this source code." & vbCrLf & vbCrLf & _
                     "Please feel free to use this source code in your own " & _
                     "projects without obligation.  This source code is in " & _
                     "the Public Domain.  It would be appreciated if you " & _
                     "could contact the author to show this source code in " & _
                     "your own applications."
                     
End Sub
