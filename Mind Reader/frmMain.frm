VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Mind Reader"
   ClientHeight    =   6180
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   10620
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   6180
   ScaleWidth      =   10620
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTab 
      BackColor       =   &H00FAF0E7&
      Height          =   3375
      Left            =   480
      ScaleHeight     =   3315
      ScaleWidth      =   9675
      TabIndex        =   1
      Top             =   2040
      Width           =   9735
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmMain.frx":1CA4
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label lbSym 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00783F0C&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbNum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "0  >  "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCapture 
      BackColor       =   &H00F3DDC5&
      Caption         =   "&Capture"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdAgain 
      BackColor       =   &H00F3DDC5&
      Caption         =   "Again"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lbMin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9840
      TabIndex        =   13
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   10080
      Picture         =   "frmMain.frx":1D3D
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here is your Result >>"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   10
      Top             =   3120
      Width           =   2745
   End
   Begin VB.Label lbShow 
      Alignment       =   2  'Center
      BackColor       =   &H00F4C091&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   72
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0088480F&
      Height          =   1590
      Left            =   5040
      TabIndex        =   9
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4.  Click 'Capture' button to Read  your mind and capture that symbol. >> "
      Height          =   240
      Left            =   600
      TabIndex        =   7
      Top             =   5520
      Width           =   6345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3.  Find the corresponding symbol ( ie  45 > .... )  from the Table bellow. >> "
      Height          =   240
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   6450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.  Reduse the value of the each digit  from  the Number.  >> ie  ( 54 - 5 - 4 ) = 45"
      Height          =   240
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   6840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.  Think  a  number with two Digits  >> Eg. 54 "
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   3930
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "frmMain.frx":1E73
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mind Reader"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************
'By     Jim Jose
'email  jimjosev33@yahoo.com
'****************************

'PLEASE READ THIS
'This project is made to show 'How we can load controls at RunTime'.

'The  Idea behind the 'Mind Reading' is
'for  every numbers ,the value(number)-value(Each Digits)
'is  always a multiple of  'Nine'
'The symbol for the multiples of nine(1,9,18,...,99) is  fixed
'to a  common one and that  will be always  the result

'If you feel Satisfactory
'   Please 'Rate' this code
'Else
'   Give feedback to improve this code.
'End If
'Good luck
'****************************
Option Explicit
Private Sub cmdAgain_Click()
'Resetting the symbols and rearranging the form
    Reload
    picTab.Visible = True
    cmdCapture.Visible = True
End Sub
Private Sub cmdCapture_Click()
'Shows the  result
'1.lbShow > the result is displayed in it
    picTab.Visible = False
    lbShow = lbSym(9)
    cmdCapture.Visible = False
End Sub
Private Sub Form_Load()
'****************************
'This project is made to show
''How we can load controls at RunTime'.
'****************************

'Declaring the loop variables
Dim X  As Integer
Dim Y As Integer
'IndexNum determines the index of the newly loaded control
Dim IndexNum   As Integer

'****************************
'We  have two Labels,
'1.lbNum > Displays the numbers ( 0 to 99 )
'2.lbSym > Displays the corresponding  symbols

'the range 0-8 and 1-11  is a little confusing.
'they are only to disorder the numbers
'***************************

'Staring the loading
For X = 0 To 8  'loops through the columns
    For Y = 1 To 11 'loops through the rows
        'determining the indexNum.They can also can use as the label Captions
        IndexNum = X * 11 + Y
        'Loading the control
        Load lbNum(IndexNum)
        'Moves to correct position 1050,255,735 are the lbNum.width ,lbNum.height,lbSym.Width respectively
        lbNum(IndexNum).Move X * 1050, Y * 255, 735, 255
        'Make it visible.Because the default loaded is in Invisible condition
        lbNum(IndexNum).Visible = True
        'Setting the Capton
        lbNum(IndexNum) = IndexNum & "  >  "
        
        'Loading the control and all similer
        Load lbSym(IndexNum)
        lbSym(IndexNum).Move X * 1050 + 735, Y * 255, 315, 255
        lbSym(IndexNum).Visible = True
    Next Y
Next X
'Load the symbols
Reload
End Sub
Public Sub Reload()
'Loop variables
Dim X  As Integer
Dim Y As Integer
'Trick' follows
Dim Trick As String
Dim IndexNum   As Integer

'**************************************
'The  Idea behind the 'Mind Reading' is
'for  every numbers ,the value(number)-value(Each Digits)
'is  always a multiple of  'Nine'
'The symbol for the multiples of nine(1,9,18,...,99) is  fixed
'to a  common one ( ie Trick ) and that  will be always  the result
'**************************************

'Setting the common value for all Multiples of nine
Trick = Chr(40 + Int((100 * Rnd)))

For X = 0 To 8 'looping through the columns
    For Y = 1 To 11 'looping through the rows
        IndexNum = X * 11 + Y
        If IndexNum / 9 = Int(IndexNum / 9) Then 'Catching the multiples of nine
            'Setting the common value
            'The Font of lbSym is alreadyselected as 'Wingdings' which is a symbol font
            lbSym(IndexNum) = Trick
        Else
            'Setting random values for all other labels
            lbSym(IndexNum) = Chr(40 + Int((100 * Rnd) + 1))
        End If
    Next Y
Next X
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Mind Reader") = vbYes Then
        MsgBox "Please Rate this as an entertainer.The site address will be copied to your clipboard", vbInformation, "ThankYou"
    Else
        MsgBox "Please give FeedBack,The site address will be copied to your clipboard", vbInformation, "Please Give FeedBack"
    End If
    Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57560&lngWId=1")
End Sub
Private Sub imgClose_Click()
    Unload Me
End Sub
Private Sub lbMin_Click()
    Me.WindowState = 1
End Sub
