VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "modMouseEvents Test - by Edgars Priede"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Mouse Move"
      Height          =   1410
      Left            =   585
      TabIndex        =   15
      Top             =   4215
      Width           =   3630
      Begin VB.CommandButton cmdGetPos 
         Caption         =   "Get Mouse Position"
         Height          =   300
         Left            =   705
         TabIndex        =   21
         Top             =   990
         Width           =   2190
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "Move Mouse"
         Height          =   300
         Left            =   705
         TabIndex        =   20
         Top             =   675
         Width           =   2190
      End
      Begin VB.TextBox txtY 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2310
         TabIndex        =   18
         Top             =   315
         Width           =   585
      End
      Begin VB.TextBox txtX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         TabIndex        =   16
         Top             =   315
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2100
         TabIndex        =   19
         Top             =   345
         Width           =   225
      End
      Begin VB.Label Label3 
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   690
         TabIndex        =   17
         Top             =   345
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mouse Buttons"
      Height          =   3195
      Left            =   585
      TabIndex        =   2
      Top             =   945
      Width           =   3630
      Begin VB.CommandButton Command1 
         Caption         =   "Left Click"
         Height          =   615
         Left            =   135
         TabIndex        =   14
         Top             =   300
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Right Click"
         Height          =   615
         Left            =   1275
         TabIndex        =   13
         Top             =   300
         Width           =   1065
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Middle Click"
         Height          =   615
         Left            =   2415
         TabIndex        =   12
         Top             =   300
         Width           =   1065
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Left Double Click"
         Height          =   615
         Left            =   135
         TabIndex        =   11
         Top             =   990
         Width           =   1065
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Right Double Click"
         Height          =   615
         Left            =   1275
         TabIndex        =   10
         Top             =   990
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Middle Double Click"
         Height          =   615
         Left            =   2415
         TabIndex        =   9
         Top             =   990
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Left Down"
         Height          =   615
         Left            =   135
         TabIndex        =   8
         Top             =   1680
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Right Down"
         Height          =   615
         Left            =   1275
         TabIndex        =   7
         Top             =   1680
         Width           =   1065
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Middle Down"
         Height          =   615
         Left            =   2415
         TabIndex        =   6
         Top             =   1680
         Width           =   1065
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Left Up"
         Height          =   615
         Left            =   135
         TabIndex        =   5
         Top             =   2370
         Width           =   1065
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Right Up"
         Height          =   615
         Left            =   1275
         TabIndex        =   4
         Top             =   2370
         Width           =   1065
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Middle Up"
         Height          =   615
         Left            =   2415
         TabIndex        =   3
         Top             =   2370
         Width           =   1065
      End
   End
   Begin VB.Shape Shape2 
      Height          =   810
      Left            =   510
      Top             =   60
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      Height          =   330
      Left            =   1170
      Top             =   5820
      Width           =   2340
   End
   Begin VB.Label Label2 
      Caption         =   "All code by Edgars Priede"
      Height          =   240
      Left            =   1425
      TabIndex        =   1
      Top             =   5880
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "To test modMouseEvents, drag mouse over something, and from keyboard click on these buttons using TAB, Navigation and SPACE keys."
      Height          =   675
      Left            =   615
      TabIndex        =   0
      Top             =   135
      Width           =   3555
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////
'//       Form: frmTest                                   //
'//       Author: Edgars Priede                           //
'//       E-Mail: edgars.software@inbox.lv                //
'//       Date: 15.07.2005                                //
'//       Description: modMouseEvent test application.    //
'///////////////////////////////////////////////////////////

Private Sub cmdGetPos_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position
    
    txtX.Text = X '\
'                   Show coordinates
    txtY.Text = Y '/
    
End Sub

Private Sub cmdMove_Click()
    
    Call MouseMove(txtX.Text, txtY.Text) 'Move mouse
    
End Sub

Private Sub Command1_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseLeftClick(X, Y) 'Call Mouse button event

End Sub

Private Sub Command10_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseMiddleDbClick(X, Y, 0.4) 'Call Mouse button event, interval 400 milliseconds

End Sub

Private Sub Command11_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseMiddleDown(X, Y) 'Call Mouse button event

End Sub

Private Sub Command12_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseMiddleUp(X, Y) 'Call Mouse button event

End Sub

Private Sub Command2_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseRightClick(X, Y) 'Call Mouse button event

End Sub

Private Sub Command3_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseLeftDbClick(X, Y, 0.4) 'Call Mouse button event, interval 400 milliseconds

End Sub

Private Sub Command4_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseRightDbClick(X, Y, 0.4) 'Call Mouse button event, interval 400 milliseconds

End Sub

Private Sub Command5_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseLeftDown(X, Y) 'Call Mouse button event

End Sub

Private Sub Command6_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseRightDown(X, Y) 'Call Mouse button event

End Sub

Private Sub Command7_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseLeftUp(X, Y) 'Call Mouse button event

End Sub

Private Sub Command8_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseRightUp(X, Y) 'Call Mouse button event

End Sub

Private Sub Command9_Click()

    Dim X As Long, Y As Long 'Dim variables

    Call GetMousePos(X, Y) 'Get mouse position

    Call MouseMiddleClick(X, Y) 'Call Mouse button event

End Sub
