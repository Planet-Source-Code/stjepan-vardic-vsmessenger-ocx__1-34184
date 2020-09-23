VERSION 5.00
Object = "*\AVSMessengerControl.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VSButton Control - Test Project"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Retrieve the message"
      Height          =   330
      Left            =   75
      TabIndex        =   4
      Top             =   900
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set 3rd Message Active"
      Height          =   360
      Left            =   75
      TabIndex        =   3
      Top             =   495
      Width           =   1980
   End
   Begin VSMessengerControl.VSMessenger VSMessenger1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1695
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   556
      ForeColor       =   -2147483640
   End
   Begin VB.Label Label4 
      Caption         =   "You can also use the PopUp Menu of VSMessenger Control to navigate thru the message list "
      Height          =   450
      Left            =   2130
      TabIndex        =   6
      Top             =   2685
      Width           =   3825
   End
   Begin VB.Line Line3 
      X1              =   416
      X2              =   227
      Y1              =   171
      Y2              =   171
   End
   Begin VB.Line Line2 
      X1              =   416
      X2              =   416
      Y1              =   118
      Y2              =   172
   End
   Begin VB.Line Line1 
      X1              =   403
      X2              =   417
      Y1              =   118
      Y2              =   118
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2190
      TabIndex        =   5
      Top             =   945
      Width           =   3690
   End
   Begin VB.Label Label2 
      Caption         =   "Total number of messages is : "
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1170
      TabIndex        =   0
      Top             =   2040
      Width           =   3690
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
VSMessenger1.CurrentMessage 2 'Remember it starts from 0, so if you put 2 it will be 3rd message


End Sub

Private Sub Command2_Click()
Label3.Caption = VSMessenger1.Message

End Sub

Private Sub Form_Load()
VSMessenger1.AddMessage "PlanetSorceCode"
VSMessenger1.AddMessage "VSMessenger Control"
VSMessenger1.AddMessage "Copyright © 2002 by Stjepan Vardic"
VSMessenger1.AddMessage "All Rights Reserved."
VSMessenger1.AddMessage "Please, send comments and suggestions"
VSMessenger1.AddMessage "to"
VSMessenger1.AddMessage "topctc2001@yahoo.co.uk"
VSMessenger1.ForeColor = vbBlue
Label2.Caption = Label2.Caption & VSMessenger1.MessageCount

End Sub

Private Sub VSMessenger1_DownClick() 'We have clicked the Down Arrow
Label1.Caption = "Down arrow is clicked"

End Sub

Private Sub VSMessenger1_UpClick() 'We have clicked the Up Arrow
Label1.Caption = "UpArrow is clicked"

End Sub
