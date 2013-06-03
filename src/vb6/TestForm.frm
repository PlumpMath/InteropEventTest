VERSION 5.00
Begin VB.Form TestForm 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Text            =   "no message received yet"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Text            =   "the message"
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "send the message"
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   3495
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents tester As InteropEventTest.InteropConnection

Private Sub Command1_Click()
    Call tester.DoEvent(Text1.Text)
End Sub

Private Sub Form_Load()
    Set tester = New InteropConnection
End Sub

Private Sub tester_OnHappened(ByVal theMessage As String)
    Text2.Text = theMessage
End Sub

Private Sub Text2_Change()

End Sub
