VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Exit"
      Height          =   2175
      Left            =   6960
      TabIndex        =   11
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Connect"
      Height          =   2175
      Left            =   6960
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Host"
      Height          =   2175
      Left            =   6960
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "?"
      Height          =   2175
      Left            =   4680
      TabIndex        =   8
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "?"
      Height          =   2175
      Left            =   2400
      TabIndex        =   7
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "?"
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "?"
      Height          =   2175
      Left            =   4680
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   2175
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   2175
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "?"
      Height          =   2175
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YourID As String
Dim HisID As String
Dim Win As Boolean
Dim MyGo As Boolean
Dim Cmd1 As String
Dim Cmd2 As String
Dim Cmd3 As String
Dim Cmd4 As String
Dim Cmd5 As String
Dim Cmd6 As String
Dim Cmd7 As String
Dim Cmd8 As String
Dim Cmd9 As String

Private Sub Command1_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "1"
Command1.Caption = YourID
Cmd1 = YourID
Command1.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command10_Click()
Winsock1.LocalPort = 2000
Winsock1.Listen
Command10.Enabled = False
Command11.Enabled = False
YourID = "X"
HisID = "0"
MyGo = True
End Sub

Private Sub Command11_Click()
Host = InputBox("Enter the host's computer name or ip address:")
Winsock1.Connect Host, 2000
Command10.Enabled = False
Command11.Enabled = False
YourID = "0"
HisID = "X"
MyGo = False
End Sub

Private Sub Command12_Click()
End
End Sub

Private Sub Command2_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "2"
Command2.Caption = YourID
Cmd2 = YourID
Command2.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command3_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "3"
Command3.Caption = YourID
Cmd3 = YourID
Command3.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command4_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "4"
Command4.Caption = YourID
Cmd4 = YourID
Command4.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command5_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "5"
Command5.Caption = YourID
Cmd5 = YourID
Command5.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command6_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "6"
Command6.Caption = YourID
Cmd6 = YourID
Command6.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command7_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "7"
Command7.Caption = YourID
Cmd7 = YourID
Command7.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command8_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "8"
Command8.Caption = YourID
Cmd8 = YourID
Command8.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Command9_Click()
If MyGo = False Then Exit Sub
Winsock1.SendData YourID & "9"
Command9.Caption = YourID
Cmd9 = YourID
Command9.Enabled = False
CheckWin
MyGo = False
End Sub

Private Sub Form_Load()
Cmd1 = 1
Cmd2 = 2
Cmd3 = 3
Cmd4 = 4
Cmd5 = 5
Cmd6 = 6
Cmd7 = 7
Cmd8 = 8
Cmd9 = 9
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close  'Got to do this to make sure the Winsock control isn't already being used.
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Data, vbString
If Left(Data, 1) = YourID Then MyGo = False Else MyGo = True
If Right(Data, 1) = 1 Then
Command1.Caption = HisID
Cmd1 = HisID
Command1.Enabled = False
End If
If Right(Data, 1) = 2 Then
Command2.Caption = HisID
Cmd2 = HisID
Command2.Enabled = False
End If
If Right(Data, 1) = 3 Then
Command3.Caption = HisID
Cmd3 = HisID
Command3.Enabled = False
End If
If Right(Data, 1) = 4 Then
Command4.Caption = HisID
Cmd4 = HisID
Command4.Enabled = False
End If
If Right(Data, 1) = 5 Then
Command5.Caption = HisID
Cmd5 = HisID
Command5.Enabled = False
End If
If Right(Data, 1) = 6 Then
Command6.Caption = HisID
Cmd6 = HisID
Command6.Enabled = False
End If
If Right(Data, 1) = 7 Then
Command7.Caption = HisID
Cmd7 = HisID
Command7.Enabled = False
End If
If Right(Data, 1) = 8 Then
Command8.Caption = HisID
Cmd8 = HisID
Command8.Enabled = False
End If
If Right(Data, 1) = 9 Then
Command9.Caption = HisID
Cmd9 = HisID
Command9.Enabled = False
End If
CheckWin
End Sub

Sub CheckWin()
If Command1.Enabled = False And Command2.Enabled = False And Command3.Enabled = False And Command4.Enabled = False And Command5.Enabled = False And Command6.Enabled = False And Command7.Enabled = False And Command8.Enabled = False And Command9.Enabled = False Then MsgBox "Game Over - Draw"
If (Cmd1 = Cmd2) And (Cmd2 = Cmd3) Then MsgBox Cmd1 & " is the winner"
If (Cmd1 = Cmd4) And (Cmd4 = Cmd7) Then MsgBox Cmd1 & " is the winner"
If (Cmd1 = Cmd5) And (Cmd5 = Cmd9) Then MsgBox Cmd1 & " is the winner"
If (Cmd2 = Cmd5) And (Cmd5 = Cmd8) Then MsgBox Cmd2 & " is the winner"
If (Cmd3 = Cmd5) And (Cmd5 = Cmd7) Then MsgBox Cmd3 & " is the winner"
If (Cmd3 = Cmd6) And (Cmd6 = Cmd9) Then MsgBox Cmd3 & " is the winner"
If (Cmd4 = Cmd5) And (Cmd5 = Cmd6) Then MsgBox Cmd4 & " is the winner"
If (Cmd7 = Cmd8) And (Cmd8 = Cmd9) Then MsgBox Cmd7 & " is the winner"

End Sub
