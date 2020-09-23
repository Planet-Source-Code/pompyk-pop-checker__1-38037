VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmainpop 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "POP CHECKER"
   ClientHeight    =   5970
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6360
      TabIndex        =   20
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "DELE 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6360
      TabIndex        =   19
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "DELE 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "RETR 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "RETR 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "LIST"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   15
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "STAT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "PASS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmduser 
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdsend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEND"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtsendcommand 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONNECT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtscreen 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IP ADDRESS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   3375
      Begin VB.TextBox txtipaddress 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "AUTHENTIFICATION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   3375
      Begin VB.TextBox txtusername 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtpassword 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   645
         Width           =   1215
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5040
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   7200
      X2              =   7200
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "POP COMMANDS...plz enter command one at a time...in order.."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   5400
      TabIndex        =   22
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   7
      X1              =   360
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "POP CHECKER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      Top             =   120
      Width           =   7455
   End
   Begin VB.Menu mnufile 
      Caption         =   "FILE"
      Begin VB.Menu mnuexit 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "ABOUT"
      Begin VB.Menu mnuaboutauthor 
         Caption         =   "AUTHOR"
      End
   End
End
Attribute VB_Name = "frmmainpop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
'Author: Somdutt Ganguly
'Email: gangulysomdutt@yahoo.com
'Address: No 6, chandrodaya apt, bhaikaka nagar
'Thaltej, Ahmedabad, Gujarat, INDIA - 380059
'Date: August 2002
'About: I am right now in my first Semister of MCA
'(master of computer application) from north gujarat university
'This software is to check the command of pop server
'Pop server recognizes some commands which u can test
'using this software...
'S.W name: pop checker...
'+++++++++++++++++++++++++++++++++++++++++++++++++++++


Dim username As String
Dim password As String

Private Sub cmdclear_Click()
txtscreen.Text = ""

End Sub

Private Sub cmdok_Click()
On Error GoTo x
Winsock1.Close
Winsock1.RemoteHost = txtipaddress
Winsock1.RemotePort = 110
Winsock1.Close
Winsock1.Connect
DoEvents
txtscreen.Text = txtscreen.Text & "===Connecting Connecting===" & vbCrLf
txtscreen.SelStart = Len(txtscreen.Text)
username = txtusername.Text
password = txtpassword.Text
Exit Sub
x:
MsgBox "Error: " & Err.Description
End Sub

Private Sub cmdsend_Click()
On Error GoTo x
Winsock1.SendData txtsendcommand.Text & returns
txtscreen.Text = txtscreen.Text & "===" & txtsendcommand.Text & "===" & vbCrLf
txtscreen.SelStart = Len(txtscreen.Text)
txtsendcommand.Text = ""
Exit Sub
x:
MsgBox "Error: " & Err.Description
txtsendcommand.Text = ""
txtsendcommand.SetFocus
End Sub

Private Sub cmduser_Click(Index As Integer)
On Error GoTo x
If cmduser(Index).Caption = "USER" Then
Winsock1.SendData "USER " + username + returns
txtscreen.Text = txtscreen.Text & "===USERNAME===" & username & vbCrLf
txtscreen.SelStart = Len(txtscreen.Text)

ElseIf cmduser(Index).Caption = "PASS" Then
Winsock1.SendData "PASS " + password + returns
txtscreen.Text = txtscreen.Text & "===PASSWORD===" & password & vbCrLf
txtscreen.SelStart = Len(txtscreen.Text)

Else

Winsock1.SendData cmduser(Index).Caption & returns
txtscreen.Text = txtscreen.Text & "===" & cmduser(Index).Caption & "===" & vbCrLf
txtscreen.SelStart = Len(txtscreen.Text)
End If
Exit Sub
x:
MsgBox "Error: " & Err.Description
End Sub

Private Sub Form_Load()
returns = Chr$(10) + Chr$(13)
End Sub

Private Sub mnuaboutauthor_Click()
MsgBox "Author: Somdutt Ganguly   E-mail: gangulysomdutt@yahoo.com Product: pop checker Date: August 2002 Dedication: Dedicated to Rohan Kaul and Pramod Sinha"
End Sub

Private Sub mnuexit_Click()
If Winsock1.State <> sckClosed Then
    Winsock1.Close
End If

MsgBox "Thx. for using .... pop checker..."
Unload Me
End Sub


Private Sub txtscreen_Change()
 If Len(txtscreen.Text) >= 31000 Then
    txtscreen.Text = Mid$(txtscreen.Text, Len(txtscreen) - 31000, 31000)
  End If
End Sub

Private Sub txtsendcommand_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdsend_Click
End If
End Sub


Private Sub Winsock1_Close()
txtscreen.Text = txtscreen.Text & "===Connection Closed===" & vbCrLf
txtscreen.SelStart = Len(txtscreen.Text)

End Sub

Private Sub Winsock1_Connect()
txtscreen.Text = txtscreen.Text & "===Connection Established===" & vbCrLf
txtscreen.SelStart = Len(txtscreen.Text)

End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim somduttkadata As String
  Winsock1.GetData somduttkadata
  txtscreen.Text = txtscreen.Text & "===Server Tells: " & somduttkadata & vbCrLf
  txtscreen.SelStart = Len(txtscreen.Text)
End Sub
