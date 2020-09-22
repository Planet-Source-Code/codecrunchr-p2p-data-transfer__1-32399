VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPeerB 
   Caption         =   "Peer B"
   ClientHeight    =   3240
   ClientLeft      =   4824
   ClientTop       =   456
   ClientWidth     =   3672
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   3672
   Begin VB.TextBox txtDepartment 
      Height          =   288
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3372
   End
   Begin VB.TextBox txtLastName 
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3372
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Width           =   972
   End
   Begin VB.TextBox txtFirstName 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3372
   End
   Begin MSWinsockLib.Winsock udpPeerB 
      Left            =   840
      Top             =   2280
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label lblMessage 
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   3372
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      X1              =   120
      X2              =   3480
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      Caption         =   "Department"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2292
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2292
   End
End
Attribute VB_Name = "frmPeerB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oPeer As New clsPeerB

Private Sub cmdSend_Click()
    ' define RemoteHost and send data
    With udpPeerB
        .RemoteHost = "127.0.0.1"
        .SendData oPeer.StoreData
    End With
End Sub

Private Sub udpPeerB_DataArrival(ByVal bytesTotal As Long)
    ' show returned data
    Dim strData As String
    udpPeerB.GetData strData
    lblMessage.Caption = strData
End Sub

Private Sub Form_Load()
    ' define remote port and bind
    With udpPeerB
        .RemotePort = 1636
        .Bind 1363
    End With
End Sub
