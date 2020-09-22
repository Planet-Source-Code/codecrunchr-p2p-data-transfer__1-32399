VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPeerA 
   Caption         =   "Peer A"
   ClientHeight    =   1944
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4584
   LinkTopic       =   "Form1"
   ScaleHeight     =   1944
   ScaleWidth      =   4584
   Begin VB.TextBox txtOutput 
      Height          =   1692
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   4092
   End
   Begin VB.TextBox txtSend 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   4092
   End
   Begin MSWinsockLib.Winsock udpPeerA 
      Left            =   4440
      Top             =   120
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmPeerA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oPeer As New clsPeerA

Private Sub Form_Load()
    ' define RemoteHost, RemotePort and bind
    With udpPeerA
        .RemoteHost = "127.0.0.1"
        .RemotePort = 1363
        .Bind 1636
    End With
    ' open peer b
    frmPeerB.Show
End Sub

Private Sub udpPeerA_DataArrival(ByVal bytesTotal As Long)
    ' collect and show data and return message to sender
    Dim strData As String
    udpPeerA.GetData strData
    txtOutput.Text = ""
    oPeer.ReturnData strData
    udpPeerA.SendData "Data Transfer Successful..."
End Sub
