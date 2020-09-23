VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inet Sockets Demo"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSockets 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "100"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Sockets Baby!"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2800
      Width           =   1695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Index           =   0
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Sockets:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'A lot of people don't know about the cheap way of doing inet sockets
'but they are easy and really cool way of avoiding using winsock for
'a little while longer. I hope you get something out of this..

'VOTE FOR ME!!!

Private Sub cmdGo_Click()
    Dim SOCK As Long
    Dim X As Integer
    
    'Sets the loop to how many sockets you pick..
    For X = 0 To txtSockets.Text
        List1.ListIndex = X
        SOCK = Inet1.Count
        Load Inet1(SOCK)
        Inet1(SOCK).Execute List1.Text
    Next
End Sub

Private Sub Form_Load()
    Dim intX As Integer
    
    'Just random ass sites for inet to execute..
    For intX = 0 To 25
        List1.AddItem "http://www.google.com"
        List1.AddItem "http://www.yahoo.com"
        List1.AddItem "http://www.xspork.com"
        List1.AddItem "http://www.dictionary.com"
        List1.AddItem "http://www.download.com"
        List1.AddItem "http://www.realm-x.net"
        List1.AddItem "http://www.microsoft.com"
        List1.AddItem "http://www.powerade.com"
        List1.AddItem "http://www.planetsourcecode.com"
        List1.AddItem "http://www.redhat.com"
        List1.AddItem "http://www.hotmail.com"
        List1.AddItem "http://www.goarmy.com"
        List1.AddItem "http://www.billybussey.com"
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'Unloads all the sockets when program is closed..
    Dim i As Integer
    For i = 0 To Inet1.UBound
        Unload Inet1(i)
    Next i
End Sub

Private Sub Inet1_StateChanged(Index As Integer, ByVal State As Integer)
    'Takes the sockets when they are finished executing and gives them
    'something else to execute... loops the sockets till listbox is out..
    If State = 12 Then
        If List1.ListIndex = List1.ListCount - 1 Then: Exit Sub
        List1.ListIndex = List1.ListIndex + 1
        Inet1(Index).Execute List1.Text
    End If
End Sub
