VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find the I.P. Address"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Host name:"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   4335
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "IP Addresses:"
      Height          =   2415
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter website below EX: www.Planet-source-code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    Text1.Text = vbNullString
    List1.Clear
    
End Sub

Private Sub cmdFind_Click()
    
    Dim lngPtrToHOSTENT As Long
   
    Dim udtHostent      As HOSTENT
   
    Dim lngPtrToIP      As Long
   
    Dim arrIpAddress()  As Byte
      
    Dim strIpAddress    As String
    '
    '----------------------------------------------------
   
    List1.Clear
   
    lngPtrToHOSTENT = gethostbyname(Trim$(Text1.Text))
    
    If lngPtrToHOSTENT = 0 Then
       
        ShowErrorMsg (Err.LastDllError)
        
    Else
        
        
        RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
        
        RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
        
        Do Until lngPtrToIP = 0
            
            ReDim arrIpAddress(1 To udtHostent.hLength)
            
            RtlMoveMemory arrIpAddress(1), lngPtrToIP, udtHostent.hLength
           
            For i = 1 To udtHostent.hLength
                strIpAddress = strIpAddress & arrIpAddress(i) & "."
            Next
           
            strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
           
            List1.AddItem strIpAddress
          
            strIpAddress = ""
           
            udtHostent.hAddrList = udtHostent.hAddrList + LenB(udtHostent.hAddrList)
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
            '
         Loop
        '
    End If
    '
End Sub



Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    '
    Dim lngRetVal      As Long
    Dim strErrorMsg    As String
    Dim udtWinsockData As WSAData
    Dim lngType        As Long
    Dim lngProtocol    As Long
    '

    lngRetVal = WSAStartup(&H101, udtWinsockData)
    '
    If lngRetVal <> 0 Then
        '
        '
        Select Case lngRetVal
            Case WSASYSNOTREADY
                strErrorMsg = "The underlying network subsystem is not " & _
                    "ready for network communication."
            Case WSAVERNOTSUPPORTED
                strErrorMsg = "The version of Windows Sockets API support " & _
                    "requested is not provided by this particular " & _
                    "Windows Sockets implementation."
            Case WSAEINVAL
                strErrorMsg = "The Windows Sockets version specified by the " & _
                    "application is not supported by this DLL."
        End Select
        '
        MsgBox strErrorMsg, vbCritical
        '
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WSACleanup
End Sub

Private Sub ShowErrorMsg(lngError As Long)
    '
    Dim strMessage As String
    '
    Select Case lngError
        Case WSANOTINITIALISED
            strMessage = "A successful WSAStartup call must occur " & _
                         "before using this function."
        Case WSAENETDOWN
            strMessage = "The network subsystem has failed."
        Case WSAHOST_NOT_FOUND
            strMessage = "Authoritative answer host not found."
        Case WSATRY_AGAIN
            strMessage = "Nonauthoritative host not found, or server failure."
        Case WSANO_RECOVERY
            strMessage = "A nonrecoverable error occurred."
        Case WSANO_DATA
            strMessage = "Valid name, no data record of requested type."
        Case WSAEINPROGRESS
            strMessage = "A blocking Windows Sockets 1.1 call is in " & _
                         "progress, or the service provider is still " & _
                         "processing a callback function."
        Case WSAEFAULT
            strMessage = "The name parameter is not a valid part of " & _
                         "the user address space."
        Case WSAEINTR
            strMessage = "A blocking Windows Socket 1.1 call was " & _
                         "canceled through WSACancelBlockingCall."
    End Select
    '
    MsgBox strMessage, vbExclamation
    '
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

