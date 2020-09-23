VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "WinSock API"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Auto Update"
      Height          =   375
      Left            =   1800
      TabIndex        =   34
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9480
      TabIndex        =   32
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   2535
      Left            =   120
      TabIndex        =   31
      Top             =   3600
      Width           =   10695
      Begin VB.ListBox List3 
         Height          =   1815
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   10455
      End
      Begin VB.Label Label13 
         Caption         =   "Status"
         Height          =   255
         Index           =   6
         Left            =   8880
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Remote Port"
         Height          =   255
         Index           =   5
         Left            =   7800
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Remote Computer"
         Height          =   255
         Index           =   4
         Left            =   6000
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Remote IP"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   39
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Local Port"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Local Computer"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Local IP"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Get Host By IP Address"
      Height          =   1935
      Left            =   7320
      TabIndex        =   23
      Top             =   1560
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "Get Host Name"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Text            =   "Type IP Address Here"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "Host Name"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Get Local Information"
      Height          =   1935
      Left            =   3720
      TabIndex        =   19
      Top             =   1560
      Width           =   3495
      Begin VB.ListBox List2 
         Height          =   1035
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Get"
         Height          =   285
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Get Host By Name"
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   3495
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get"
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "Type Host Name Here"
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Winsock API Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Vendor Info:"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Max UdpDg:"
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Max Sockets:"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "System Status:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Description:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "High Version:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Version:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Program by John Phillips - vbjack@nyc.rr.com"
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
      Left            =   7320
      TabIndex        =   26
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   25
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WinSock API"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   24
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example was made possible by a lot of research from the web
' some from variouse sources on the internet
' MSDN Library
' Microsoft Knowledge base
' Planet-Source-Code.com
' friends and collegues from work
' There a lot of of useful functions and examples of the
' winsock api functions
' Use them as you see fit, all I ask is if you rip the code
' from this example please give me credit for the functions I created
' I wouldnt mind positive and negative feedback (negative feedback makes the app better) not too much though
' I also wouldnt mind a few votes on PSC - thanks and I hope this benifits someone

Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type
'
Private Const ERROR_BUFFER_OVERFLOW = 111&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_NO_DATA = 232&
Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&
'
Private Const MIB_TCP_STATE_CLOSED = 1
Private Const MIB_TCP_STATE_LISTEN = 2
Private Const MIB_TCP_STATE_SYN_SENT = 3
Private Const MIB_TCP_STATE_SYN_RCVD = 4
Private Const MIB_TCP_STATE_ESTAB = 5
Private Const MIB_TCP_STATE_FIN_WAIT1 = 6
Private Const MIB_TCP_STATE_FIN_WAIT2 = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT = 8
Private Const MIB_TCP_STATE_CLOSING = 9
Private Const MIB_TCP_STATE_LAST_ACK = 10
Private Const MIB_TCP_STATE_TIME_WAIT = 11
Private Const MIB_TCP_STATE_DELETE_TCB = 12
'
Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPROW) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
'
Private aTcpTblRow() As MIB_TCPROW

Private mWSData As WSAData ' this will hold the wsadata we need

Private Sub Command1_Click()
If Text1.Text <> "Type Host Name Here" Then
Screen.MousePointer = vbHourglass
WSAPIFun1 2, Text1, List1
Screen.MousePointer = vbNormal
End If
End Sub

Private Sub Command2_Click()
Screen.MousePointer = vbHourglass
WSAPIFun1 1, Text2, List2
Screen.MousePointer = vbNormal
End Sub

Private Sub Command3_Click()
If Text3.Text = "Type IP Address Here" Then Exit Sub
Screen.MousePointer = vbHourglass
' The inet_addr function returns a long value
    Dim lInteAdd As Long
' pointer to the HOSTENT
    Dim lPointtoHost As Long
' host name we are looking for
    Dim sHost As String
' Hostent
    Dim mHost As HOSTENT
' IP Address
    Dim sIP As String

    sIP = Trim$(Text3.Text)

' Convert the IP address
    lInteAdd = inet_addr(sIP)

' if the wrong IP format was entered there is an err generated
    If lInteAdd = INADDR_NONE Then

        WSErrHandle (Err.LastDllError)

    Else

' pointer to the Host
        lPointtoHost = gethostbyaddr(lInteAdd, 4, PF_INET)

' if zero is returned then there was an error
        If lPointtoHost = 0 Then

            WSErrHandle (Err.LastDllError)

        Else

            RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

            sHost = String(256, 0)

' Copy the host name
            RtlMoveMemory ByVal sHost, ByVal mHost.hName, 256

' Cut the chr(0) character off
            sHost = Left(sHost, InStr(1, sHost, Chr(0)) - 1)

' Return the host name
            Text4.Text = sHost

        End If

    End If
Screen.MousePointer = vbNormal
End Sub

Private Sub Command4_Click()
Unload Me
End
End Sub

Private Sub Command5_Click()
UpdateList
End Sub

Private Sub Command6_Click()
MsgBox "I didnt add this into this project" & vbCrLf & "if you want to know what to do double click the cmd button in vb and read the comments", vbOKOnly + vbInformation, "Not Put In"
' just add a timer and set the enabled to false
' set the interval to how ever long you want inbetween updates
' 1000 = 1 sec
' then in this command buttons click event
' put some code like this
' if command6.caption = "Auto Update" then
' timer1.enabled = true
' command6.caption = "Stop"
' exit sub
' elseif command6.caption = "Stop" then
' timer1.enables = false
' command6.caption = "Auto Update"
' exit sub
' end if

' then in the timer1 event put code like this
'
' updatelist
'
' and thats it when the button is clicked the timer will run the
' updatelist function every second.
End Sub

Private Sub Form_Load()
' I would go into more detail here but most of this information can be found in the MSDN
' Library that came with VB when you bought it
' Otherwise the knowledge base on the microsoft web site has
' almost all of the information needed if not all of it
' for this version we are using winsock version 1.1
' if you want to use winsock version 2.2 then change
' lV = WSAStartup(&H101, mWSD) to
' lV = WSAStartup(&H202, mWSD)

Dim lV As Long
Dim mWSD As WSAData

' start the winsock service
' we need to load this before we can do any type of winsocking :)

    lV = WSAStartup(&H101, mWSD)

' this is to check and make sure the winsock service has started
' before we proceed any further

    If lV <> 0 Then

    Select Case lV
        Case WSASYSNOTREADY ' winsock error system not ready
            MsgBox "The system is not ready!", vbOKOnly + vbInformation, "Winsock Error"
        Case WSAVERNOTSUPPORTED ' winsock error API not supported
            MsgBox "The version of Windows Sockets API is not supported!", vbOKOnly + vbInformation, "Winsock Error"
        Case WSAEINVAL ' winsock error the socket version is not supported
            MsgBox "The Windows Sockets version is not supported!", vbOKOnly + vbInformation, "Winsock Error"
        Case Else
            MsgBox "An unknown error has occured!", vbOKOnly + vbInformation, "Winsock Error"
        End Select

    End If
    
mWSData = mWSD 'set our declaration to the wsadata

' set up our labels on our form with the winsock information
Label2.Caption = mWSData.wVersion \ 256 & "." & mWSData.wVersion Mod 256

Label3.Caption = mWSData.wHighVersion \ 256 & "." & mWSData.wHighVersion Mod 256
                  
Label4.Caption = mWSData.szDescription

Label5.Caption = mWSData.szSystemStatus

Label6.Caption = IntegerToUnsigned(mWSData.iMaxSockets)

Label7.Caption = IntegerToUnsigned(mWSData.iMaxUdpDg)

Label8.Caption = mWSData.lpVendorInfo


End Sub

Private Sub Form_Unload(Cancel As Integer)
' call the winsock cleanup to unload the winsock service
Call WSACleanup

' make sure the program ends and is unloaded from memory
End
End Sub

Private Function HostNameFromLong(lngInetAdr As Long) As String

    Dim lPointtoHost As Long
    
    Dim lPointtoHostName As Long
    
    Dim sHName As String
    
    Dim mHost As HOSTENT

' Get the pointer to the Host
    lPointtoHost = gethostbyaddr(lngInetAdr, 4, 1)

' put data into the Host
    RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

    sHName = String(256, 0)

' Copy the host name
    RtlMoveMemory ByVal sHName, ByVal mHost.hName, 256

    sHName = Left(sHName, InStr(1, sHName, Chr(0)) - 1)

    HostNameFromLong = sHName

End Function

Private Function UpdateList()

    Dim aBuf() As Byte
    Dim lSize As Long
    Dim lV As Long
    Dim lR As Long
    Dim i As Long
    Dim TCPtr As MIB_TCPROW

    List3.Clear

    Me.MousePointer = vbHourglass

    lSize = 0

' Call the GetTcpTable just to get the buffer size into the lSize variable
    lV = GetTcpTable(ByVal 0&, lSize, 0)

    If lV = ERROR_NOT_SUPPORTED Then
' API is not supported
        MsgBox "not supported by this system.", vbOKOnly + vbInformation, "Error"
        Exit Function
    End If
    
    ReDim aBuf(0 To lSize - 1) As Byte

    lV = GetTcpTable(aBuf(0), lSize, 0)

    If lV = ERROR_SUCCESS Then

        CopyMemory lR, aBuf(0), 4

        ReDim aTcpTblRow(1 To lR)

            Dim sListItem As String
            Dim lcIP As String
            Dim lcHst As String
            Dim lcPrt As String
            Dim rmIP As String
            Dim rmHst As String
            Dim rmPrt As String
            Dim tStat As String

        For i = 1 To lR

            DoEvents

' Copy the table row data to the TCPtr structure
            CopyMemory TCPtr, aBuf(4 + (i - 1) * Len(TCPtr)), Len(TCPtr)

' Add data to the listbox
           
                With TCPtr
                
                 
                 lcIP = GetIp(.dwLocalAddr)
                 lcHst = HostNameFromLong(.dwLocalAddr)
                 lcPrt = GetPort(.dwLocalPort)
                 rmIP = GetIp(.dwRemoteAddr)
                 rmHst = HostNameFromLong(.dwRemoteAddr)
                 rmPrt = GetPort(.dwRemotePort)
                 tStat = GetState(.dwState)
                
                ' just a check to see the type of data returned
                ' in the list box we need to add an extra tab space if
                ' localhost is returned - this may only be the case on my network
                ' because of the name sceme on the network - you may see the data displayed
                ' all messed up - in my case this make sthe data display nice.
                ' there are better ways to display this data but I just threw this
                ' together quick and didnt want to get into the listview control
                ' or a datagrid control - you should be able to get the idea and
                ' change this part as you see fit to display this data
                
                If lcHst = "localhost" Then
                lcHst = lcHst & vbTab
                End If
                
                If rmHst = "localhost" Then
                rmHst = rmHst & vbTab
                End If
                
                sListItem = sListItem & lcIP & vbTab & vbTab & lcHst & vbTab & vbTab & lcPrt & vbTab & rmIP & vbTab & vbTab & rmHst & vbTab & vbTab & rmPrt & vbTab & tStat
                    List3.AddItem sListItem
                End With
            sListItem = ""
            aTcpTblRow(i) = TCPtr

        Next i

    End If

    Me.MousePointer = vbNormal
End Function


Private Function GetState(lngState As Long) As String

    Select Case lngState
        Case MIB_TCP_STATE_CLOSED: GetState = "CLOSED"
        Case MIB_TCP_STATE_LISTEN: GetState = "LISTEN"
        Case MIB_TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
        Case MIB_TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
        Case MIB_TCP_STATE_ESTAB: GetState = "ESTAB"
        Case MIB_TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
        Case MIB_TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
        Case MIB_TCP_STATE_CLOSING: GetState = "CLOSING"
        Case MIB_TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
        Case MIB_TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
        Case MIB_TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
    End Select

End Function

Private Function WSAPIFun1(icType As Integer, tText As TextBox, tlist As ListBox)
' we are seeting up a function here to do some of the
' WS api calls - again we set up functions so there isnt much code being repeated
' ictype returns 1 = get name and IP address of local system
' ictype returns 2 = get remote host by name


' Pointer to host
    Dim lPointtoHost As Long
' stores all the host info
    Dim mHost As HOSTENT
' pointer to the IP address list - there may be several IP address for 1 host
    Dim lPointtoIP As Long
' array that holds elemets of an IP address
    Dim aIPAdd() As Byte
' IP address to add into the ListBox
    Dim sIPAdd As String

    tlist.Clear
' here we are checking to see what type of call we need
' if we want the host by name then we do not need the following code
' else if we want local ip address and name then we do
If icType = 1 Then
    Dim sHostN As String * 256
    Dim lV As Long

    lV = gethostname(sHostN, 256)

    If lV = SOCKET_ERROR Then
        WSErrHandle (Err.LastDllError)
        Exit Function
    End If

    tText.Text = Left(sHostN, InStr(1, sHostN, Chr(0)) - 1)
End If

' Call the gethostbyname Winsock API function
    lPointtoHost = gethostbyname(Trim$(tText.Text))

' Check to see if the lPointtoHost value has returned anything
' if we get a 0 then that means there was an error getting the host info
' here is where we saved time typeing and we call the error function
' we created for the winsock api
    If lPointtoHost = 0 Then
        WSErrHandle (Err.LastDllError)
    Else
' Copy data to mHost structure
        RtlMoveMemory mHost, lPointtoHost, LenB(mHost)

        RtlMoveMemory lPointtoIP, mHost.hAddrList, 4

        Do Until lPointtoIP = 0
            
            ReDim aIPAdd(1 To mHost.hLength)

            RtlMoveMemory aIPAdd(1), lPointtoIP, mHost.hLength

            For i = 1 To mHost.hLength
                sIPAdd = sIPAdd & aIPAdd(i) & "."
            Next

            sIPAdd = Left$(sIPAdd, Len(sIPAdd) - 1)

' Add the IP address to the listbox
            tlist.AddItem sIPAdd

            sIPAdd = ""

            mHost.hAddrList = mHost.hAddrList + LenB(mHost.hAddrList)
            RtlMoveMemory lPointtoIP, mHost.hAddrList, 4

         Loop
    End If

End Function

Private Function GetIp(lIPAdd As Long) As String

    GetIp = GetString(inet_ntoa(lIPAdd))

End Function
 

Private Function GetPort(lPort As Long) As Long

    GetPort = IntegerToUnsigned(ntohs(UnsignedToInteger(lPort)))

End Function
 


