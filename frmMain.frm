VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBMarketWatch"
   ClientHeight    =   4260
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   270
      Top             =   6570
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   44
      Top             =   6660
      Width           =   105
   End
   Begin VB.Frame Frame1 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      Begin VB.PictureBox picTitleBar 
         Height          =   645
         Left            =   90
         ScaleHeight     =   585
         ScaleWidth      =   7875
         TabIndex        =   35
         Top             =   135
         Width           =   7935
         Begin VB.CommandButton cmdCommand1 
            Height          =   195
            Left            =   4140
            TabIndex        =   45
            Top             =   675
            Width           =   375
         End
         Begin VB.CommandButton cmdGetQuote 
            Caption         =   "&Get Quote"
            Height          =   310
            Left            =   1770
            TabIndex        =   41
            Top             =   180
            Width           =   1032
         End
         Begin VB.CommandButton cmdViewGraph 
            Caption         =   "&View Graph"
            Height          =   310
            Left            =   6390
            TabIndex        =   40
            Top             =   135
            Width           =   1455
         End
         Begin VB.ComboBox cboGraph 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            ItemData        =   "frmMain.frx":030A
            Left            =   4770
            List            =   "frmMain.frx":031A
            Style           =   2  'Dropdown List
            TabIndex        =   39
            ToolTipText     =   "Select the type of graph you would like"
            Top             =   180
            Width           =   1512
         End
         Begin VB.ComboBox cboSymbol 
            Height          =   315
            ItemData        =   "frmMain.frx":0354
            Left            =   630
            List            =   "frmMain.frx":0356
            TabIndex        =   38
            Text            =   "Symbol"
            ToolTipText     =   "Enter you stock symbol here"
            Top             =   180
            Width           =   1092
         End
         Begin VB.CommandButton cmdAddSymbol 
            Caption         =   "+"
            Height          =   312
            Left            =   2850
            TabIndex        =   37
            ToolTipText     =   "Add symbol to list"
            Top             =   180
            Width           =   372
         End
         Begin VB.CommandButton cmdRemoveSymbol 
            Caption         =   "-"
            Height          =   312
            Left            =   3270
            TabIndex        =   36
            ToolTipText     =   "Remove symbol from list"
            Top             =   180
            Width           =   372
         End
         Begin VB.Label Label20 
            Caption         =   "Symbol:"
            Height          =   255
            Left            =   30
            TabIndex        =   43
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Graph Type:"
            Height          =   255
            Left            =   3750
            TabIndex        =   42
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.Label lblConnection 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You are currently connected to the internet "
         Height          =   255
         Left            =   90
         TabIndex        =   48
         Top             =   3870
         Width           =   5055
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5220
         TabIndex        =   47
         Top             =   3870
         Width           =   1365
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6660
         TabIndex        =   46
         Top             =   3870
         Width           =   1365
      End
      Begin VB.Label lblDateTime 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   4905
         TabIndex        =   1
         Top             =   945
         Width           =   3075
      End
      Begin VB.Label lblLastGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   25
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label lblSEGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   2
         Top             =   3540
         Width           =   1245
      End
      Begin VB.Label lblPERGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   18
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label lblOpenGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   24
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblHighGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   23
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label lblLowGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblVolumeGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   21
         Top             =   2595
         Width           =   1215
      End
      Begin VB.Label lblPSEGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   20
         Top             =   2910
         Width           =   1215
      End
      Begin VB.Label lblPSOGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   19
         Top             =   3225
         Width           =   1215
      End
      Begin VB.Label lblChangeGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6780
         TabIndex        =   9
         Top             =   1335
         Width           =   1245
      End
      Begin VB.Label lblChangepGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   8
         Top             =   1650
         Width           =   1245
      End
      Begin VB.Label lbl52HighGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6810
         TabIndex        =   7
         Top             =   1965
         Width           =   1245
      End
      Begin VB.Label lbl52LowGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   6
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label lblBidGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   5
         Top             =   2595
         Width           =   1245
      End
      Begin VB.Label lblAskGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6825
         TabIndex        =   4
         Top             =   2925
         Width           =   1245
      End
      Begin VB.Label lblMCGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   3
         Top             =   3225
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Stock Exchange:"
         Height          =   255
         Index           =   15
         Left            =   4050
         TabIndex        =   10
         Top             =   3540
         Width           =   2775
      End
      Begin VB.Label lblSymbol 
         BackColor       =   &H80000005&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   90
         TabIndex        =   34
         Top             =   855
         Width           =   7935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Last:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   33
         Top             =   1335
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Open:"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   32
         Top             =   1650
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "High:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   31
         Top             =   1965
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Low:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   30
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Volume:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   29
         Top             =   2595
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Per Share Earning:"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   28
         Top             =   2910
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Per Share Outstanding:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   27
         Top             =   3225
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "P/E Ratio:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   26
         Top             =   3540
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Change:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   4050
         TabIndex        =   17
         Top             =   1335
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Change(%):"
         Height          =   255
         Index           =   9
         Left            =   4050
         TabIndex        =   16
         Top             =   1650
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "52 Week High:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   10
         Left            =   4050
         TabIndex        =   15
         Top             =   1965
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "52 Week Low:"
         Height          =   255
         Index           =   11
         Left            =   4050
         TabIndex        =   14
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Bid:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   12
         Left            =   4035
         TabIndex        =   13
         Top             =   2595
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Ask:"
         Height          =   255
         Index           =   13
         Left            =   4050
         TabIndex        =   12
         Top             =   2910
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Market Capitilization:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   14
         Left            =   4035
         TabIndex        =   11
         Top             =   3225
         Width           =   2775
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpSelect 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Begin VB.Menu mnuThisProgram 
            Caption         =   "This Program"
         End
         Begin VB.Menu mnuYourCompany 
            Caption         =   "About (Your Company)"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AppName As String 'Used for registry
Dim Section As String 'Used for registry
Dim sCustKey As String * 50 'Used for registry
Dim sCustVal As String * 50 'Used for registry
Dim Counter1 As Integer 'Used to enumerate the registry entries
Dim Counter As Integer  'Used to enumerate the registry entries
Dim j As String 'For Delete for loop
Dim j1 As Integer 'For Delete for loop
Dim Mysettings1 As Variant, intSettings1 As Integer 'For custlist
Dim intSet1 As Integer 'For custlist
Dim symbols As String 'For GetQuote


Private Sub Form_Load()
Dim eR As EIGCInternetConnectionState
Dim sName As String
Dim bConnected As Boolean

    With frmMain
    .Left = (Screen.Width - Me.Width) / 2
    .Top = (Screen.Height - Me.Height) / 2
    End With

   bConnected = InternetConnected(eR, sName)
   
   If (eR And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM Then
     lblConnection.Caption = lblConnection.Caption & "via modem." & vbCrLf
   End If
   
   If (eR And INTERNET_CONNECTION_LAN) = INTERNET_CONNECTION_LAN Then
     lblConnection.Caption = lblConnection.Caption & "via LAN." & vbCrLf
   End If
   
   If (eR And INTERNET_CONNECTION_PROXY) = INTERNET_CONNECTION_PROXY Then
     lblConnection.Caption = lblConnection.Caption & "via Proxy." & vbCrLf
   End If
   
   If (eR And INTERNET_CONNECTION_OFFLINE) = INTERNET_CONNECTION_OFFLINE Then
     lblConnection.Caption = "You are currently not connected to the internet." & vbCrLf
   End If
   
        lblSymbol.Caption = ""
        lblTime.Caption = Time
        lblDate.Caption = Date

        'Registry Load Stuff
        SaveSetting AppName:="Header", Section:="CustList", Key:="0", setting:=".."
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")

    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        cboSymbol.AddItem LTrim(Mysettings1(intSettings1, 1))
    Next intSettings1
    
End Sub

Private Sub cmdGetQuote_Click()
On Error GoTo ErrorHandler
Dim objhttp As New MSXML.XMLHTTPRequest
Dim compname, datetime1, last1, open1, high1, low1, high52, changeper1, volume1, exchange, change1, marketcap, ask1, bid1, low52, peratio1, pershareprofit1, shareoutstanding1 As String
On Error GoTo ErrorHandler

        symbol = cboSymbol.Text
        
    If cboSymbol.Text = "Symbol" Or cboSymbol.Text = ".." Then
        MsgBox "You must enter a ticker symbol first", vbOKOnly + vbCritical, "Duh!"
        cboSymbol.SetFocus
        Exit Sub
    End If

    If lblConnection.Caption = "You are currently not connected to the internet." Then
        MsgBox "You need to connect to the internet before you can use this program", _
        vbOKOnly + vbCritical, "Not Connected!"
    Exit Sub
    End If
    
        frmWait.Show
        frmWait.SetFocus
    
        objhttp.Open "GET", "http://www.stockpoint.com/quote.asp?Exchange=US&Symbol=" & symbol & "&Company=&x=0&y=0", False
        objhttp.send
        strResponse = objhttp.responseText
                    
        'compname
        retval1 = InStr(1, strResponse, "<FONT COLOR=WHITE><B>") + 21
        retval2 = InStr(1, strResponse, "&nbsp;") - 1
        compname = Mid(strResponse, retval1, Len(strResponse) - retval2)
        retval1 = InStr(1, compname, "&nbsp;")
        compname = Left(compname, retval1 - 1)
        lblSymbol.Caption = compname & "(" & symbol & ")"
        
        'date-time
        retval1 = InStr(1, strResponse, "As of ") + 6
        retval2 = InStr(1, strResponse, "(E.T.)") + 6
        datetime1 = Mid(strResponse, retval1, retval2 - retval1)
        temp = datetime1
        retval1 = InStr(1, datetime1, "&nbsp;&nbsp;") - 1
        datetime1 = Left(datetime1, retval1)
        retval1 = retval1 + 12
        datetime1 = datetime1 & " " & Right(temp, retval1)
        temp = Right(datetime1, 11)
        datetime1 = Left(datetime1, Len(datetime1) - 11)
        datetime1 = datetime1 & " " & Right(temp, 6)
        lblDateTime.Caption = datetime1
        
        'last
        retval1 = InStr(1, strResponse, "Last") + 4
        retval2 = InStr(retval1, strResponse, "<B>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</B>") - 1
        last1 = Left(strResponse, retval1)
        lblLastGet.Caption = last1
        
        'open
        retval1 = InStr(retval1, strResponse, "Open") + 4
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        open1 = Left(strResponse, retval1)
        Text1.Text = open1
        open1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblOpenGet.Caption = open1
          
        'high
        retval1 = InStr(retval1, strResponse, "high1") + 4
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        high1 = Left(strResponse, retval1)
        Text1.Text = high1
        high1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblHighGet.Caption = high1
        
        'low
        retval1 = InStr(retval1, strResponse, "low1") + 6
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        low1 = Left(strResponse, retval1)
        Text1.Text = low1
        low1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblLowGet.Caption = low1
           
        'volume
        retval1 = InStr(retval1, strResponse, "volume") + 6
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        volume1 = Left(strResponse, retval1)
        Text1.Text = volume1
        volume1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblVolumeGet.Caption = volume1
        
        'Per share earning
        retval1 = InStr(retval1, strResponse, "P/Share") + 7
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        pershareprofit1 = Left(strResponse, retval1)
        Text1.Text = pershareprofit1
        pershareprofit1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblPSEGet.Caption = pershareprofit1
                
        'per share outstanding
        retval1 = InStr(retval1, strResponse, "Outstanding") + 11
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        shareoutstanding1 = Left(strResponse, retval1)
        Text1.Text = shareoutstanding1
        shareoutstanding1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblPSOGet.Caption = shareoutstanding1
                
        'P/E ratio
        retval1 = InStr(retval1, strResponse, "Ratio") + 5
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        peratio1 = Left(strResponse, retval1)
        Text1.Text = peratio1
        peratio1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblPERGet.Caption = peratio1
        
        'change
        retval1 = 0
        retval1 = InStr(1, strResponse, "#339933")
        If retval1 <> 0 Then
        retval1 = retval1 + 9
        ElseIf InStr(1, strResponse, "RED") <> 0 Then
        retval1 = InStr(1, strResponse, "RED") + 6
        ans = 1
        Else
        retval1 = InStr(1, strResponse, "SIZE=-1>") + 8
        ans = 2
        End If
        strResponse = Mid(strResponse, retval1, Len(strResponse) - retval1)
        retval2 = InStr(1, strResponse, "</FONT>")
        change1 = Left(strResponse, retval2 - 1)
        If ans = 2 Then change1 = Mid(change1, 3, Len(change1) - 3)
        If ans = 1 Then change1 = "-" & change1
        lblChangeGet.Caption = change1
        
        '% change
        retval1 = 0
        retval1 = InStr(1, strResponse, "#339933")
        If retval1 <> 0 Then
        retval1 = retval1 + 9
        ElseIf InStr(1, strResponse, "RED") <> 0 Then
        retval1 = InStr(1, strResponse, "RED") + 6
        ans = 1
        Else
        retval1 = InStr(1, strResponse, "SIZE=-1>") + 8
        ans = 2
        End If
        strResponse = Mid(strResponse, retval1, Len(strResponse) - retval1)
        retval2 = InStr(1, strResponse, "</FONT>")
        changeper1 = Left(strResponse, retval2 - 1)
        Text1.Text = changeper1
        If ans = 2 Then changeper1 = Mid(changeper1, 3, Len(changeper1) - 3)
        If ans = 1 Then
        changeper1 = "-" & changeper1
        ans = 0
        End If
        lblChangepGet.Caption = changeper1
                
        '52high
        retval1 = InStr(retval1, strResponse, "High") + 4
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        high52 = Left(strResponse, retval1)
        Text1.Text = high52
        high52 = Right(Text1.Text, Len(Text1.Text) - 2)
        lbl52HighGet.Caption = high52
        
        '52low
        retval1 = InStr(retval1, strResponse, "Low") + 3
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        low52 = Left(strResponse, retval1)
        Text1.Text = low52
        low52 = Right(Text1.Text, Len(Text1.Text) - 2)
        lbl52LowGet.Caption = low52
        
        'bid
        retval1 = InStr(retval1, strResponse, "bid1") + 3
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        bid1 = Left(strResponse, retval1)
        Text1.Text = bid1
        bid1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblBidGet.Caption = bid1
        
        'ask
        retval1 = InStr(retval1, strResponse, "ask1") + 3
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        ask1 = Left(strResponse, retval1)
        Text1.Text = ask1
        ask1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblAskGet.Caption = ask1
        
        'Market Cap
        retval1 = InStr(retval1, strResponse, "Capitalization") + 14
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        marketcap = Left(strResponse, retval1)
        Text1.Text = marketcap
        marketcap = Right(Text1.Text, Len(Text1.Text) - 2)
        lblMCGet.Caption = marketcap
        
        'exchange
        retval1 = InStr(retval1, strResponse, "Exchange") + 8
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        exchange = Left(strResponse, retval1)
        Text1.Text = exchange
        exchange = Right(Text1.Text, Len(Text1.Text) - 2)
        lblSEGet.Caption = exchange
        
        Unload frmWait
        Exit Sub
        
ErrorHandler:
        MsgBox Err.Description & vbCrLf & Err.Number
        Unload frmWait
        Resume Next
    
    
End Sub

Public Sub GetGraph(GraphType As String, symbol As String)
Dim url As String
Dim Bite() As Byte
On Error GoTo ErrorHandler
Top:

symbol = LCase(cboSymbol)
        
    If cboGraph.Text = "1 year big" Then
        url = "http://chart.yahoo.com/c/1y/" & Left(symbol, 1) & "/" & symbol & ".gif"
    ElseIf cboGraph.Text = "2 year big" Then
        url = "http://chart.yahoo.com/c/2y/" & Left(symbol, 1) & "/" & symbol & ".gif"
    ElseIf cboGraph.Text = "3 months big" Then
        url = "http://chart.yahoo.com/c/3m/" & Left(symbol, 1) & "/" & symbol & ".gif"
    ElseIf cboGraph.Text = "6 months small" Then
        url = "http://chart.yahoo.com/c/0b/" & Left(symbol, 1) & "/" & symbol & ".gif"
    End If
    
        Bite() = Inet1.OpenURL(url, icByteArray) ' Download picture.s = Bilden()
        x = Bite()
        
    If Len(x) <> 75 Then
        Open "C:\graph.gif" For Binary Access Write As #1 ' Save the file.
        Put #1, , Bite()
        Close #1
    Else
    End If
    
        frmGraph.Picture1.Picture = LoadPicture("C:\graph.gif")
        frmGraph.Show
        
ErrorHandler:
    Resume Next
    If Err.Number = 35764 Then
        MsgBox "Still executiong last request", vbOKOnly, "Oops"
    GoTo Top
    End If
    
End Sub

Private Sub cmdViewGraph_Click()
    If cboSymbol.Text = "" Or cboSymbol = "Symbol" Or IsNumeric(cboSymbol.Text) = True Or cboGraph.Text = "" Then
        MsgBox "Could not process your request.  Check that you have entered a" & vbCrLf _
        & "valid ticker symbol and that you have selected a valid graph type.", vbOKOnly + vbCritical, "Oops"
        Exit Sub
    Else
        GetGraph cboGraph.Text, cboSymbol.Text
    End If

    
End Sub

Private Sub mnuExit_Click()
        Unload Me
End Sub

Private Sub mnuYourCompany_Click()
        'frmAbout.Show
End Sub

Private Sub mnuThisProgram_Click()
        'frmAboutThisProgram.Show
End Sub

'Private Sub GraphDownloadCompleted(filename As String)
        'frmGraph.Picture1.Picture = LoadPicture("C:\graph.gif")
        'frmGraph.Show
'End Sub

Private Sub cmdAddSymbol_Click()
    
    If cboSymbol.Text = "Symbol" Or cboSymbol.Text = ".." Then
        MsgBox "You must enter a ticker symbol first", vbOKOnly + vbCritical, "Duh!"
        cboSymbol.SetFocus
    Exit Sub
    End If
    
        'Get Customer List form registry
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")
        'Set the array upper and lower parameters
        
    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        Counter = Mysettings1(intSettings1, 0)
    Next intSettings1
    
        sCustKey = "CustList"
        sCustVal = RTrim(cboSymbol.Text)
        intSet1 = Counter + 1
        'Saves combo text in CustList Registry
        SaveSetting AppName:="Header", Section:="CustList", Key:=intSet1, setting:=RTrim(sCustVal)
        Counter1 = Counter1 + 1
        cboSymbol.Clear
        'Clears then fills CboSymbol with new list
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")
        
    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        cboSymbol.AddItem Mysettings1(intSettings1, 1)
    Next intSettings1
    
        cboSymbol.ListIndex = 0
End Sub

Private Sub cmdRemoveSymbol_Click()

    If cboSymbol.Text = "Symbol" Or cboSymbol.Text = ".." Then
        MsgBox "You must enter a ticker symbol first", vbOKOnly + vbCritical, "Duh!"
        cboSymbol.SetFocus
    Exit Sub
    End If
    
    'Loop through the registry looking for a match to the cboSymbol.Text
    'If it is found delete it from the registry
    
    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        j = Mysettings1(intSettings1, 1)
        j1 = Mysettings1(intSettings1, 0)
        
    If j = cboSymbol.Text Then GoTo Del
    
    Next intSettings1

Del:
        DeleteSetting "Header", "CustList", j1
        cboSymbol.Clear 'Clear before re-loading combo with new values
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")
        'Re-Read the registry and fill cboSymbol
        
   For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        cboSymbol.AddItem LTrim(Mysettings1(intSettings1, 1))
   Next intSettings1
   
        cboSymbol.ListIndex = 0
    
End Sub

