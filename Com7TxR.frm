VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox WTscale2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   12480
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox WTscale1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   12480
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.Timer ReadTimer 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   480
      Top             =   4080
   End
   Begin VB.TextBox RAWText 
      Height          =   1095
      Left            =   7920
      TabIndex        =   5
      Top             =   6480
      Width           =   6255
   End
   Begin VB.TextBox TextSummedSCL 
      Height          =   735
      Left            =   7920
      TabIndex        =   4
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox TextSCL3 
      Height          =   615
      Left            =   7920
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox TextSCL2 
      Height          =   615
      Left            =   7920
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox TextSCL1 
      Height          =   615
      Left            =   7920
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox RxText 
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   6480
      Width           =   6975
   End
   Begin MSCommLib.MSComm Comm5 
      Left            =   14880
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MSComm1_OnComm()

End Sub

Private Sub Form_Load()
 Dim Instring As String
 Dim AnyString As String
 Dim SCALe3String, SCALe3MOString As String
 Dim SCALe2String, SCALe2MOString As String
 Dim SCALe1String As String
 Dim ASC02Str As String
 Dim summedWT As String

 If Comm5.PortOpen Then
        Comm5.PortOpen = False
        MsgBox "COM port closed."
        Else
        Comm5.InputLen = 0   ' Tell the control to read entire buffer when Input is used.
        Comm5.PortOpen = True
        Comm5.Output = "#1  49180KG" & Chr$(13)            'ascii 2,#,1,7 digits weight,KG,CR
        Comm5.Output = "#2  49180KG" & Chr$(13)             '#,2,7 digits weight,KG,CR
        Comm5.Output = "#3  80000KG" & Chr$(13)             '#,3,7 digits weight,KG,CR
        Comm5.Output = "#0 178360KG" & Chr$(13)             '#,0,7 digits weight,KG,CR
  ''  Do
   ''   DoEvents
    'AnyString = Comm5.Input
'' Buffer$ = Buffer$ & Comm5.Input
 
' Instring = Buffer$
 ''RxText.Text = Buffer$

 ' ASC02Str = Left(AnyString, 1)    ' Returns ascii 02
  'If ASC02Str = Chr$(2) Then TextSCL1.Text = AnyString
   'Buffer$ = ""
   
   
   'Loop Until InStr(Buffer$, "KG" & vbCrLf)
 '  Loop Until InStr(Buffer$, "KG" & vbCr)
   ''Loop Until Left(Comm5.Input, 1) = Chr$(2)
   ' Read the "OK" response data in the serial port.
    ''Comm5.PortOpen = False      ' Close the serial port.
    
    ''RAWText.Text = Left(RxText.Text, 48)
   ' RAWText.Text = Left(Instring, 48)
    
    ''SCALe1String = RAWText.Text
    ''TextSCL1.Text = Left(SCALe1String, 13)
    
    ''SCALe2String = Left(RxText.Text, 25)
    ''SCALe2MOString = Right(SCALe2String, 13)
    ''TextSCL2.Text = SCALe2MOString
    
    ''SCALe3String = Right(RAWText.Text, 23)
    ''SCALe3MOString = Left(SCALe3String, 12)
    ''TextSCL3.Text = SCALe3MOString
    
    ''summedWT = Right(RAWText.Text, 12) ' Summed Wt. Returns
    ''TextSummedSCL.Text = summedWT
    
    ''Buffer$ = ""
    ''RxText.Text = ""
    ''RAWText.Text = ""
    MsgBox "COM port opened."
    
End If
ReadTimer.Enabled = True
     ' Buffer to hold input string
  
   
  ' MSComm1.CommPort = 1         ' Use COM1.
   ' MSComm1.Settings = "9600,N,8,1"          ' 9600 baud, no parity, 8 data, and 1 stop bit.
   '   MSComm1.PortOpen = True        ' Open the port.
 '   MSComm1.Output = "ATV1Q0" & Chr$(13)  ' Send the attention command to the modem.
   ' Ensure that
   ' the modem responds with "OK".
   ' Wait for data to come back to the serial port.
   
'With Comm5
       ' .CommPort = 5          ' COM5
       ' .Settings = "9600,N,8,1" ' Baud rate, parity, data bits, stop bits
 '       .PortOpen = True
  '  End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Comm5.PortOpen Then
        Comm5.PortOpen = False
        MsgBox "COM port closed."
    End If
End Sub

Private Sub ReadTimer_Timer()
'AnyString = Comm5.Input
 Buffer$ = Buffer$ & Comm5.Input
 
' Instring = Buffer$
 RxText.Text = Buffer$

 ' ASC02Str = Left(AnyString, 1)    ' Returns ascii 02
  'If ASC02Str = Chr$(2) Then TextSCL1.Text = AnyString
   'Buffer$ = ""
   
   
   'Loop Until InStr(Buffer$, "KG" & vbCrLf)
 '  Loop Until InStr(Buffer$, "KG" & vbCr)
 '''  If (Left(Comm5.Input, 1) = Chr$(2)) Then
   ' Read the "OK" response data in the serial port.
   ' Comm5.PortOpen = False      ' Close the serial port.
   
   ''' ReadTimer.Enabled = False
    
 ''   RAWText.Text = Left(RxText.Text, 48)
   ' RAWText.Text = Left(Instring, 48)
    
   '' SCALe1String = RAWText.Text
   '' TextSCL1.Text = Left(SCALe1String, 13)
    
   '' SCALe2String = Left(RxText.Text, 25)
   '' SCALe2MOString = Right(SCALe2String, 13)
   '' TextSCL2.Text = SCALe2MOString
    
    '''''SCALe3String = Right(RAWText.Text, 23)
    ''SCALe3MOString = Left(SCALe3String, 12)
    ''TextSCL3.Text = SCALe3MOString
    
    ''summedWT = Right(RAWText.Text, 12) ' Summed Wt. Returns
    ''TextSummedSCL.Text = summedWT
    
   '' Buffer$ = ""
   ' RxText.Text = ""
  '  RAWText.Text = ""
  '''  End If
End Sub

Private Sub RxText_Change()
  RAWText.Text = Left(RxText.Text, 48)
   ' RAWText.Text = Left(Instring, 48)
    
    SCALe1String = RAWText.Text
    TextSCL1.Text = Left(SCALe1String, 13)
    
    SCALe2String = Left(RxText.Text, 25)
    SCALe2MOString = Right(SCALe2String, 13)
    TextSCL2.Text = SCALe2MOString
    
    SCALe3String = Right(RAWText.Text, 23)
    SCALe3MOString = Left(SCALe3String, 12)
    TextSCL3.Text = SCALe3MOString
    
    summedWT = Right(RAWText.Text, 12) ' Summed Wt. Returns
    TextSummedSCL.Text = summedWT
End Sub

Private Sub TextSCL1_Change()
Dim SCALe1Wt As String
SCALe1Wt = Left(TextSCL1.Text, 10)
WTscale1.Text = Right(SCALe1Wt, 7)
End Sub

Private Sub TextSCL2_Change()
Dim SCALe2Wt As String
SCALe2Wt = Left(TextSCL2.Text, 10)
WTscale2.Text = Right(SCALe2Wt, 7)
End Sub
