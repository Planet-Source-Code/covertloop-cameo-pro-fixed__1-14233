VERSION 5.00
Object = "{DF6D6558-5B0C-11D3-9396-008029E9B3A6}#1.0#0"; "EZVIDC60.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Cameo Pro"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   2760
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1200
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Capturer 
      Left            =   2520
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   4920
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   255
   End
   Begin VB.HScrollBar FPSscroll 
      Height          =   135
      LargeChange     =   10
      Left            =   2880
      Max             =   100
      Min             =   1
      TabIndex        =   5
      Top             =   1680
      Value           =   15
      Width           =   1455
   End
   Begin vbVidC60.ezVidCap Cam 
      Height          =   2160
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   3810
      AutoSize        =   0   'False
      BorderStyle     =   0
      Preview         =   0   'False
      AbortLeftMouse  =   0
      AbortRightMouse =   0
   End
   Begin MSWinsockLib.Winsock Send 
      Left            =   1680
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connections - 0"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Dragger 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Cam Address"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "127.0.0.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      MouseIcon       =   "Form1.frx":2A24C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FPS - 15"
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Stream"
      Height          =   195
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Image Image4 
      Height          =   180
      Left            =   2880
      Picture         =   "Form1.frx":2A39E
      Top             =   1200
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Format"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   900
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   2880
      Picture         =   "Form1.frx":2A590
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   540
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   2880
      Picture         =   "Form1.frx":2BE92
      Top             =   480
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   4080
      Picture         =   "Form1.frx":2D794
      Top             =   0
      Width           =   315
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   3720
      Picture         =   "Form1.frx":2DD16
      Top             =   0
      Width           =   315
   End
   Begin VB.Image MinDown 
      Height          =   300
      Left            =   720
      Picture         =   "Form1.frx":2E258
      Top             =   4800
      Width           =   315
   End
   Begin VB.Image MinUp 
      Height          =   300
      Left            =   360
      Picture         =   "Form1.frx":2E79A
      Top             =   4800
      Width           =   315
   End
   Begin VB.Image CloseDown 
      Height          =   315
      Left            =   1200
      Picture         =   "Form1.frx":2ECDC
      Top             =   4440
      Width           =   315
   End
   Begin VB.Image CloseUp 
      Height          =   315
      Left            =   840
      Picture         =   "Form1.frx":2F25E
      Top             =   4440
      Width           =   315
   End
   Begin VB.Image CheckOn 
      Height          =   180
      Left            =   480
      Picture         =   "Form1.frx":2F7E0
      Top             =   4440
      Width           =   180
   End
   Begin VB.Image CheckOff 
      Height          =   180
      Left            =   240
      Picture         =   "Form1.frx":2F9D2
      Top             =   4440
      Width           =   180
   End
   Begin VB.Image ButtonDown 
      Height          =   330
      Left            =   240
      Picture         =   "Form1.frx":2FBC4
      Top             =   3960
      Width           =   1425
   End
   Begin VB.Image ButtonUp 
      Height          =   330
      Left            =   240
      Picture         =   "Form1.frx":314C6
      Top             =   3600
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public i As Long
Public q As Long
Public M As Long

Private Declare Function ConvertBMPtoJPG Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean, ByVal JPGCompressQuality As Integer, ByVal blnKeepBMP As Boolean) As Integer

Private Declare Function SendMessage Lib "User32" _
Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "User32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Const OK = 0
Private Const InputFileMissing = -1
Private Const OutputFileAlreadyExists = -2

Dim strSource As String
Dim strSource2 As String
Dim strDestination As String

Private Sub Capturer_Timer()
'Capture The Image
If App.Path = "C:\" Then Cam.SaveDIB "C:\Cap.bmp"
If App.Path <> "C:\" Then Cam.SaveDIB App.Path & "\Cap.bmp"


'Set Settings To Convert To JPG
If App.Path = "C:\" Then strSource = "C:\Cap.bmp"
If App.Path <> "C:\" Then strSource = App.Path & "\Cap.bmp"

If App.Path = "C:\" Then strSource2 = "C:\Cap.bmp"
If App.Path <> "C:\" Then strSource2 = App.Path & "\Cap.bmp"

If App.Path = "C:\" Then strDestination = "C:\Cap.jpg"
If App.Path <> "C:\" Then strDestination = App.Path & "\Cap.jpg"

'Convert To JPG
Dim retval As Integer
retval = ConvertBMPtoJPG(strSource2, strDestination, True, 100, True)
End Sub

Private Sub Dragger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Allow Dragging Of Form Without Title Bar
Dim lngReturnValue As Long

If Button = 1 Then
    Call ReleaseCapture
    lngReturnValue = SendMessage(Form1.hWnd, WM_NCLBUTTONDOWN, _
    HTCAPTION, 0&)
End If
End Sub


Private Sub Form_Load()
Show

'Refresh Sockets For Use
Send.Close
Winsock1.Close
End Sub


Private Sub FPSscroll_Change()
'Set Scrollbar Values To Preview/Capture Rate
'Of The Capture Control
Cam.PreviewRate = FPSscroll.Value
Cam.CaptureRate = FPSscroll.Value

'Show Scrollbar Value as Frames Per Second
Label4.Caption = "FPS - " & FPSscroll.Value
End Sub


Private Sub FPSscroll_Scroll()
'Set Scrollbar Values To Preview/Capture Rate
'Of The Capture Control
Cam.PreviewRate = FPSscroll.Value
Cam.CaptureRate = FPSscroll.Value

'Show Scrollbar Value as Frames Per Second
Label4.Caption = "FPS - " & FPSscroll.Value
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = CloseDown.Picture
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Image1.Picture = CloseUp.Picture
Cam.Preview = False
Send.Close
Unload Me
End
End Sub


Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ButtonDown.Picture
End Sub


Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ButtonUp.Picture

'Open The "Source" Dialog
Cam.ShowDlgVideoSource
End Sub


Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = ButtonDown.Picture
End Sub


Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = ButtonUp.Picture

'Open The "Format" Dialog
Cam.ShowDlgVideoFormat
End Sub


Private Sub Image4_Click()
On Error Resume Next

'Check The Users "Video Preview" Choice
If Check1.Value = 1 Then
    Check1.Value = 0
    Image4.Picture = CheckOff.Picture
    Cam.Preview = False
    Send.Close
    Exit Sub
End If

'Check The Users "Video Preview" Choice
If Check1.Value = 0 Then
    Check1.Value = 1
    Image4.Picture = CheckOn.Picture
    Cam.Preview = True
    
    'Start Capturing An Image From The
    'Video Every Second (1000 ms)
    Capturer.Interval = 1000
    
    'Refresh And Open The HTML Port For
    'Users To Connect To (Port 80)
    Send.Close
    Send.LocalPort = CLng(80)
    Send.Listen
    Exit Sub
End If
End Sub


Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Picture = MinDown.Picture
End Sub


Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Picture = MinUp.Picture
Form1.WindowState = 1
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ButtonDown.Picture
End Sub


Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = ButtonUp.Picture

'Show "Source" Dialog
Cam.ShowDlgVideoSource
End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = ButtonDown.Picture
End Sub


Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = ButtonUp.Picture

'Show "Format" Dialog
Cam.ShowDlgVideoFormat
End Sub


Private Sub Label5_Click()
'Alert User
Dim AlertUser
AlertUser = MsgBox("If your web browser is covering up the live image," & Chr$(13) & "the picture may not display properly." & Chr$(13) & Chr$(13) & "Note:  May not work for 'Netscape' users.", vbInformation + vbOKOnly, "Cameo-Pro")

'Open Default Web Browser For
'Proof Preview Of Live Image
Result = Shell("start.exe http://" & CurrentIP(True), vbHide)
End Sub

Private Sub Send_ConnectionRequest(ByVal requestID As Long)
'Allow Winsock1 To Accept The Request, Which Will
'Keep This Winsock Control Open For New Requests
M = M + 1
Label7.Caption = "Connections - " & M
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Timer1_Timer()
'Update The IP Number Visibly
'Every Second, In Case Of Change
Label5.Caption = "http://" & CurrentIP(True)
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Free File
f = FreeFile
temp = ""

'Retrieve The Path Of The Captured Image
'And Convert It To A Short Path Name
'Ex:
'Long Path Name - C:\Program Files
'Short Path Name - C:\Progra~1
Dim ConvertedPath As String
Dim sFile As String, sShortFile As String * 67
Dim lRet As Long
If App.Path = "C:\" Then sFile = "C:\Cap.jpg"
If App.Path <> "C:\" Then sFile = App.Path & "\Cap.jpg"
lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
sFile = Left(sShortFile, lRet)
ConvertedPath = sFile

'Open Captured/Converted Image
Open ConvertedPath For Binary As #f
    temp = Input(FileLen(ConvertedPath), #f)
Close #f

getimg = temp

'Send The Picture
Winsock1.SendData getimg
End Sub

Private Sub Winsock1_SendComplete()
'Close The Socket
Winsock1.Close
End Sub


