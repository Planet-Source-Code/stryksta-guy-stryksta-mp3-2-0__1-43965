VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd2 
      Left            =   4920
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer timerplay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   2160
   End
   Begin VB.HScrollBar VSvolume 
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Left            =   2640
      Top             =   4320
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   360
      Picture         =   "Form1.frx":1042
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   228
      TabIndex        =   3
      Top             =   720
      Width           =   3420
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   3840
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1680
      Top             =   3840
   End
   Begin VB.PictureBox PicDisplay 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   2685
      TabIndex        =   1
      ToolTipText     =   "Song  Playng"
      Top             =   360
      Width           =   2685
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stryksta  Player 2.0"
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   600
      Picture         =   "Form1.frx":2B3C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   228
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   600
      Picture         =   "Form1.frx":4636
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   228
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   600
      Picture         =   "Form1.frx":6130
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   228
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   3420
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   240
      TabIndex        =   9
      Tag             =   "sx"
      Top             =   1030
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   6085
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Filename"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image iexit 
      Height          =   255
      Left            =   4690
      Picture         =   "Form1.frx":7C2A
      Top             =   50
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00D67335&
      X1              =   200
      X2              =   4800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00D67335&
      X1              =   4800
      X2              =   4800
      Y1              =   960
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D67335&
      X1              =   200
      X2              =   4800
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D67335&
      X1              =   190
      X2              =   190
      Y1              =   960
      Y2              =   2760
   End
   Begin VB.Image next2 
      Height          =   225
      Left            =   3840
      Picture         =   "Form1.frx":7FE0
      Tag             =   "rx"
      Top             =   3120
      Width           =   210
   End
   Begin VB.Image next1 
      Height          =   225
      Left            =   3600
      Picture         =   "Form1.frx":82B6
      Tag             =   "rx"
      Top             =   3120
      Width           =   210
   End
   Begin VB.Image prev2 
      Height          =   225
      Left            =   2520
      Picture         =   "Form1.frx":858C
      Tag             =   "rx"
      Top             =   3240
      Width           =   210
   End
   Begin VB.Image prev1 
      Height          =   225
      Left            =   2280
      Picture         =   "Form1.frx":8862
      Tag             =   "rx"
      Top             =   3240
      Width           =   210
   End
   Begin VB.Image Inext 
      Height          =   225
      Left            =   3645
      Picture         =   "Form1.frx":8B38
      Tag             =   "rx"
      Top             =   60
      Width           =   210
   End
   Begin VB.Image iopen 
      Height          =   225
      Left            =   4080
      Picture         =   "Form1.frx":8E0E
      Tag             =   "rx"
      Top             =   75
      Width           =   210
   End
   Begin VB.Image pause 
      Height          =   225
      Left            =   3390
      Picture         =   "Form1.frx":90E4
      Tag             =   "rx"
      Top             =   60
      Width           =   210
   End
   Begin VB.Image stopb 
      Height          =   225
      Left            =   3135
      Picture         =   "Form1.frx":93BA
      Tag             =   "rx"
      ToolTipText     =   "Stop"
      Top             =   60
      Width           =   210
   End
   Begin VB.Image prev 
      Height          =   225
      Left            =   2625
      Picture         =   "Form1.frx":9690
      Tag             =   "rx"
      Top             =   60
      Width           =   210
   End
   Begin VB.Image play 
      Height          =   225
      Left            =   2880
      Picture         =   "Form1.frx":9966
      Tag             =   "rx"
      Top             =   60
      Width           =   210
   End
   Begin VB.Image iopen2 
      Height          =   225
      Left            =   2760
      Picture         =   "Form1.frx":9C3C
      Tag             =   "rxy"
      Top             =   3840
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image iopen1 
      Height          =   225
      Left            =   2520
      Picture         =   "Form1.frx":9F12
      Tag             =   "rxy"
      Top             =   3840
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Iexit1 
      Height          =   255
      Left            =   3360
      Picture         =   "Form1.frx":A1E8
      Tag             =   "rxy"
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Iexit2 
      Height          =   255
      Left            =   3120
      Picture         =   "Form1.frx":A59E
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00 - 00:00"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image FTop3 
      Height          =   330
      Left            =   3390
      Picture         =   "Form1.frx":A954
      Tag             =   "rx"
      Top             =   0
      Width           =   1620
   End
   Begin VB.Image FTop2 
      Height          =   330
      Left            =   2580
      Picture         =   "Form1.frx":C56E
      Stretch         =   -1  'True
      Tag             =   "sxrx"
      Top             =   0
      Width           =   855
   End
   Begin VB.Image FRight 
      Height          =   3450
      Left            =   4950
      Picture         =   "Form1.frx":D3C8
      Stretch         =   -1  'True
      Tag             =   "syrx"
      Top             =   330
      Width           =   60
   End
   Begin VB.Image FButtom 
      Height          =   45
      Left            =   60
      Picture         =   "Form1.frx":DC1A
      Stretch         =   -1  'True
      Tag             =   "sxry"
      Top             =   2865
      Width           =   5010
   End
   Begin VB.Image pause2 
      Height          =   225
      Left            =   1200
      Picture         =   "Form1.frx":E688
      Tag             =   "ry"
      Top             =   5040
      Width           =   210
   End
   Begin VB.Image pause1 
      Height          =   225
      Left            =   960
      Picture         =   "Form1.frx":E95E
      Tag             =   "ry"
      Top             =   5040
      Width           =   210
   End
   Begin VB.Image stopb2 
      Height          =   225
      Left            =   840
      Picture         =   "Form1.frx":EC34
      Tag             =   "ry"
      Top             =   5520
      Width           =   210
   End
   Begin VB.Image stopb1 
      Height          =   225
      Left            =   600
      Picture         =   "Form1.frx":EF0A
      Tag             =   "ry"
      Top             =   5520
      Width           =   210
   End
   Begin VB.Image play2 
      Height          =   225
      Left            =   360
      Picture         =   "Form1.frx":F1E0
      Tag             =   "ry"
      Top             =   5040
      Width           =   210
   End
   Begin VB.Image play1 
      Height          =   225
      Left            =   0
      Picture         =   "Form1.frx":F4B6
      Tag             =   "ry"
      Top             =   5040
      Width           =   210
   End
   Begin VB.Image resize2 
      Height          =   255
      Left            =   4920
      Tag             =   "rxy"
      Top             =   2880
      Width           =   375
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   675
      Left            =   3960
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1185
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   12648384
      DisplayForeColor=   12648384
      DisplayMode     =   1
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   -1  'True
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   -1  'True
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Image Ftop 
      Height          =   330
      Left            =   0
      Picture         =   "Form1.frx":F78C
      Tag             =   "sx"
      Top             =   0
      Width           =   2580
   End
   Begin VB.Image right 
      Height          =   4440
      Left            =   5665
      Picture         =   "Form1.frx":12426
      Stretch         =   -1  'True
      Tag             =   "syrx"
      Top             =   0
      Width           =   45
   End
   Begin VB.Image FLeft 
      Height          =   2580
      Left            =   0
      Picture         =   "Form1.frx":13278
      Stretch         =   -1  'True
      Tag             =   "sy"
      Top             =   330
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim positioned As Boolean, down As Boolean, counting As Integer, fullscreen As Boolean
Dim oldtop As Integer, oldleft As Integer, oldwidth As Integer, oldheight As Integer, testx As Long, testy As Long, p As Integer
Private Sizedata() As CtlAdj
Dim sFileName As String
Dim mystr As String
Private Sub Command1_Click()

End Sub

 Private Sub Form_Load()
Dim i As Integer, TestHighlight As ItemColourType
RegisterForm Me, Me.width, Me.height, Sizedata()

    ModLVSubClass.Attach Me.hWnd, ListView1
    
    ModLVSubClass.UseCustomHighLight True
    
    TestHighlight.BackGround = RGB(53, 115, 214)
    TestHighlight.ForeGround = RGB(255, 255, 255)
    
    ModLVSubClass.SetHighLightColour TestHighlight
    
    
    ModLVSubClass.SetCustomColour TestHighlight
  
End Sub
Private Sub Form_Resize()
    ResizeOMatic Me, Sizedata()
End Sub

Private Sub Image10_Click()
If MediaPlayer1.CurrentPosition < MediaPlayer1.SelectionEnd - 5 Then MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition + 5
End Sub

Private Sub Image11_Click()
On Error Resume Next
If MediaPlayer1.CurrentPosition < 5 Then MediaPlayer1.CurrentPosition = 0
If MediaPlayer1.CurrentPosition > 5 Then MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition - 5
End Sub

Private Sub Ftop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNSnap Me, Button, X, Y
    prev.Picture = prev1.Picture
    play.Picture = play1.Picture
    stopb.Picture = stopb1.Picture
End Sub
Private Sub FTop2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prev.Picture = prev1.Picture
    play.Picture = play1.Picture
    stopb.Picture = stopb1.Picture
End Sub
Private Sub FTop3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iexit.Picture = Iexit1.Picture
pause.Picture = pause1.Picture
Inext.Picture = next1.Picture
iopen.Picture = iopen1.Picture
End Sub
Private Sub Iexit_Click()
Unload Me
End Sub
Private Sub Iexit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iexit.Picture = Iexit2.Picture
End Sub

Private Sub iopen_Click()
 Dim strFile As String
 Dim ex
  Dim i%, tempname As String
On Error GoTo ClickError
  cd.DialogTitle = "Open Audio"
cd.Filter = "Supported Audio formats (*.mp3;*.wav;*.cda;*.divx;*.asf;*.wmv;*.dat) *.mpg;*.mpeg;*.avi;*.divx;*.asf;*.wmv;*.dat|*.mp3;*.wav;*.cda;*.wma;*.asf;*.wmv;*.dat|All files|*.*"
cd.FilterIndex = 1
cd.Action = 1
MediaPlayer1.FileName = cd.FileName
ReadID3.LoadMp3File cd.FileName
 MediaPlayer1.DisplayMode = mpTime
        sectomin (MediaPlayer1.Duration)
        tottime = MediaPlayer1.Duration
       MediaPlayer1.DisplayMode = mpFrames
        totfps = MediaPlayer1.Duration
        PlaybackSpeed = totfps / tottime
        PlaybackSpeed = Int(PlaybackSpeed)
        mhour = shour: mmin = smin: msec = ssec
        sectomin (MediaPlayer1.Duration)
        tottime = MediaPlayer1.Duration
        'Label2.Caption = "00:00:00 - " & mhour & ":" & mmin & ":" & msec
        Label24.Caption = Text2.Text
        Text3.Text = ReadID3.Artist
        Text4.Text = ReadID3.Title
        Text2.Text = Text3.Text & " - " & Text4.Text
         'Label24.Caption = Text2.Text
        timerplay.Enabled = True
        Timer1.Enabled = True
        Picture3.Enabled = True
        Picture3.Picture = Picture1.Picture
ClickError:
    If Err.Number <> 0 Then Exit Sub
End Sub
Private Sub Iopen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iopen.Picture = iopen2.Picture
End Sub

Private Sub ListView1_DblClick()
Dim ouy As Integer
ouy = ListView1.SelectedItem.index
MediaPlayer1.FileName = ListView1.ListItems.Item(ouy).SubItems(3)
MediaPlayer1.CurrentPosition = 0
Picture3.Picture = Picture1.Picture
Picture3.Enabled = False
Picture4.Picture = Picture1.Picture
Picture3.Enabled = True
MediaPlayer1.play
Label24.Caption = ListView1.ListItems.Item(ouy).SubItems(1)
Label2.Caption = "00:00 - " & ":" & ListView1.ListItems.Item(ouy).SubItems(2)
End Sub

Private Sub next_Click()
'
End Sub
Private Sub inext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Inext.Picture = next2.Picture
End Sub
Private Sub pause_Click()
On Error GoTo ClickError

If MediaPlayer1.PlayState = mpPaused Then
    MediaPlayer1.play
Else
    MediaPlayer1.pause
End If

ClickError:
    If Err.Number <> 0 Then Exit Sub
End Sub
Private Sub pause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    stopb.Picture = stopb1.Picture
    pause.Picture = pause2.Picture
    Inext.Picture = next1.Picture
End Sub
Private Sub play_Click()
On Error GoTo ClickError

Dim ouy As Integer
ouy = ListView1.SelectedItem.index
MediaPlayer1.FileName = ListView1.ListItems.Item(ouy).SubItems(2)
MediaPlayer1.play

ClickError:
    If Err.Number <> 0 Then Exit Sub
End Sub
Private Sub play_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    stopb.Picture = stopb1.Picture
    play.Picture = play2.Picture
    prev.Picture = prev1.Picture
End Sub

Private Sub prev_Click()
'
End Sub
Private Sub prev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pause.Picture = pause1.Picture
    prev.Picture = prev2.Picture
End Sub
Private Sub resize2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
End If
End Sub
Private Sub Image6_Click()
Unload Me
End Sub

Private Sub stopb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    play.Picture = play1.Picture
    stopb.Picture = stopb2.Picture
    pause.Picture = pause1.Picture
End Sub
Private Sub stopb_Click()
MediaPlayer1.Stop
MediaPlayer1.CurrentPosition = 0
Picture3.Picture = Picture1.Picture
Picture3.Enabled = False
Picture4.Picture = Picture1.Picture
End Sub

Private Sub Timer1_Timer()
Dim ouy As Integer
ouy = ListView1.SelectedItem.index
tempval = MediaPlayer1.CurrentPosition / MediaPlayer1.SelectionEnd
tempval = tempval * Picture3.ScaleWidth
tempval = tempval
Call BitBlt(Picture4.hDC, 0, 0, tempval, Picture3.ScaleHeight, Picture5.hDC, 0, 0, vbSrcCopy)
Call BitBlt(Picture3.hDC, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture4.hDC, 0, 0, vbSrcCopy)
Picture3.Refresh
temptime = MediaPlayer1.CurrentPosition / PlaybackSpeed
sectomin (Int(temptime))

Label2.Caption = smin & ":" & ssec & " - " & ListView1.ListItems.Item(ouy).SubItems(2)
End Sub

Private Sub Timer2_Timer()
'
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If down = True Then
If X < 0 Then X = 0
If X > Picture3.ScaleWidth Then X = Picture3.ScaleWidth
currentpos = (X / Picture3.ScaleWidth) * 100
MediaPlayer1.CurrentPosition = MediaPlayer1.SelectionEnd * (currentpos / 100)
Picture4.Cls
Call BitBlt(Picture4.hDC, 0, 0, X, Picture3.ScaleHeight, Picture5.hDC, 0, 0, vbSrcCopy)
Call BitBlt(Picture3.hDC, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture4.hDC, 0, 0, vbSrcCopy)
Picture3.Refresh
End If
pause.Picture = pause1.Picture
stopb.Picture = stopb1.Picture
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
down = True
looking = True
currentpos = (X / Picture3.ScaleWidth) * 100
Picture4.Cls
MediaPlayer1.CurrentPosition = MediaPlayer1.SelectionEnd * (currentpos / 100)
Call BitBlt(Picture4.hDC, 0, 0, X, Picture3.ScaleHeight, Picture5.hDC, 0, 0, vbSrcCopy)
Call BitBlt(Picture3.hDC, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture4.hDC, 0, 0, vbSrcCopy)
Picture3.Refresh
End Sub
Private Sub Picture3_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
down = False
looking = False
Picture3.Refresh
End Sub
Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
Timer1.Enabled = False
MediaPlayer1.CurrentPosition = 0
Picture3.Picture = Picture1.Picture
Picture3.Enabled = False
Picture4.Picture = Picture1.Picture
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNSnap Me, Button, X, Y
   play.Picture = play1.Picture
stopb.Picture = stopb1.Picture
pause.Picture = pause1.Picture
iexit.Picture = Iexit1.Picture
iopen.Picture = iopen1.Picture
End Sub


Private Sub timerplay_Timer()
Dim titleartist As String
Dim ex
ex = ListView1.ListItems.Count + 1
titleartist = Text2.Text
ListView1.ListItems.Add = ex & "."
ListView1.ListItems.Item(ex).SubItems(1) = titleartist
ListView1.ListItems.Item(ex).SubItems(2) = mmin & ":" & msec
ListView1.ListItems.Item(ex).SubItems(3) = cd.FileName

timerplay.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ModLVSubClass.UnAttach Me.hWnd
End Sub
