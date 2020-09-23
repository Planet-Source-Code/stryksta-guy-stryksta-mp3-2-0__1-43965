VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   30
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   1725
   LinkTopic       =   "Form2"
   ScaleHeight     =   30
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      Begin VB.Menu mnuadd 
         Caption         =   "Add File"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprev 
         Caption         =   "Previous"
      End
      Begin VB.Menu mnuplay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnupause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnunext 
         Caption         =   "Next"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuadd_Click()
On Error GoTo MS ' Error control
    Dim vFiles As Variant
    Dim lFile As Long
    With CommonDialog1
        .Filename = "" 'Clear the filename
        .CancelError = True 'Gives an error if cancel is pressed
        .DialogTitle = "Select File(s)..."
        .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Falgs, allows Multi select, Explorer style and hide the Read only tag
        .Filter = "All Supported Formats|*.mp3;*.wav;*.wma;*.asf;*.asx;*.lsf;*.lsx;*.mid;*.midi;*.rmi;*.aif;*.aifc;*.aiff;*.au;*.snd"
        .ShowOpen
        vFiles = Split(.Filename, Chr(0)) 'Splits the filename up in segments
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
    Form1.List.AddItem .FileTitle
    Else
    For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
    Form1.List1.AddItem vFiles(0) + "\" & vFiles(lFile)
    Form1.List.AddItem vFiles(lFile)
    Form1.lblItems.Caption = List.ListCount
    Next
    End If
    End With
MS:
End Sub

Private Sub mnunext_Click()
On Error Resume Next
If Form1.List.ListCount = 0 Then
Exit Sub
Else
If Form1.List.ListIndex + 1 < Form1.List.ListCount Then
Form1.List.ListIndex = Form1.List.ListIndex + 1
Form1.am1.FileTitle = Form1.List.Text
Else
Form1.List.ListIndex = 0
Form1.am1.Filename = Form1.List1.Text
End If
End If
End Sub

Private Sub mnupause_Click()
If Form1.List.ListCount = 0 Then Exit Sub
If Form1.Labelpause.Caption = "Pause" Then
Form1.am1.pause
Form1.Labelpause.Caption = "Resume"
Else
Form1.am1.play
Form1.Labelpause.Caption = "Pause"
End If
End Sub

Private Sub mnuplay_Click()
On Error Resume Next
Form1.List1.ListIndex = Form1.List.ListIndex
ext.Text = right$(Form1.List.Text, 3)
Form1.am1.Open (Form1.List.ListIndex + 1)
Form1.am1.play
Form1.am1.Filename = List.Text
Form1.Label24.Caption = List.Text
End Sub

Private Sub mnuprev_Click()
 On Error Resume Next
If Form1.List.ListCount = 0 Then
Exit Sub
Else
If Form1.List.ListIndex - 1 > -1 Then
Form1.List.ListIndex = Form1.List.ListIndex - 1
Form1.am1.Filename = Form1.List.Text
Form1.Label24.Caption = Form1.List.Text
Else
Form1.List.ListIndex = Form1.List.ListCount - 1
Form1.am1.Filename = Form1.List.Text
Form1.Label24.Caption = Form1.List.Text
End If
End If
End Sub

Private Sub mnustop_Click()
On Error GoTo hell
Form1.am1.Stop
Form1.am1.CurrentPosition = 0
hell:
Exit Sub
End Sub
