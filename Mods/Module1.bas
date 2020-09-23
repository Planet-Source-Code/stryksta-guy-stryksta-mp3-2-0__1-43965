Attribute VB_Name = "Module1"
Public mhour As String, mmin As String, msec As String
Public chour As String, cmin As String, csec As String
Public shour As String, smin As String, ssec As String
Public check As String
Global PlaybackSpeed As Integer
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Type CtlAdj
   AdjX As Long
   AdjY As Long
End Type
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const HTBOTTOMRIGHT = 17
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Sub sectomin(lsecs As Long)
'Den här snutten översätter sekunder till "tt:mm:ss"
smin = Format(Fix(lsecs / 60), "#0")
ssec = Format(lsecs Mod 60, "00")
shour = Format(Fix(smin / 60), "#0")
smin = Format(smin Mod 60, "00")
End Sub

Public Sub ResizeOMatic(frm As Form, adj() As CtlAdj)
    Dim tmpControl As Control
    Dim index As Long
    On Error Resume Next
    index = 0


    For Each tmpControl In frm
        index = index + 1


        Select Case LCase$(tmpControl.Tag)
            Case "rx" 'relative X
            tmpControl.Left = frm.width - tmpControl.width - adj(index).AdjX
            Case "ry" 'relative Y
            tmpControl.Top = frm.height - tmpControl.height - adj(index).AdjY
            
            Case "rxy" 'relative XY
            tmpControl.Left = frm.width - tmpControl.width - adj(index).AdjX
            tmpControl.Top = frm.height - tmpControl.height - adj(index).AdjY
            Case "sx" 'stretch X
            tmpControl.width = frm.width - tmpControl.Left - adj(index).AdjX
            Case "sy" 'stretch Y
            tmpControl.height = frm.height - tmpControl.Top - adj(index).AdjY
            Case "sxy" 'stretch XY
            tmpControl.width = frm.width - tmpControl.Left - adj(index).AdjX
            tmpControl.height = frm.height - tmpControl.Top - adj(index).AdjY
            Case "sxry" 'stretch X relative to Y
            tmpControl.width = frm.width - tmpControl.Left - adj(index).AdjX
            tmpControl.Top = frm.height - tmpControl.height - adj(index).AdjY
            Case "syrx" 'stretch Y, relative x
            tmpControl.height = frm.height - tmpControl.Top - adj(index).AdjY
            tmpControl.Left = frm.width - tmpControl.width - adj(index).AdjX
        Case "sxrx" 'stretch Y, relative x
            tmpControl.width = frm.width - tmpControl.Left - adj(index).AdjX
            tmpControl.Left = frm.width - tmpControl.width - adj(index).AdjX
        End Select
Next
End Sub


Public Sub RegisterForm(frm As Form, width As Long, height As Long, adj() As CtlAdj)

    Dim tmpControl As Control
    ReDim adj(0)
    On Error Resume Next


    For Each tmpControl In frm
        ReDim Preserve adj(UBound(adj) + 1)
        adj(UBound(adj)).AdjX = width - (tmpControl.Left + tmpControl.width)
        adj(UBound(adj)).AdjY = height - (tmpControl.Top + tmpControl.height)
    Next
    
End Sub
Public Sub LVFlatColumnHeaders(LV As ListView)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong hHeader, GWL_STYLE, InitLVStyle Xor HDS_BUTTONS
End Sub
