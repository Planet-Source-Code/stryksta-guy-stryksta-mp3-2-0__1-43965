Attribute VB_Name = "DragNSnap_Mod"
'DragNSnap By Sean Siegel (SeanMSiegel@hotmail.com)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Dim CurXY As POINTAPI
Dim Cx As Long
Dim Cy As Long
Dim Tpx As Long
Dim Tpy As Long
Public Function Get_Mouse_X() As Long
    resp = GetCursorPos(CurXY) 'load mouse xy into user type curxy
    Get_Mouse_X = CurXY.X 'return the x
End Function
Public Function Get_Mouse_Y() As Long
    resp = GetCursorPos(CurXY) 'load mouse xy into user type curxy
    Get_Mouse_Y = CurXY.Y 'return the y
End Function
Sub DragNSnap(TheForm As Form, TheButton As Integer, X As Single, Y As Single, Optional ScreenSnapping As Boolean = True, Optional SnapPixels As Byte = 10)
    If Tpx = 0 Then Tpx = Screen.TwipsPerPixelX 'i use if so we only load the variable once to save process
    If Tpy = 0 Then Tpy = Screen.TwipsPerPixelY
    If TheButton <> 1 Then 'make sure they are left clicking to drag
        Cx = X 'save the mouse x so we can calculte the left later on if a button is pressed
        Cy = Y 'save the mouse y so we can calculte the top later on if a button is pressed
    Else
        tx = Get_Mouse_X * Tpx - Cx 'get the mousex in pixels and subtract its x to set the left of the form
        If tx / Tpx < -SnapPixels Then GoTo cnty
        If ScreenSnapping And tx / Tpx < SnapPixels Then tx = 0
        If (tx + TheForm.width) / Tpx > (Screen.width / Tpx) + SnapPixels Then GoTo cnty
        If ScreenSnapping And (tx + TheForm.width) / Tpx > (Screen.width / Tpx) - SnapPixels Then tx = Screen.width - TheForm.width
cnty:
        ty = Get_Mouse_Y * Tpy - Cy 'get the mousey in pixels and subtract its y to set the top of the form
        If ty / Tpy < -SnapPixels Then GoTo cnt:
        If ScreenSnapping And ty / Tpx < SnapPixels Then ty = 0
        If (ty + TheForm.height) / Tpy > (Screen.height / Tpy) + SnapPixels Then GoTo cnt:
        If ScreenSnapping And (ty + TheForm.height) / Tpx > (Screen.height / Tpx) - SnapPixels Then ty = Screen.height - TheForm.height
cnt:
        TheForm.Move tx, ty 'move the form to its new location
    End If
End Sub
