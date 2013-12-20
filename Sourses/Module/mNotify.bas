Attribute VB_Name = "mNotify"
Option Explicit

'File for the MSIM notify sound
Private sNotifySound As String

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetMsimNotifySound
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function GetMsimNotifySound() As String

    Dim sKey As String

    'valid values for the second-last member
    'of this string are:
    'MSMSGS_ContactOnline
    'MSMSGS_NewAlert
    'MSMSGS_NewMail
    'MSMSGS_NewMessage
    '
    'You could also use sounds listed under
    'current user \ Schemes \ apps such as"
    'HKEY_CURRENT_USER\AppEvents\Schemes\
    'Apps\.Default\MailBeep\.Current
    sKey = "AppEvents\Schemes\Apps\MSMSGS\MSMSGS_ContactOnline\.Current"
    GetMsimNotifySound = GetRegString(HKEY_CURRENT_USER, sKey, vbNullString)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ShowNotifyMessage
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   Msg (String)
'!--------------------------------------------------------------------------------
Public Sub ShowNotifyMessage(Msg As String)
    'sMsg: string to display
    'ico: Image to display in the notify window -
    '     can be icon or a bitmap
    'ImageX: X coordinate of image relative to
    '     upper left corner of the form
    'ImageY: Y coordinate of image relative to
    '     upper left corner of the form
    'Duration: specify the duration
    'BgColour1: Colour of gradient background (top)
    'BgColour2: Colour of gradient background (bottom)
    'ImgTransColour: specifies the transparency colour
    '     for bitmap image. Ignored for icons
    'msShowTime: milliseconds between reveal increments, default=50
    'msHangTime: milliseconds form remains on-screen, default=4000
    'msHideTime: milliseconds between hide increments, default=50
    'bPlacement: True for top right, false for top left
    'sSound: Path of the sound to be played
    sNotifySound = GetMsimNotifySound()
    frmNotify.ShowMessage sMsg:=Msg, img:=frmMain.Icon, ImageX:=88, ImageY:=4, BgColour1:=RGB(133, 112, 243), BgColour2:=RGB(255, 255, 255), ImgTransColour:=RGB(255, 0, 0), msShowTime:=20, msHangTime:=11000, msHideTime:=10, bPlacement:=False, _
                                sSound:=sNotifySound
    'here's the same call without
    'the inline variable names
    'Call frmNotify.ShowMessage(msg,
    'Image1.Picture,
    '88,
    '4,
    'RGB(133, 112, 243),
    'RGB(255, 255, 255),
    'RGB(255, 0, 0),
    '10,
    '4000,
    '10,
    'False,
    'sNotifySound)
End Sub
