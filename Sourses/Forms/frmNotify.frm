VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin prjDIADBS.LabelW lblNameProg 
      Height          =   1065
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Label2"
   End
   Begin prjDIADBS.LabelW Label1 
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Label1"
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private msHangDuration                  As Long
Private msShowDuration                  As Long
Private msHideDuration                  As Long
Private twipsx                          As Long
Private twipsy                          As Long

Private Const notify_mode_show          As Integer = 1
Private Const notify_mode_wait          As Integer = 2
Private Const notify_mode_hide          As Integer = 3

Private notify_mode                     As Long

Private Const SW_SHOWNA                 As Integer = 8
Private strFormName                     As String

Private Sub FontCharsetChange()
' Выставляем шрифт
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

End Sub

Private Sub DrawGradientBackground(Colour1 As Long, Colour2 As Long)

Dim vert(0 To 1)                        As TRIVERTEX
Dim grc                                 As GRADIENT_RECT

    'gradient start colour
    With vert(0)
        .X = 0
        .Y = 0
        .Red = LongToSignedShort(CLng((Colour1 And &HFF&) * 256))
        .Green = LongToSignedShort(CLng(((Colour1 And &HFF00&) \ &H100&) * 256))
        .Blue = LongToSignedShort(CLng(((Colour1 And &HFF0000) \ &H10000) * 256))
        .Alpha = 0

    End With

    'gradient end colour
    With vert(1)
        .X = Me.ScaleWidth \ twipsx
        .Y = Me.ScaleHeight \ twipsx
        .Red = LongToSignedShort(CLng((Colour2 And &HFF&) * 256))
        .Green = LongToSignedShort(CLng(((Colour2 And &HFF00&) \ &H100&) * 256))
        .Blue = LongToSignedShort(CLng(((Colour2 And &HFF0000) \ &H10000) * 256))
        .Alpha = 0

    End With

    grc.UpperLeft = 0
    grc.LowerRight = 1
    GradientFill hDC, vert(0), 2, grc, 1, GRADIENT_FILL_RECT_V

End Sub

Private Sub DrawIconPicture(img As StdPicture, _
                            ImageX As Long, _
                            ImageY As Long, _
                            ImgTransColour As Long)

Dim hbmDc                               As Long
Dim hBmp                                As Long
Dim hBmpOld                             As Long
Dim Bmp                                 As BITMAP

    'if the picture is a bitmap...
    If img.Type = vbPicTypeBitmap Then
        hBmp = img.Handle
        'create a memory device context
        hbmDc = CreateCompatibleDC(0&)

        If hbmDc <> 0 Then
            'select the bitmap into the context
            hBmpOld = SelectObject(hbmDc, hBmp)

            'retrieve information for the
            'specified graphics object
            If GetObject(hBmp, Len(Bmp), Bmp) <> 0 Then
                'draw the bitmap with the
                'specified transparency colour
                TransparentBlt Me.hDC, ImageX, ImageY, Bmp.BMWidth, Bmp.BMHeight, hbmDc, 0, 0, Bmp.BMWidth, Bmp.BMHeight, ImgTransColour

            End If

            'GetObject
            SelectObject hbmDc, hBmpOld
            DeleteObject hBmpOld
            DeleteDC hbmDc

        End If

        'hbmDc
    ElseIf img.Type = vbPicTypeIcon Then
        'if the picture is an icon
        Me.PaintPicture img, ImageX, ImageY

    End If

End Sub

Private Sub Form_Click()
    Timer1.Enabled = False
    ReleaseCapture
    Unload Me

End Sub

Private Sub Form_Initialize()
'position the elements and
'set some initial settings
    twipsx = Screen.TwipsPerPixelX
    twipsy = Screen.TwipsPerPixelY
    Me.KeyPreview = True
    Me.AutoRedraw = True

    With Label1
        .Move 4 * twipsx, 40 * twipsx, Me.ScaleWidth - (7 * twipsx), Me.ScaleHeight - (44 * twipsx)
        .AutoSize = False
        .WordWrap = True
        '.BackStyle = vbTransparent
        .Alignment = vbCenter

    End With

    'LABEL1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me

    End If

End Sub

Private Sub Form_Load()
    SetupVisualStyles Me

    strFormName = Me.Name

    lblNameProg.Caption = strFrmMainCaptionTemp & vbNewLine & " v." & strProductVersion & " " & strFrmMainCaptionTempDate & strDateProgram & ")"
    ' Выставляем шрифт
    FontCharsetChange

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'trap the mouse movements while
'in the form
    If GetCapture() = Me.hWnd Then
        If X < 0 Or X > Me.Width Or Y < 0 Or Y > Me.Height Then
            ReleaseCapture
            Label1.ForeColor = &H80000012
            Label1.Font.Underline = False

        End If

    Else
        Label1.ForeColor = RGB(0, 0, 255)
        Label1.Font.Underline = True
        SetCapture Me.hWnd

    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    ReleaseCapture
    CheckUpd False
    Unload Me

End Sub

'Private Sub Form_Terminate()
'    On Error Resume Next
'
'    If Forms.Count = 0 Then
'        UnloadApp
'
'    End If
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNotify = Nothing

End Sub

Private Sub Label1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Form_MouseMove Button, Shift, X, Y

End Sub

Private Function LongToSignedShort(dwUnsigned As Long) As Integer

'convert from long to signed short
    If dwUnsigned < 32768 Then
        LongToSignedShort = CInt(dwUnsigned)
    Else
        LongToSignedShort = CInt(dwUnsigned - &H10000)

    End If

End Function

Public Sub ShowMessage(ByVal sMsg As String, _
                       Optional img As StdPicture, _
                       Optional ImageX As Long = 0, _
                       Optional ImageY As Long = 0, _
                       Optional BgColour1 As Long = &HFFFFFF, _
                       Optional BgColour2 As Long = &HFFFFFF, _
                       Optional ImgTransColour As Long = &HFFFFFF, _
                       Optional ByVal msShowTime As Long = 50, _
                       Optional ByVal msHangTime As Long = 4000, _
                       Optional ByVal msHideTime As Long = 50, _
                       Optional ByVal bPlacement As Boolean = False, _
                       Optional sSound As String)

Dim RC                                  As RECT

    'ensure the notification window
    'is not already visible
    If Not Me.Visible Then
        'clear form
        Me.Cls
        'draw gradient background
        DrawGradientBackground BgColour1, BgColour2

        'draw picture
        If Not img Is Nothing Then
            DrawIconPicture img, ImageX, ImageY, ImgTransColour

        End If

        'set the sMsg
        Label1.Caption = sMsg
        'assign the intervals for the
        'respective timer events
        msShowDuration = msShowTime
        msHangDuration = msHangTime
        msHideDuration = msHideTime

        'ready to go, so first play
        'the notification sound
        If LenB(sSound) > 0 Then
            PlaySound sSound, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If

        'retrieve the work area (the
        'available real estate available)
        SystemParametersInfo SPI_GETWORKAREA, 0, RC, 0

        'move the form in the upper-right corner
        'of the work area and set the form as
        '"topmost" (always on top). We pass
        'SWP_NOACTIVATE so the form does not
        'take focus from the active app. The
        'initial height of the form is 0
        Select Case bPlacement

            Case True
                'show top left
                SetWindowPos Me.hWnd, HWND_TOPMOST, 0, RC.Top, (Me.Width / twipsx), 0, SWP_NOACTIVATE

            Case False
                'show top right
                SetWindowPos Me.hWnd, HWND_TOPMOST, RC.Right - (Me.Width / twipsx), RC.Top, (Me.Width / twipsx), 0, SWP_NOACTIVATE

        End Select

        'show the form without activating
        ShowWindow Me.hWnd, SW_SHOWNA
        'begin the animation by setting
        'the notify mode to notify_mode_show,
        'and setting the interval to the value
        'passed as msShowDuration, and starting
        'the timer
        notify_mode = notify_mode_show
        Timer1.Interval = msShowDuration
        Timer1.Enabled = True

    End If

End Sub

Private Sub Timer1_Timer()

    Select Case notify_mode

        Case notify_mode_show

            If Me.Height + 4 * twipsx < 1900 Then
                Me.Height = Me.Height + 4 * twipsx
            Else
                Me.Height = 1900
                Timer1.Enabled = False
                Timer1.Interval = msHangDuration
                notify_mode = notify_mode_wait
                Timer1.Enabled = True

            End If

        Case notify_mode_wait
            Timer1.Enabled = False
            Timer1.Interval = msHideDuration
            notify_mode = notify_mode_hide
            Timer1.Enabled = True

        Case notify_mode_hide

            If (Me.Height - (Me.Height - Me.ScaleHeight * twipsx)) - 4 * twipsx > (Me.Height - Me.ScaleHeight * twipsx) Then
                Me.Height = Me.Height - 4 * twipsx
            Else
                Me.Height = 0
                notify_mode = 0
                Timer1.Enabled = False
                Unload Me

            End If

    End Select

End Sub
