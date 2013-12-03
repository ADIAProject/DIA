Attribute VB_Name = "mHighlight"
Option Explicit

Public glHighlightColor                 As Long

'---------------------------------------------------------------------------------------
' Procedure   : HighlightActiveControl
' DateTime    : 06/06/2011 16:23
' Author      : Giorgio Brausi
' Purpose     :
' Description :
' Comments    :
' Returns     :
'---------------------------------------------------------------------------------------
Public Sub HighlightActiveControl(ByRef frm As Form, _
                                  ByVal ctrl As Control, _
                                  ByVal bState As Boolean)

Dim l                                   As Long

    l = 45

    On Error Resume Next
    If bState Then

        frm.Controls.Add "VB.Shape", "ShapeHL", ctrl.Container

        With frm!ShapeHL
            .BackStyle = 1
            .BorderColor = glHighlightColor
            '.BorderColor = 32896
            '.BackColor = 32896
            .BackColor = glHighlightColor
            .Move ctrl.Left - l, ctrl.Top - l, ctrl.Width + l * 2, ctrl.Height + l * 2
            .ZOrder
            .Visible = True

        End With

    Else
        frm.Controls.Remove "ShapeHL"

    End If

End Sub
