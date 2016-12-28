Attribute VB_Name = "mHighlight"
'---------------------------------------------------------------------------------------
' Module      : HighlightActiveControl
' DateTime    : 06/06/2011 16:23
' Author      : Giorgio Brausi
'---------------------------------------------------------------------------------------

Option Explicit

Public glHighlightColor As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub HighlightActiveControl
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   frm (Form)
'                              ctrl (Control)
'                              bState (Boolean)
'!--------------------------------------------------------------------------------
Public Sub HighlightActiveControl(ByRef frm As Form, ByVal ctrl As Control, ByVal bState As Boolean)

    Dim l As Long

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
