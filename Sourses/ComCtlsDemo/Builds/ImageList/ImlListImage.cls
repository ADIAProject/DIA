VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImlListImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowImageList As ImageList
Private PropIndex As Long, PropKey As String
Private PropTag As String
Private PropPicture As IPictureDisp

Friend Sub FInit(ByVal Object As ImageList, ByVal Key As String, ByVal Picture As IPictureDisp)
Set ShadowImageList = Object
PropKey = Key
Set PropPicture = Picture
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the list image."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = PropKey
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = PropTag
End Property

Public Property Let Tag(ByVal Value As String)
PropTag = Value
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the image picture."
Set Picture = PropPicture
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Sub Draw(ByVal hDC As Long, Optional ByVal X As Long, Optional ByVal Y As Long, Optional ByVal Style As ImlDrawConstants)
Attribute Draw.VB_Description = "Draws the image to a given device context (DC) at a specified location using a specified style."
ShadowImageList.FListImageDraw PropIndex, hDC, X, Y, Style
End Sub

Public Function ExtractIcon() As IPictureDisp
Attribute ExtractIcon.VB_Description = "Creates an icon from the list image."
Set ExtractIcon = ShadowImageList.FListImageExtractIcon(PropIndex)
End Function