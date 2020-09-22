VERSION 5.00
Begin VB.UserControl ThreeDLine 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   ScaleHeight     =   315
   ScaleWidth      =   6225
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5280
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "ThreeDLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum lOrientation
    Horizontal
    Vertical
End Enum

Private m_Orientation As lOrientation

Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_Orientation = PropBag.ReadProperty("Orientation", 0)
End Sub
Public Sub UserControl_Resize()

'3D Line OCX with Source Code
'----------------------------
'Written by:    Tom Garratt
'Version:       1.0.0
'VB Version:    VB 6.0 Enterprise
'Email:         tom@garrattsoftware.co.uk or t.garratt@lineone.net

'Thanks for downloading the 3d line source. I hope you find this useful. The project is ready and
'compiled for you to use with your VB apps if you don't like including pages and pages of code.

'You can freely distribute this code but you must include this file in the zip.
'If you are going to use this in your app, please include me in the credits etc...

'If you want you can visit my site at: www.garrattsoftware.co.uk. I am rebuilding it at the moment to make it
'more VB oriented.At presesnt I built it to see if I could make some money. It didn't work.

If Me.Orientation = 0 Then

    UserControl.Height = 40

    Line1.X1 = 0
    Line1.X2 = UserControl.Width - 10
    Line1.Y1 = 0
    Line1.Y2 = 0
    
    Line2.Y1 = 20
    Line2.Y2 = 20
    Line2.X1 = 20
    Line2.X2 = UserControl.Width
    
End If

If Me.Orientation = 1 Then

    UserControl.Width = 40

    Line1.X1 = 0
    Line1.X2 = 0
    Line1.Y1 = 0
    Line1.Y2 = UserControl.Height - 10
    
    Line2.Y1 = 20
    Line2.Y2 = UserControl.Height
    Line2.X1 = 20
    Line2.X2 = 20
    
End If

End Sub
Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Orientation", Me.Orientation, 0
End Sub



Public Property Get Orientation() As lOrientation
Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal vNewValue As lOrientation)
m_Orientation = vNewValue
UserControl_Resize
PropertyChanged "Orientation"
If vNewValue = Vertical Then
    UserControl.Height = 2000
Else
    UserControl.Width = 2000
End If

End Property
