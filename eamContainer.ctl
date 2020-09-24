VERSION 5.00
Begin VB.UserControl eamContainer 
   Alignable       =   -1  'True
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   ScaleHeight     =   5280
   ScaleWidth      =   3390
   ToolboxBitmap   =   "eamContainer.ctx":0000
   Begin VB.Shape shpMain 
      BorderColor     =   &H00A19D9D&
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Container Header"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   1275
   End
   Begin VB.Shape shpHeader 
      BackColor       =   &H00D8E9EC&
      BorderColor     =   &H00A19D9D&
      FillColor       =   &H00D8E9EC&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "eamContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Event Resize()
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let Caption(ByVal New_Caption As String)
    lblHeader.Caption() = New_Caption
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "General"
    Caption = lblHeader.Caption
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblHeader.Font = New_Font
End Property

Public Property Get Font() As Font
    Set Font = lblHeader.Font
End Property

Public Property Let HeaderBackColor(ByVal New_BackColor As OLE_COLOR)
    shpHeader.FillColor() = New_BackColor
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
    HeaderBackColor = shpHeader.FillColor
End Property

Public Property Let HeaderBorderColor(ByVal New_BorderColor As OLE_COLOR)
shpHeader.BorderColor() = New_BorderColor
End Property

Public Property Get HeaderBorderColor() As OLE_COLOR
Attribute HeaderBorderColor.VB_ProcData.VB_Invoke_Property = "General"
    HeaderBorderColor = shpHeader.BorderColor
End Property
Public Property Let HeaderForeground(ByVal New_ForegroundColor As OLE_COLOR)
    lblHeader.ForeColor() = New_ForegroundColor
End Property

Public Property Get HeaderForeground() As OLE_COLOR
    HeaderForeground = lblHeader.ForeColor
End Property
Public Property Let MainBackColor(ByVal New_BackColor As OLE_COLOR)
    shpMain.FillColor() = New_BackColor
End Property

Public Property Get MainBackColor() As OLE_COLOR
    MainBackColor = shpMain.BackColor
End Property
Public Property Let MainBorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpMain.BorderColor() = New_BorderColor
End Property

Public Property Get MainBorderColor() As OLE_COLOR
Attribute MainBorderColor.VB_ProcData.VB_Invoke_Property = "General"
    MainBorderColor = shpMain.BorderColor
End Property

Private Sub UserControl_Resize()
shpHeader.Move 50, 50, UserControl.Width - 100, shpHeader.Height
lblHeader.Move 70, 53, lblHeader.Width, lblHeader.Height
shpMain.Move 50, 50 + shpHeader.Height + 30, shpHeader.Width, UserControl.Height - 50 - 30 - shpHeader.Height - 50

RaiseEvent Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' Control Background Colour
UserControl.BackColor = PropBag.ReadProperty("BackColor", &HD8E9EC)

' Header Background Colour
shpHeader.FillColor = PropBag.ReadProperty("HeaderBackGround", &HD8E9EC)

' Header Border Colour
shpHeader.BorderColor = PropBag.ReadProperty("HeaderBorderColor", &HA19D9D)

' Caption Font
Set lblHeader.Font = PropBag.ReadProperty("Font", Ambient.Font)

' Caption
lblHeader.Caption = PropBag.ReadProperty("Caption", "Container Header")

' Caption Foreground
lblHeader.ForeColor = PropBag.ReadProperty("HeaderForeground", &H80000012)

' Main Border Color
shpMain.BorderColor = PropBag.ReadProperty("HeaderBorderColor", &HA19D9D)

' Main Background Color
shpMain.FillColor = PropBag.ReadProperty("HeaderBackColor", &H80000005)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' Control Background Colour
Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HD8E9EC)

' Header Background Colour
Call PropBag.WriteProperty("HeaderBackColor", shpHeader.FillColor, &HD8E9EC)

' Header Border Colour
Call PropBag.WriteProperty("HeaderBorderColor", shpHeader.BorderColor, &HA19D9D)

' Caption Font
Call PropBag.WriteProperty("Font", lblHeader.Font, Ambient.Font)

' Caption
Call PropBag.WriteProperty("Caption", lblHeader.Caption, "Container Caption")

' Caption Foreground
Call PropBag.WriteProperty("HeaderForeground", lblHeader.ForeColor, &H80000012)

' Main Border Color
Call PropBag.WriteProperty("MainBorderColor", shpMain.BorderColor, &HA19D9D)

' Main Background Color
Call PropBag.WriteProperty("MainBorderColor", shpMain.FillColor, &H80000005)
End Sub






