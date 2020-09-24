VERSION 5.00
Object = "{04039C52-2BE7-45B9-A9A2-94A3EA714C22}#2.0#0"; "eamContainer.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin Container.eamContainer eamContainer1 
      Align           =   3  'Align Left
      Height          =   8910
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   15716
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Container Header"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub eamContainer1_Resize()
Command1.Move 0, 0
End Sub


