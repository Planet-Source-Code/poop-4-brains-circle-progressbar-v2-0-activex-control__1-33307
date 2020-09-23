VERSION 5.00
Object = "{C7BBFD4A-3D12-11D6-8D8D-000244057B3B}#2.0#0"; "circleprog.ocx"
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circle Progressbar OCX"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   3  'Windows Default
   Begin prjCircleProgOCX.CircleProg cir 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5953
      YRAD            =   0
      XRAD            =   0
      PSHOWPER        =   0   'False
      PVALUE          =   0
      PMAX            =   0
      CVALUE          =   0
      CNONVALUE       =   0
      CPER            =   0
      CBACK           =   0
      PDEPTH          =   0
      PMODE           =   0   'False
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   4335
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw Bar"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   4335
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************
'Circle Progressbar 2.0 OCX
'by Kevin Fleet
'a cool progressbar that is in a circle
'PLEASE VOTE AT PSC!
'********************************

Private Sub cmdAbout_Click()
cir.About 'show the about box for the control
End Sub

Private Sub cmdDraw_Click()
cir.UpdatePicture
'you must now call the function instead of it drawing it
'automaticly because the function is too slow and if it updated
'continuosly it would be extremely slow
End Sub

Private Sub Form_Load()
' * * * * NOTE!!! * * * *
'you need to set all of the values prior to updating the control's picture

'(i set them in the code so you see that you have to set them before use)
cir.Depth = 10
cir.Max = 10
cir.Value = 7
cir.ValueColor = vbGreen
cir.BackColor = vbBlack
cir.NonValueColor = vbRed
cir.Is3D = True
cir.progWidth = 150
cir.progHeight = 75
cir.ShowCaption = True
End Sub
