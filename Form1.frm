VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Form2"
      Height          =   360
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2415
      Begin VB.Label lblGitSource 
         BackStyle       =   0  'Transparent
         Caption         =   "Git Source"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCommand1_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Dim s1 As String, s2 As String

    If started = False Then

        If MsgBox("真的需要启动吗?", vbOKCancel) = vbCancel Then
            Unload Me
        End If
        started = True
    End If
End Sub
