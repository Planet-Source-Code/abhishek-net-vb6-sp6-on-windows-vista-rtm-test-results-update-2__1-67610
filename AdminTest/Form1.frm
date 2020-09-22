VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 On Vista RTM Test"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Shell Instance"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IsAdmin:"
      Height          =   225
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   225
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function IsUserAnAdmin Lib "shell32" () As Long

Private Sub Form_Load()

   Select Case IsUserAnAdmin()
      Case 1:
         Label1.Caption = "Yes"
         
      Case False:
         Label1.Caption = "No"
   End Select
   
End Sub

Private Sub Command1_Click()
    Call Shell(App.Path & "\isadmin.exe", vbNormalFocus)
End Sub

