VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   2055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Set Priority"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
End Sub
