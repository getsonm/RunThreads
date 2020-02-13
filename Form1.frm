VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   4785
   ClientTop       =   4440
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5010
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   690
      Left            =   810
      TabIndex        =   1
      Top             =   1935
      Width           =   2715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   690
      Left            =   810
      TabIndex        =   0
      Top             =   990
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements clsCallerObj

Private Sub clsCallerObj_Done(sIndexThreadServerLivre As Long)

    MsgBox "Fim da Thread " & sIndexThreadServerLivre, vbInformation

End Sub

Private Sub Command1_Click()

    Dim Ret As Integer

    Ret = CriaThread(1, 1, "1", Me)

    MsgBox Ret

End Sub

Private Sub Command2_Click()

    Dim Ret As Integer

    Ret = FinalizarThread(1)
    
    MsgBox Ret

End Sub
