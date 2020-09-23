VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_SkinnedForm 
   BorderStyle     =   0  'None
   Caption         =   "Vb Skinner Ultra !"
   ClientHeight    =   8775
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin Skinned_Form.Skin Skin1 
      Align           =   1  'Align Top
      Height          =   8745
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   15425
      Caption         =   "Title Bar"
      Skin            =   "Frm_SkinnedForm.frx":0000
      RememberSkin    =   -1  'True
      Begin VB.CommandButton Command1 
         Caption         =   "Change skin During Run time"
         Height          =   525
         Left            =   3750
         TabIndex        =   1
         Top             =   2460
         Width           =   2295
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   2280
         Top             =   4230
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Can Hold other controls !"
         Height          =   765
         Left            =   960
         TabIndex        =   2
         Top             =   2490
         Width           =   2715
      End
   End
End
Attribute VB_Name = "Frm_SkinnedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd.Filter = "*.bmp"
cd.InitDir = App.Path & "\skins"
cd.ShowOpen
Skin1.ChangeSkin cd.filename
End Sub

Private Sub Form_Load()
Skin1.Init_Skin Me
End Sub


Private Sub Skin1_MenuClick(ItemName As String)
    If ItemName = "Vote" Then
        MsgBox "Please vote for me"
    Else
        MsgBox ItemName & " was Clicked"
    End If
End Sub
