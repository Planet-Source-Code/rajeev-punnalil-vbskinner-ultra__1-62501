VERSION 5.00
Begin VB.UserControl Skin 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   ControlContainer=   -1  'True
   DataSourceBehavior=   1  'vbDataSource
   EditAtDesignTime=   -1  'True
   HasDC           =   0   'False
   PropertyPages   =   "Skin.ctx":0000
   ScaleHeight     =   2730
   ScaleWidth      =   8850
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Bar"
         Height          =   195
         Left            =   570
         TabIndex        =   1
         Top             =   90
         Width           =   585
      End
   End
   Begin VB.PictureBox MenuBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   8760
      TabIndex        =   3
      Top             =   315
      Width           =   8790
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.PictureBox Skinpic 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   2100
      Picture         =   "Skin.ctx":000F
      ScaleHeight     =   795
      ScaleWidth      =   3135
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1410
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   0
      Begin VB.Menu a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   1
      Begin VB.Menu b 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   2
      Begin VB.Menu c 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   3
      Begin VB.Menu d 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   4
      Begin VB.Menu e 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   5
      Begin VB.Menu f 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   6
      Begin VB.Menu g 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   7
      Begin VB.Menu h 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "Skin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'----------------------------------------------------------------------------------------------------------
'
'
'           Author   : Rajeev P
'           Email ID : Rajeev_Punnalil@hotmail.com
'
'           All of u might have used vbskinner which is not free
'           for pro version .Here i have included almost all skinner features of vbskinner pro.
'           If u guys have any suggestions please contact me at Rajeev_punnalil@hotmail.com. U may make
'           and may redistribute this code as long as this commented lines
'           are retainded in all of them.
'
'----------------------------------------------------------------------------------------------------------
'           Note : Retain The above lines in all redistributed versions
'
'
'           This code uses skins from vbskinner so u can go there and download
'           more skin files if u want . Enjoy!
'
'           IMPORTANT !
'           ------------
'           Remember to change form borderstyle to 0-none
'           Use 'send to back' on the skinner activex
'
'
'           Improvements over last version
'           -------------------------------
'           Thanks a lot for ur wonderful response which made me look back to my code
'           and i found some junk stuff in there . so i got rid of some and added some nice ones.
'           1) It doesnot use iterative method any more and hence it is really fast now
'               ,Thanks to the guy who suggested c++ which made me think about optimising the code
'           2) Form resize has been added ,Thanks , that was a nice suggestion
'           3) Rounded Form added
'           4)Load from Saved setting made optional
'           4)Skinning is initiated even at design time ! nice isnt it !
'           5)The control can now contain other controls so that u can do all the design work on the control
'           6)Finally i have managed to pulloff the menus as well ! so thats good going.. thanks a lot for
'             all your comments which helped me to make this skinner this wonderful, the menu support is not there
'             for the actual vbskinner as well, so he/she can take some thing away from my project !!! payback time u know !
'           Thanks
'           ------
'               Thanks for all ur suggestions . A special thanks to merlin,he got
'           the idea of the project correct ! plz add suggestions !
'----------------------------------------------------------------------------------------------------------
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


'Api Functions of picture manipulations and windows rgn settings


'Structure for saving location of buttons
Private Type ButtonLocation
    Close_Left As Integer
    Close_Right As Integer
    Min_Left As Integer
    Min_Right As Integer
    Max_Left As Integer
    Max_Right As Integer
End Type

Private skinned As Boolean
Private skinfile As String
Private Bool_Min As Boolean
Private Bool_Max As Boolean
Private Bool_Remember  As Boolean

Private frm As Form

Private initok As Boolean 'init check variable
Private locate As ButtonLocation

'Public Evenets
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MenuClick(ItemName As String)
Private inrgn As Boolean
Private MenuNameArray() As String
Private MenuCaptionArray() As String
Private rounded As Boolean
Private a1() As String, b1() As String, c1() As String, d1() As String, e1() As String, f1() As String, g1() As String, h1() As String


Private Sub MakeRoundedEdge() 'Creates a Ronded Rect region if the skin supports it !
   Skinpic.ScaleMode = 3
   If GetPixel(Skinpic.hdc, 0, 14) = 16711935 Then
        Skinpic.ScaleMode = 1
        If initok = True Then
            Call SetWindowRgn(frm.hwnd, CreateRoundRectRgn(2, 0, UserControl.Width, UserControl.Height, 14, 14), True)
            rounded = True
        End If

   Else
        Skinpic.ScaleMode = 1
   End If
    
End Sub
Private Sub Init_PaintSkin() ' Runs Routines which paints the skin
If initok = True Then
    UserControl.Cls
    DrawTitleBar
    Draw_Close_Defaults
    Draw_Min_Defaults
    Draw_Max_Defaults frm.WindowState
    Draw_BackGround
    MenuBar.BackColor = UserControl.BackColor
End If
End Sub
Private Sub Init_PaintDesigner() 'Runs skinner in designer mode
    UserControl.Cls
    TitleBar.Cls
    DrawTitleBar
    Draw_Close_Defaults
    Draw_Min_Defaults
    Draw_Max_Defaults 0
    Draw_DesignerBackGround
    MenuBar.BackColor = UserControl.BackColor
End Sub

Public Sub ChangeSkin(filename As String) ' For changing skin during run time
    If initok = False Then Exit Sub
    SaveSetting frm.Caption, "skin", "main", filename
    On Error Resume Next
    If Not filename = "" Then
        Skinpic.Picture = LoadPicture(filename)
    End If
    Init_PaintSkin
End Sub
Private Sub InitAllocate() 'avoiding some redundant code from allocating effort to increase speed
    frm.BorderStyle = 0
    TitleBar.Top = 0
    TitleBar.Left = 0
    CapLabel.Top = 75
    CapLabel.Caption = frm.Caption
    UserControl.Height = frm.Height
End Sub

Private Sub AllocateDesigner() 'Allocates in design mode
    Dim level As Integer
    
    TitleBar.Width = UserControl.Width
    TitleBar.Height = 300
    MenuBar.Top = TitleBar.Height + 100
    MenuBar.Width = UserControl.Width - 170
    MenuBar.Left = 90
    
    locate.Close_Left = TitleBar.Width - 300
    locate.Close_Right = TitleBar.Width - 105
    
    MakeRoundedEdge
    
    If Me.ButtonMax Then  'allocate Max button if present
        locate.Max_Left = TitleBar.Width - 510
        locate.Max_Right = TitleBar.Width - 330
        level = level + 1
    End If
    
    If Me.ButtonMin Then  'allocate min button if present depend on weather maxbutton is present or not
        If level = 1 Then
            locate.Min_Left = TitleBar.Width - 720
            locate.Min_Right = TitleBar.Width - 480
        Else
            locate.Min_Left = TitleBar.Width - 510
            locate.Min_Right = TitleBar.Width - 330
        End If
    End If
End Sub
Public Sub allocate() ' allocate locations of various buttons depending on settings
    UserControl.Width = frm.Width
    UserControl.Height = frm.Height
    initok = True
End Sub

Public Function GetSkinTheme() As Long ' Returns usercontrol's backcolor
    GetSkinTheme = UserControl.BackColor
End Function
Private Sub Draw_BackGround()
    'This Part has been drastically improved from last version
    'iteration techniques are avoided producing faster o/p
If initok = True Then
   Skinpic.ScaleMode = 3
    UserControl.BackColor = GetPixel(Skinpic.hdc, 164, 22)
    Skinpic.ScaleMode = 1
    UserControl.PaintPicture Skinpic.Picture, 0, 0, 75, frm.ScaleHeight, 2370, 240, 75, 240
    UserControl.PaintPicture Skinpic.Picture, frm.ScaleWidth - 75, 0, 75, frm.ScaleHeight, 3000, 225, 135, 255
    UserControl.PaintPicture Skinpic.Picture, 0, TitleBar.Height, frm.ScaleWidth, 75, 2595, 0, 270, 75
    UserControl.PaintPicture Skinpic.Picture, 0, frm.ScaleHeight - 75, frm.ScaleWidth, 75, 2595, 645, 270, 75
    
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub
Private Sub Draw_DesignerBackGround() 'Draws background in designer mode
    UserControl.ScaleMode = 3
    UserControl.BackColor = GetPixel(Skinpic.hdc, 164, 22)
    UserControl.ScaleMode = 1
    UserControl.PaintPicture Skinpic.Picture, 10, 0, 75, UserControl.ScaleHeight, 2370, 240, 75, 240
    UserControl.PaintPicture Skinpic.Picture, UserControl.ScaleWidth - 75, 0, 75, UserControl.ScaleHeight, 3000, 225, 135, 255
    UserControl.PaintPicture Skinpic.Picture, 0, TitleBar.Height, UserControl.ScaleWidth, 75, 2595, 0, 270, 75
    UserControl.PaintPicture Skinpic.Picture, 0, UserControl.ScaleHeight - 75, UserControl.ScaleWidth, 75, 2595, 645, 270, 75
End Sub
Private Sub DrawTitleBar() 'Draws Title Bar, iteration avoided !

 TitleBar.PaintPicture Skinpic.Picture, 0, 0, TitleBar.ScaleWidth, 375, 300, 210, 270, 435
 If rounded = True Then TitleBar.PaintPicture Skinpic.Picture, 0, 0, 235, 375, 0, 210, 270, 435


End Sub
Private Sub a_Click(index As Integer)
    RaiseEvent MenuClick(a1(index))
End Sub
Private Sub b_Click(index As Integer)
    RaiseEvent MenuClick(b1(index))
End Sub
Private Sub c_Click(index As Integer)
    RaiseEvent MenuClick(c1(index))
End Sub
Private Sub d_Click(index As Integer)
    RaiseEvent MenuClick(d1(index))
End Sub
Private Sub e_Click(index As Integer)
    RaiseEvent MenuClick(e1(index))
End Sub
Private Sub f_Click(index As Integer)
    RaiseEvent MenuClick(f1(index))
End Sub
Private Sub g_Click(index As Integer)
    RaiseEvent MenuClick(g1(index))
End Sub
Private Sub h_Click(index As Integer)
    RaiseEvent MenuClick(h1(index))
End Sub

Private Sub CapLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Helps in moving the form
If initok = True Then
    TitleBar_MouseMove Button, 0, X, Y
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub






Private Sub Label1_Click()

End Sub

Private Sub Label_Click(index As Integer)
    PopupMenu Core(index), , Label(index).Left, MenuBar.Top + MenuBar.Height
End Sub

Private Sub TitleBar_DblClick()
If Bool_Max = True Then
    If frm.WindowState = 2 Then
        frm.WindowState = 0
        
    Else
        frm.WindowState = 2
    End If
End If
Draw_BackGround
End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' Determines titlebar mousedown
If initok = True Then
    If X > locate.Close_Left And X < locate.Close_Right Then
        Unload frm
    Else
    If X > locate.Min_Left And X < locate.Min_Right Then
        frm.WindowState = 1
        Draw_Min_Defaults
    Else
    If X > locate.Max_Left And X < locate.Max_Right Then
        If frm.WindowState = 0 Then
            UserControl.Cls
            TitleBar.Cls
            frm.WindowState = 2
            allocate
            Init_PaintSkin
        Else
            TitleBar.Cls
            UserControl.Cls
            frm.WindowState = 0
            allocate
            Init_PaintSkin
        End If
    Draw_Max_Defaults frm.WindowState
    End If
    End If
    End If
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub

Private Sub ResetBar() 'Resetes to default
If initok = True Then
    Draw_Close_Defaults
    Draw_Min_Defaults
    Draw_Max_Defaults frm.WindowState
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub


Private Sub Draw_Close_Defaults() 'Draws close default
If locate.Close_Right = 0 Then Exit Sub
    TitleBar.PaintPicture Skinpic.Picture, locate.Close_Left, 90, 195, 195, 0, 0, 195, 195
End Sub
Private Sub Draw_Close_Move() 'Draws close mouse over
If locate.Close_Right = 0 Then Exit Sub
    TitleBar.PaintPicture Skinpic.Picture, locate.Close_Left, 90, 195, 195, 210, 0, 195, 195
End Sub
Private Sub Draw_Close_Down() 'Draws Close Mouse Down
If locate.Close_Right = 0 Then Exit Sub

    TitleBar.PaintPicture Skinpic.Picture, locate.Close_Left, 90, 195, 195, 420, 0, 195, 195

End Sub

Private Sub Draw_Min_Defaults() 'Draws Min Default
If locate.Min_Right = 0 Then Exit Sub
TitleBar.PaintPicture Skinpic.Picture, locate.Min_Left, 90, 195, 195, 1890, 0, 195, 195

End Sub
Private Sub Draw_Min_Move() 'Draws min mouse move
If locate.Min_Right = 0 Then Exit Sub

    TitleBar.PaintPicture Skinpic.Picture, locate.Min_Left, 90, 195, 195, 2100, 0, 195, 195
End Sub
Private Sub Draw_Max_Defaults(ByVal State As Integer) 'Draws max defaults
If locate.Max_Right = 0 Then Exit Sub
If State = 2 Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 630, 0, 195, 195
    UserControl.Height = frm.Height
    UserControl.Width = frm.Width
    
Else
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 1260, 0, 195, 195
End If
End Sub
Private Sub Draw_Max_Move(ByVal State As Integer) 'Draws max mouse move
If locate.Max_Right = 0 Then Exit Sub
If State = 2 Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 840, 0, 195, 195
Else
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 1470, 0, 195, 195
End If
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) ' Title Bar Mouse move
If initok = True Then
If X > locate.Close_Left And X < locate.Close_Right Then
    Draw_Close_Move
    Draw_Min_Defaults
    Draw_Max_Defaults frm.WindowState
Else
    Draw_Close_Defaults
If X > locate.Min_Left And X < locate.Min_Right Then
    Draw_Min_Move
    Draw_Max_Defaults frm.WindowState
    Draw_Close_Defaults
Else
If X > locate.Max_Left And X < locate.Max_Right Then
    Draw_Max_Move frm.WindowState
    Draw_Min_Defaults
    Draw_Close_Defaults
Else
If Not frm.WindowState = 2 Then
    ReleaseCapture
    SendMessage frm.hwnd, &HA1, 2, 0
End If
End If
End If
End If
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
        
End If
End Sub



Private Sub UserControl_Initialize()
    TitleBar.Width = UserControl.Width
    LoadMenu
End Sub

' Various propertis of skinner
Public Property Let MenuVisible(Temp As Boolean)
    MenuBar.Visible = Temp
    PropertyChanged "MenuVisible"
End Property

Public Property Let RememberSkin(Temp As Boolean) '**newly added
    Bool_Remember = Temp
End Property
Public Property Let ButtonMin(Temp As Boolean)
    Bool_Min = Temp
    UserControl_Resize
End Property

Public Property Let ButtonMax(Temp As Boolean)
    Bool_Max = Temp
     UserControl_Resize
End Property
Public Property Let Caption(Temp As String)
    CapLabel = Temp
End Property
Public Property Get MenuVisible() As Boolean
    MenuVisible = MenuBar.Visible
End Property

Public Property Get ButtonMin() As Boolean
    ButtonMin = Bool_Min
End Property
Public Property Get RememberSkin() As Boolean
    RememberSkin = Bool_Remember
End Property
Public Property Get ButtonMax() As Boolean
Attribute ButtonMax.VB_ProcData.VB_Invoke_Property = "MenuItem"
    ButtonMax = Bool_Max
End Property

Public Property Get Caption() As String
    Caption = CapLabel.Caption
End Property
Public Property Set Skin(Skn As Picture)
    Set Skinpic.Picture = Skn
    If initok Then Init_PaintSkin
End Property
Public Property Get Skin() As Picture
       Set Skin = Skinpic.Picture
End Property


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
'Form resize added in routine usercontrol - mouseup,mousemove!
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If X > UserControl.Width - 200 Or Y > UserControl.Height - 200 Then
            inrgn = True
        Else
            If Not inrgn And Not frm.WindowState = 2 Then
                ReleaseCapture
                SendMessage frm.hwnd, &HA1, 2, 0
            End If
        End If
    Else
        If X > UserControl.Width - 200 And Y > UserControl.Height - 200 Then
        UserControl.MousePointer = 8
    Else
    If Y > 400 Then
        If X > UserControl.Width - 200 And Y < UserControl.Height - 200 Then
            UserControl.MousePointer = 9
        Else
            If X < UserControl.Width - 200 And Y > UserControl.Height - 200 Then
                UserControl.MousePointer = 7
            Else
                UserControl.MousePointer = 0
            End If
        End If
    Else
        UserControl.MousePointer = 0
    End If
    End If
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
    ResetBar
End Sub



Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
       On Error Resume Next
       If inrgn = True Then
            If UserControl.MousePointer = 8 Or UserControl.MousePointer = 9 Then
                If X < 3750 Then X = 3750
                frm.Width = X
            End If
            If UserControl.MousePointer = 8 Or UserControl.MousePointer = 7 Then
                If Y < 3270 Then Y = 3270
                frm.Height = Y
            End If
            allocate
            
        inrgn = False
       End If
       UserControl.MousePointer = 0
End Sub

Private Sub UserControl_Resize() ' Repaints on resize event
    AllocateDesigner
    Init_PaintDesigner
End Sub

' Property Bag Additions
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
    Call .ReadProperty("Caption", "")
    Bool_Max = .ReadProperty("ButtonMax", True)
    Bool_Min = .ReadProperty("ButtonMin", True)
    Set Skinpic = .ReadProperty("Skin", Skinpic.Picture)
    Bool_Remember = .ReadProperty("RememberSkin", False)
    MenuBar.Visible = .ReadProperty("MenuVisible", True)
 End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
    Call .WriteProperty("Caption", CapLabel.Caption, "")
    Call .WriteProperty("Skin", Skinpic, Skinpic.Picture)
    Call .WriteProperty("ButtonMax", Bool_Max, True)
    Call .WriteProperty("ButtonMin", Bool_Min, True)
    Call .WriteProperty("RememberSkin", Bool_Remember, False)
    Call .WriteProperty("MenuVisible", MenuBar.Visible, True)
End With
End Sub

'Initializes with form control .. Vital part of the code

Public Sub Init_Skin(frm1 As Form)
    initok = True
    Set frm = frm1
    Dim filename As String
    filename = GetSetting(frm.Caption, "skin", "main", "")
    InitAllocate
    allocate
    If Me.RememberSkin = True And Not filename = "" Then ChangeSkin filename
    
End Sub

Private Sub LoadMenu()

If Dir(App.Path + "\menu.txt") = "" Then Exit Sub
Dim Temp As String, Temp1 As String
Dim i As Integer
Dim level As Integer
Dim ctr As Integer

Open App.Path + "\menu.txt" For Input As #1
While Not EOF(1)
    Line Input #1, Temp
    On Error Resume Next
    Temp1 = Mid(Temp, 1, InStr(Temp, "!") - 1)
    Temp = Mid(Temp, InStr(Temp, "!") + 1)
    level = GetLevel(Temp)
    If level = 0 Then
        i = 0
        If Not ctr = 0 Then
            Load Label(ctr)
            Label(ctr).Left = Label(ctr - 1).Left + Label(ctr - 1).Width + 200
            Label(ctr).Caption = Temp
            Label(ctr).Visible = True
            ctr = ctr + 1
        Else
            Label(ctr).Caption = Temp
            ctr = ctr + 1
        End If
    Else
    Temp = Trimed(Temp)
        Select Case (ctr - 1)
    Case 0:
            Load a(i)
            a(i).Caption = Temp
            ReDim Preserve a1(i)
            a1(i) = Temp1
    Case 1:
            Load b(i)
            b(i).Caption = Temp
            ReDim Preserve b1(i)
            b1(i) = Temp1
    Case 2:
            Load c(i)
            c(i).Caption = Temp
            ReDim Preserve c1(i)
            c1(i) = Temp1
    Case 3:
            Load d(i)
            d(i).Caption = Temp
            ReDim Preserve d1(i)
            d1(i) = Temp1
    Case 4:
            Load e(i)
            e(i).Caption = Temp
            ReDim Preserve e1(i)
            e1(i) = Temp1
    Case 5:
            Load f(i)
            f(i).Caption = Temp
            ReDim Preserve f1(i)
            f1(i) = Temp1
    Case 6:
            Load g(i)
            g(i).Caption = Temp
            ReDim Preserve g1(i)
            g1(i) = Temp1
    Case 7:
            Load h(i)
            h(i).Caption = Temp
            ReDim Preserve h1(i)
            h1(i) = Temp1
    End Select
    i = i + 1
    End If
    
    
Wend
Close #1
End Sub
Private Function Trimed(Temp As String) As String
again:
    If InStr(Temp, "....") Then
        Temp = Mid(Temp, 5)
        GoTo again
    End If
    Trimed = Temp
End Function


Private Function GetLevel(ByVal Temp As String) As Integer
    Dim i As Integer
again:
    If InStr(Temp, "....") Then
        Temp = Mid(Temp, 5)
        i = i + 1
        GoTo again
    End If
    GetLevel = i

End Function

