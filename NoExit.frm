VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Hehe, no-one can exit me"
   ClientHeight    =   1575
   ClientLeft      =   480
   ClientTop       =   825
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3720
   Begin VB.Label AutoText 
      Alignment       =   2  'Zentriert
      Caption         =   "Have a look to my other program - AutoText v3.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   3960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   -240
      X2              =   3720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Info 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEnableExit 
         Caption         =   "E&nable ""Exit"""
      End
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Info"
      Begin VB.Menu mnuInfoCloseSymbol 
         Caption         =   "What about the ""&Close-Symbol"" ? ..."
      End
      Begin VB.Menu Spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfoAutoText 
         Caption         =   "&What is ""AutoText"" ?   ..."
      End
      Begin VB.Menu Spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfoAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Enables API Functions
Private Declare Function GetSystemMenu _
  Lib "user32" _
   (ByVal hWnd As Long, _
    ByVal bRevert As Long) _
  As Long

Private Declare Function DeleteMenu _
  Lib "user32" _
   (ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) _
  As Long

'Exit-Handler
Dim ExitFlag As Integer

Private Sub Form_Load()
 
 Dim SYSMENU As Long
 
 SYSMENU = GetSystemMenu(Me.hWnd, False)
 
 'Disable the close menu and the separator
 DeleteMenu SYSMENU, 6, &H400&
 DeleteMenu SYSMENU, 5, &H400&
 
 'NOT (1) = you can exit
 ExitFlag = 1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  'This is called BEFORE try to unload the forms
  Cancel = ExitFlag
  Info = "Nice try, hehehe"
  
End Sub

Private Sub mnuEnableExit_Click()

  '0 = ready to go out
  ExitFlag = 0
  Info = "Oh...I'm frightend"
  
  'Disable this menu-option
  mnuEnableExit.Enabled = False

End Sub

Private Sub mnuExit_Click()
 
  Unload Me
 
End Sub

Private Sub mnuInfoAbout_Click()

  MsgBox "Only information for new programmers. I know that's" & vbCrLf _
    & "very interesting for newcomers" & vbCrLf & vbCrLf & ";-)  Stefan " _
    & "Ebert", vbInformation, "About..."
  Info = "It's from Stefan"

End Sub

Private Sub mnuInfoAutoText_Click()

  MsgBox "AutoText v3.1 © 2000 by Stefan Ebert" & vbCrLf & vbCrLf _
    & "AutoText remind you on information that you stored earlier. " & vbCrLf _
    & "Put a link simply in the autostart-menu" & vbCrLf _
    & "Editor integrated." & vbCrLf _
    & "at Planet-Source-Code search for ""autotext""", _
    vbExclamation, "AutoText v3.1"
  Info = "Wow ... AutoText"

End Sub

Private Sub mnuInfoCloseSymbol_Click()

  MsgBox "This Form declares two API - calls to make the ''X'' (clo" _
    & "se-Button in Top Line) looks like disabled. But that's only " _
    & "visual. The define check is the ''Form_QueryUnload''-Routine...", _
    vbQuestion, "Close Symbol"
  Info = "Yeaa. Looks Great"

End Sub
