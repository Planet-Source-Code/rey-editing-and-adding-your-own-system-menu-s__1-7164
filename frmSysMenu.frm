VERSION 5.00
Begin VB.Form frmsysmen 
   AutoRedraw      =   -1  'True
   Caption         =   "sysmenu demo"
   ClientHeight    =   2520
   ClientLeft      =   2370
   ClientTop       =   1425
   ClientWidth     =   2730
   Icon            =   "frmSysMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   182
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmSysMenu.frx":030A
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit This Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "frmsysmen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
  ' demo project showing how to manipulate a form's system menu
  ' by Bryan Stafford of New Vision Software® - newvision@imt.net
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.
  

  Private Const SC_SIZE As Long = &HF000&
  Private Const SC_MOVE As Long = &HF010&
  Private Const SC_CLOSE As Long = &HF060&
  Private Const SC_MINIMIZE As Long = &HF020&
  Private Const SC_MAXIMIZE As Long = &HF030&
  Private Const SC_NEXTWINDOW As Long = &HF040&
  Private Const SC_PREVWINDOW As Long = &HF050&
  Private Const MF_BYCOMMAND As Long = &H0&
  
  Private Const MF_STRING As Long = &H0&
  Private Const MF_SEPARATOR As Long = &H800&

  Private Const GWL_WNDPROC As Long = (-4&)
    
  Private Declare Function GetSystemMenu& Lib "user32" (ByVal hwnd&, ByVal bRevert&)
  
  Private Declare Function DeleteMenu& Lib "user32" (ByVal hMenu&, _
                                                              ByVal nPosition&, ByVal wFlags&)
  
  Private Declare Function AppendMenu& Lib "user32" Alias "AppendMenuA" (ByVal hMenu&, _
                                            ByVal wFlags&, ByVal wIDNewItem&, lpNewItem As Any)
                                                                    
  Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, _
                                                              ByVal nIndex&, ByVal dwNewLong&)
                                                              
Private Sub Command1_Click()
  ' the user want's out, so let them out
  Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
  Screen.MousePointer = vbHourglass
  Dim hSysMenu&
  hSysMenu = GetSystemMenu(hwnd, False)
 ' delete's system menu's
 ' you can get rid of the one's
 ' you want to keep
  Call DeleteMenu(hSysMenu, SC_CLOSE, MF_BYCOMMAND)
  Call DeleteMenu(hSysMenu, SC_SIZE, MF_BYCOMMAND)
  Call DeleteMenu(hSysMenu, SC_MOVE, MF_BYCOMMAND)
  Call DeleteMenu(hSysMenu, SC_MAXIMIZE, MF_BYCOMMAND)
  ' these add our new menu's
  ' for a seperator you can use:
  'Call AppendMenu(hSysMenu, MF_SEPARATOR, False, ByVal 0&)
  Call AppendMenu(hSysMenu, MF_STRING, IDM_ABOUT, ByVal "Te&sting")
  Call AppendMenu(hSysMenu, MF_SEPARATOR, False, ByVal 0&)
  Call AppendMenu(hSysMenu, MF_STRING, IDM_REY, ByVal "&Reys Escape")
 
  ' take control of message processing by installing our message handling
  ' routine into the chain of message routines for this window
  procOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MenuProc)
                          
  Screen.MousePointer = vbDefault
  
cantgetsysmenu:
  If Err Then
    Err.Clear
 
    MsgBox "Unable to load append system menu.", vbExclamation, "System Menu Demo"
  
    Resume cantgetsysmenu
  End If
  
End Sub


Private Sub Form_Unload(Cancel As Integer)

  ' give message processing control back to VB
  ' if you don't do this you WILL crash!!!
  Call SetWindowLong(hwnd, GWL_WNDPROC, procOld)
  
End Sub

Private Sub Label2_Click()

End Sub
