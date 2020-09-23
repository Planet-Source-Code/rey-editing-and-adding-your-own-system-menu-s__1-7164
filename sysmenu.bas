Attribute VB_Name = "sysmen"
  Option Explicit
' use this code at your own risk
' reyspage.cjb.net
  Public Const WM_SYSCOMMAND As Long = &H112&
  Public Const IDM_ABOUT As Long = 1&
  Public Const IDM_REY As Long = 2&
 ' U can use any id that windows isn't already using
  Public procOld As Long
  
  Public Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, _
                                                    ByVal hwnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)
  
' Be Careful with the Debug on this
' it can cause a crash
  
Public Function MenuProc(ByVal hwnd As Long, ByVal iMsg As Long, _
                                              ByVal wParam As Long, ByVal lParam As Long) As Long
      Dim ReturnVal As Long

  Select Case iMsg
    Case WM_SYSCOMMAND
       ' this checks if it was a WM_SYSCOMMAND
       ' and not something else
      If wParam = IDM_ABOUT Then
        MsgBox "Testing 123", vbInformation, "Testing 123"
        frmsysmen.Text1.Text = "Testing 123"
      End If
        If wParam = IDM_REY Then
        ReturnVal& = Shell("Start.exe " & "http://reyspage.cjb.net", vbHide)
      ' Launches My Site When the IDM_Rey is clicked
      ' in the system menu
      End If
    
  End Select
  ' returns all messages to vb for processing
  MenuProc = CallWindowProc(procOld, hwnd, iMsg, wParam, lParam)

End Function


