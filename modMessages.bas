Attribute VB_Name = "modMessages"
' ***************************************************************************
'  Module:      modMessages.bas
'
'  Purpose:     This module contains routines designed to provide standard
'               formatting for message boxes.  One routine can change the
'               captions on a message box.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added custom message box routine
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added custom message box routine
' 23-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated MessageBoxH() routine on the way button captions
'                are determined.
'              - Renamed MsgBoxHookProc() to MsgboxCallBack() for easier
'                maintenance.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Global constants
' ***************************************************************************
  Public Const IDOK         As Long = 1  ' one button return value
  Public Const IDYES        As Long = 6
  Public Const IDNO         As Long = 7
  Public Const IDCANCEL     As Long = 2
  Public Const DUMMY_NUMBER As Long = vbObjectError + 513

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MB_OK          As Long = &H0&     ' one button
  Private Const MB_YESNO       As Long = &H4&     ' two buttons
  Private Const MB_YESNOCANCEL As Long = &H3&     ' three buttons
  Private Const GWL_HINSTANCE  As Long = &HFFFA   ' (-6)
  Private Const IDPROMPT       As Long = &HFFFF&
  Private Const WH_CBT         As Long = 5
  Private Const HCBT_ACTIVATE  As Long = 5

' ***************************************************************************
' Type structures
' ***************************************************************************
  ' UDT for passing data through the hook
  Private Type MSGBOX_HOOK_PARAMS
      hwndOwner As Long
      hHook     As Long
  End Type

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumMSGBOX_ICON
      eMSG_NOICON = &H0&            ' 0
      eMSG_ICONSTOP = &H10&         ' 16
      eMSG_ICONQUESTION = &H20&     ' 32
      eMSG_ICONEXCLAMATION = &H30&  ' 48
      eMSG_ICONINFORMATION = &H40&  ' 64
  End Enum
  
' ***************************************************************************
' Global API Declarations
' ***************************************************************************
  ' The GetDesktopWindow function returns a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted.
  Public Declare Function GetDesktopWindow Lib "user32" () As Long

' ***************************************************************************
' Module API Declarations
' ***************************************************************************
  ' Retrieves the thread identifier of the calling thread.
  Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

  ' The GetWindowLong function retrieves information about the specified
  ' window. The function also retrieves the 32-bit   (long) value at the
  ' specified offset into the extra window memory.
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
          (ByVal hwnd As Long, ByVal nIndex As Long) As Long

  ' Displays a modal dialog box that contains a system icon, a set of
  ' buttons, and a brief application-specific message, such as status
  ' or error information. The message box returns an integer value that
  ' indicates which button the user clicked.
  Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" _
          (ByVal hwnd As Long, ByVal lpText As String, _
          ByVal lpCaption As String, ByVal wType As Long) As Long
   
  ' The SetDlgItemText function sets the title or text of a control
  ' in a dialog box.
  Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
          (ByVal hDlg As Long, ByVal nIDDlgItem As Long, _
          ByVal lpString As String) As Long
      
  ' The SetWindowsHookEx function installs an application-defined hook
  ' procedure into a hook chain. You would install a hook procedure to
  ' monitor the system for certain types of events. These events are
  ' associated either with a specific thread or with all threads in the
  ' same desktop as the calling thread.
  Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
          (ByVal idHook As Long, ByVal lpfn As Long, _
          ByVal hmod As Long, ByVal dwThreadId As Long) As Long
   
  ' The SetWindowText function changes the text of the specified window  ' s
  ' title bar   (if it has one). If the specified window is a control, the
  ' text of the control is changed. However, SetWindowText cannot change
  ' the text of a control in another application.
  Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
          (ByVal hwnd As Long, ByVal lpString As String) As Long

  ' The UnhookWindowsHookEx function removes a hook procedure installed
  ' in a hook chain by the SetWindowsHookEx function.
  Private Declare Function UnhookWindowsHookEx Lib "user32" _
          (ByVal hHook As Long) As Long
    
' ***************************************************************************
' Global Variables
'
'                    +------------------- Global level designator
'                    |  +---------------- Data type (Boolean)
'                    |  |       |-------- Variable subname
'                    - --- --------------
' Naming standard:   g bln StopProcessing
' Variable name:     gblnStopProcessing
'
' ***************************************************************************
  Public gblnStopProcessing As Boolean
  
' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (String)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a str Captions
' Variable name:     mastrCaptions
'
' ***************************************************************************
  Private mlngButtonCount As Long
  Private mstrTitle       As String
  Private mstrPrompt      As String
  Private mastrCaptions() As String
  Private mtypMsgHook     As MSGBOX_HOOK_PARAMS
  
  
' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
'  Routine:     InfoMsg
'
'  Description: Displays a VB MsgBox with no return values.  It is designed to
'               be used where no response from the user is expected other than
'               "OK".
'
'  Parameters:  strMsg - The message text
'               strCaption - The MsgBox caption (optional)
'
'  Returns:     None
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub InfoMsg(ByVal strMsg As String, _
          Optional ByVal strCaption As String = vbNullString)
                   
    Dim strNewCaption As String  ' Formatted MsgBox caption
                           
    ' Format the MsgBox caption
    strNewCaption = FormatCaption(strCaption)
    
    ' the MsgBox routine
    MsgBox strMsg, vbInformation Or vbOKOnly, strNewCaption
    
End Sub


' ***************************************************************************
'  Routine:     ResponseMsg
'
'  Description: Displays a standard VB MsgBox and returns the MsgBox code. It
'               is designed to be used when the user is prompted for a
'               response.
'
'  Parameters:  strMsg - The message text
'               lngButtons - The standard VB MsgBox buttons (optional)
'               strCaption - The msgbox caption (optional)
'
'  Returns:     The standard VB MsgBox return values
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function ResponseMsg(ByVal strMsg As String, _
                   Optional ByVal lngButtons As Long = vbQuestion + vbYesNo, _
                   Optional ByVal strCaption As String = vbNullString) As VbMsgBoxResult
    
    Dim strNewCaption As String  ' Formatted MsgBox caption
    
    ' Format the MsgBox caption
    strNewCaption = FormatCaption(strCaption)
    
    ' the MsgBox routine and return the user's response
    ResponseMsg = MsgBox(strMsg, lngButtons, strNewCaption)
    
End Function

' ***************************************************************************
'  Routine:     ErrorMsg
'
'  Description: Displays a standard VB MsgBox formatted to display severe
'               (Usually application-type) error messages.
'
'  Parameters:  strModule - The module where the error occurred
'               strRoutine - The routine where the error occurred
'               strMsg - The error message
'               strCaption - The MsgBox caption  (optional)
'
'  Returns:     None
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub ErrorMsg(ByVal strModule As String, _
                    ByVal strRoutine As String, _
                    ByVal strMsg As String, _
           Optional ByVal strCaption As String = vbNullString)
                     
    Dim strNewCaption As String  ' Formatted MsgBox caption
    Dim strFullMsg As String     ' Formatted message
    
    ' Make sure strModule is populated
    If Len(Trim$(strModule)) = 0 Then
       strModule = "Unknown"
    End If
    
    ' Make sure strRoutine is populated
    If Len(Trim$(strRoutine)) = 0 Then
       strRoutine = "Unknown"
    End If
    
    ' Make sure strMsg is populated
    If Len(Trim$(strMsg)) = 0 Then
       strMsg = "Unknown"
    End If
    
    ' Format the MsgBox caption
    strNewCaption = FormatCaption(strCaption, True)
    
    ' Format the message
    strFullMsg = "Module: " & vbTab & strModule & vbCr & _
                 "Routine:" & vbTab & strRoutine & vbCr & _
                 "Error:  " & vbTab & strMsg
                     
    ' the MsgBox routine
    MsgBox strFullMsg, vbCritical Or vbOKOnly, strNewCaption
    
End Sub

' ***************************************************************************
' Routine:       MessageBoxH
'
' Description:   Displays a standard msgbox with customized captions on
'                the buttons.  Wrapper function for the MessageBox API.
'
' Reference:     VBNet - API calls for Visual Basic 6.0
'                http://vbnet.mvps.org/
'
' Parameters:    hwndForm - Long integer system ID designating the form
'                hwndWindow - Long integer system ID designating the
'                           desktop window
'                strPrompt - Main body of text for msgbox
'                strTitle - Caption of msgbox
'                astrCaptions() - String array designating button text
'                           for up to three buttons
'                lngIcon - [Optional] - Designates type of icon to use
'                           Default - no icon
'
' Example:       ' Prepare message box display somewhere in your application.
'                '
'                ' These are the button captions,
'                ' in order, from left to right.
'                ReDim astrMsgBox(3)
'                astrMsgBox(0) = "Encrypt"
'                astrMsgBox(1) = "Decrypt"
'                astrMsgBox(2) = "Cancel"
'                
'                ' Prompt user with message box
'                Select Case MessageBoxH(Me.Hwnd, GetDesktopWindow(), _
'                                        "What do you want to do?  ", _
'                                        PGM_NAME, astrMsgBox(), eMSG_ICONQUESTION)
'                       
'                       ' These are valid responses
'                       Case IDYES:    lngEncrypt = eCA_ENCRYPT
'                       Case IDNO:     lngEncrypt = eCA_DECRYPT
'                       Case IDCANCEL: Exit Sub
'                End Select
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Aug-2008  Randy Birch
'              http://vbnet.mvps.org/code/hooks/messageboxhook.htm
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 23-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated the way button captions are determined
' ***************************************************************************
Public Function MessageBoxH(ByVal hwndForm As Long, _
                            ByVal hwndWindow As Long, _
                            ByVal strPrompt As String, _
                            ByVal strTitle As String, _
                            ByRef astrCaptions() As String, _
                   Optional ByVal lngIcon As enumMSGBOX_ICON = eMSG_NOICON) As Long

    Dim lngIndex    As Long
    Dim hInstance   As Long
    Dim hThreadId   As Long
    Dim lngButtonID As Long
    
    Erase mastrCaptions()                      ' Always start with empty arrays
    mlngButtonCount = UBound(astrCaptions)     ' Determine number of buttons needed
    mstrPrompt = strPrompt                     ' Save msgbox text

    ' Format message box caption
    mstrTitle = IIf(Len(Trim$(strTitle)) > 0, strTitle, FormatCaption(strTitle))
    
    ' If array size has been exceeded then
    ' reset button count to max allowed
    If mlngButtonCount > 3 then
    	mlngButtonCount = 3   
    End If	
    
    ReDim mastrCaptions(mlngButtonCount)       ' Size array to number of captions
    
    ' Transfer captions to module array
    For lngIndex = 0 To mlngButtonCount - 1
        mastrCaptions(lngIndex) = astrCaptions(lngIndex)
    Next lngIndex
                
    Select Case mlngButtonCount
           Case 1: lngButtonID = MB_OK
           Case 2: lngButtonID = MB_YESNO
           Case 3: lngButtonID = MB_YESNOCANCEL
           Case Else
                MessageBoxH = IDCANCEL
                Exit Function
    End Select
    
    ' Set up the hook
    hInstance = GetWindowLong(hwndForm, GWL_HINSTANCE)
    hThreadId = GetCurrentThreadId()
    
    ' set up the MSGBOX_HOOK_PARAMS values
    ' By specifying a Windows hook as one
    ' of the params, we can intercept messages
    ' sent by Windows and thereby manipulate
    ' the dialog
    With mtypMsgHook
        .hwndOwner = hwndWindow
        .hHook = SetWindowsHookEx(WH_CBT, _
                                  AddressOf MsgboxCallBack, _
                                  hInstance, _
                                  hThreadId)
    End With
    
    ' Call MessageBox API and return the
    ' value as the result of the function
    MessageBoxH = MessageBox(hwndWindow, _
                             mstrPrompt, _
                             mstrTitle, _
                             lngButtonID Or lngIcon)

End Function



' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function MsgboxCallBack(ByVal hInstance As Long, _
                                ByVal hThreadId As Long, _
                                ByVal lngNotUsed As Long) As Long

    ' Called by MessageBoxH()
    
    ' When the message box is about to be shown,
    ' titlebar text, prompt message and button
    ' captions will be updated
    DoEvents
    If hInstance = HCBT_ACTIVATE Then

    
        ' In a HCBT_ACTIVATE message, hThreadId 
        ' holds the handle to the messagebox
        SetWindowText hThreadId, mstrTitle
                  
        ' The ID's of the buttons on the message box
        ' correspond exactly to the values they return,
        ' so the same values can be used to identify
        ' specific buttons in a SetDlgItemText call.
        '
        ' Use default captions if array elements are empty
        Select Case mlngButtonCount
               Case 1
                    SetDlgItemText hThreadId, IDOK, IIf(Len(Trim$(mastrCaptions(0))) > 0, mastrCaptions(0), "OK")
               Case 2
                    SetDlgItemText hThreadId, IDYES, IIf(Len(Trim$(mastrCaptions(0))) > 0, mastrCaptions(0), "Yes")
                    SetDlgItemText hThreadId, IDNO, IIf(Len(Trim$(mastrCaptions(1))) > 0, mastrCaptions(1), "No")
               Case 3
                    SetDlgItemText hThreadId, IDYES, IIf(Len(Trim$(mastrCaptions(0))) > 0, mastrCaptions(0), "Yes")
                    SetDlgItemText hThreadId, IDNO, IIf(Len(Trim$(mastrCaptions(1))) > 0, mastrCaptions(1), "No")
                    SetDlgItemText hThreadId, IDCANCEL, IIf(Len(Trim$(mastrCaptions(2))) > 0, mastrCaptions(2), "Cancel")
        End Select
        
        ' Change dialog prompt text
        SetDlgItemText hThreadId, IDPROMPT, mstrPrompt
                                               
        ' Finished with the dialog, release the hook
        UnhookWindowsHookEx mtypMsgHook.hHook
             
    End If
    
    ' Return False to let normal processing continue
    MsgboxCallBack = 0

End Function

' ***************************************************************************
'  Routine:     FormatCaption
'
'  Description: Formats the caption text to use the application title as
'               default
'
'  Parameters:  strCaption - The input caption which may be appended to the
'                            application title.
'               bError - Add "Error" to the caption
'
'  Returns:     Formatted string to be used as a msgbox caption
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function FormatCaption(ByVal strCaption As String, _
                      Optional ByVal bError As Boolean = False) As String

    ' Called by InfoMsg()
    '           ResponseMsg()
    '           ErrorMsg()  

    Dim strNewCaption As String  ' The formatted caption
    
    ' Set the caption to either input parm or the application name
    If Len(Trim$(strCaption)) > 0 Then
        strNewCaption = Trim$(strCaption)
    Else
        ' Set the caption default
        strNewCaption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    End If
    
    ' Optionally, add error text
    If bError Then
        strNewCaption = strNewCaption & " Error"
    End If

    ' Return the new caption
    FormatCaption = strNewCaption
    
End Function


