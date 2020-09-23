<div align="center">

## How to call the Find / Replace DialogBox API \(REALLY\!\)


</div>

### Description

An example of *really* how to call the Find and Find/Replace DialogBoxes using the API and attaching it to your Textbox or RichTextbox!
 
### More Info
 
Here it is guys! Let me know what you think... Since the object of this was to get the Find & Find/Replace dialog boxes working without crashing, I documented that part well. I also included some find and replace code to do the actual replacement in the textbox or RichTextBox, but I didn't do a "backwards search" or a "replace all"... When these buttons are clicked the subroutine fires a messagebox to let you know where to actually put the code... ;-)

On to the code... This is a full demonstration.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[mcrider](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mcrider.md)
**Level**          |Intermediate
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mcrider-how-to-call-the-find-replace-dialogbox-api-really__1-9289/archive/master.zip)





### Source Code

```
1) Start a new project.
2) Add a textbox to Form1... You can make it MultiLine with scrollbars if you want.
3) Add two command buttons to Form1.
4) Add the following code to the Form1 Declarations Section:
'-------------------------------------------------------------------------------
  Private Sub Command1_Click()
    ShowFindDialog FindDialogBox, Me, Text1
  End Sub
  Private Sub Command2_Click()
    ShowFindDialog ReplaceDialogBox, Me, Text1
  End Sub
'-------------------------------------------------------------------------------
5) Add a module to the program and then paste the following code into the Declarations Section of the module:
'-------------------------------------------------------------------------------
  Public Const GWL_WNDPROC = (-4)
  Public Const WM_LBUTTONDOWN = &H201
  Public Const FR_NOMATCHCASE = &H800
  Public Const FR_NOUPDOWN = &H400
  Public Const FR_NOWHOLEWORD = &H1000
  Public Const EM_SETSEL = &HB1
  Public Const MaxPatternLen = 50 'Maximum Pattern Length
  Public Const GD_MATCHWORD = &H410
  Public Const GD_MATCHCASE = &H411
  Public Const GD_SEARCHUP = &H420
  Public Const GD_SEARCHDN = &H421
  Public Const BST_UNCHECKED = &H0
  Public Const BST_CHECKED = &H1
  Public Const BST_INDETERMINATE = &H2
  Public Type FINDREPLACE
    lStructSize As Long     '  size of this struct 0x20
    hwndOwner As Long      '  handle to owner's window
    hInstance As Long      '  instance handle of.EXE that
                  '  contains cust. dlg. template
    flags As Long        '  one or more of the FR_??
    lpstrFindWhat As Long    '  ptr. to search string
    lpstrReplaceWith As Long  '  ptr. to replace string
    wFindWhatLen As Integer   '  size of find buffer
    wReplaceWithLen As Integer '  size of replace buffer
    lCustData As Long      '  data passed to hook fn.
    lpfnHook As Long      '  ptr. to hook fn. or NULL
    lpTemplateName As Long   '  custom template name
  End Type
  Public Enum FR_DIALOG_TYPE
    FindDialogBox = 0
    ReplaceDialogBox = 1
  End Enum
  Public Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA" _
    (pFindreplace As FINDREPLACE) As Long
  Public Declare Function ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" _
    (pFindreplace As FINDREPLACE) As Long
  Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
  Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, _
    ByVal nIDDlgItem As Long) As Long
  Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
  Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
  Public Declare Function IsDlgButtonChecked Lib "user32" _
    (ByVal hDlg As Long, ByVal nIDButton As Long) As Long
  Public Declare Function CheckDlgButton Lib "user32" _
    (ByVal hDlg As Long, ByVal nIDButton As Long, ByVal wCheck As Long) As Long
  Global gOldFindDlgWndHandle As Long
  Global gOldCancelDlgWndHandle As Long
  Global gOldReplaceDlgWndHandle As Long
  Global gOldReplaceAllDlgWndHandle As Long
  Global frText As FINDREPLACE
  Global gHDlg As Long
  Global gFindObj As Object
  Global ghFindCmdBtn As Long     'handle of 'Find Next' command button
  Global ghCancelCmdBtn As Long    'handle of 'Cancel' command button
  Global ghReplaceCmdBtn As Long   'handle of 'Replace' command button
  Global ghReplaceAllCmdBtn As Long  'handle of 'Replace All' command button
  Global gIsDlgReplaceBox As Boolean
  Function FindTextHookProc(ByVal hDlg As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    '=========================================================
    ' This is the window hook function for the Find/Replace
    ' dialog boxes. All of the hooks are handled here!
    '=========================================================
    Dim strPtnFind As String    'pattern string
    Dim hFindTxtBox As Long     'handle of the FIND text box in dialog box
    Dim strPtnReplace As String   'pattern string
    Dim hReplaceTxtBox As Long   'handle of the REPLACE text box in dialog box
    Dim ptnLen As Integer      'actual length read by GetWindowString
    Dim lMatchWord As Boolean    'match word switch
    Dim lMatchCase As Boolean    'match case switch
    Dim lSearchUp As Boolean    'search up switch
    Dim lSearchDn As Boolean    'search down switch
    Dim iVal As Long
    strPtnFind = Space(MaxPatternLen)
    strPtnReplace = Space(MaxPatternLen)
    Select Case uMsg
      Case WM_LBUTTONDOWN
        '=========================================================
        ' We have trapped a button down event!
        '=========================================================
         'DEBUG - FIND ALL OF THE DIALOG ITEMS...
         'For iVal = 0 To 65535
         '  hFindTxtBox = GetDlgItem(gHDlg, iVal)
         '  If Not hFindTxtBox = 0 Then
         '    strPtnFind = Space(MaxPatternLen)
         '    ptnLen = GetWindowText(hFindTxtBox, strPtnFind, MaxPatternLen)
         '    Debug.Print "ITEM " + CStr(iVal) + " - " + strPtnFind
         '  End If
         'Next iVal
         'Get the switches from the dialog box
         lMatchWord = IIf(IsDlgButtonChecked(gHDlg, GD_MATCHWORD) = 1, True, False)
         lMatchCase = IIf(IsDlgButtonChecked(gHDlg, GD_MATCHCASE) = 1, True, False)
         lSearchUp = IIf(IsDlgButtonChecked(gHDlg, GD_SEARCHUP) = 1, True, False)
         lSearchDn = IIf(IsDlgButtonChecked(gHDlg, GD_SEARCHDN) = 1, True, False)
         'Get the FIND pattern string
         hFindTxtBox = GetDlgItem(gHDlg, &H480)
         ptnLen = GetWindowText(hFindTxtBox, strPtnFind, MaxPatternLen)
         strPtnFind = Left$(strPtnFind, ptnLen)
         'Get the REPLACE pattern string IF PRESENT
         hReplaceTxtBox = GetDlgItem(gHDlg, &H481)
         If Not hReplaceTxtBox = 0 Then
           ptnLen = GetWindowText(hReplaceTxtBox, strPtnReplace, MaxPatternLen)
           strPtnReplace = Left$(strPtnReplace, ptnLen)
         End If
         'Call the correct default window procedure
         'Then Customize the window procedure
         Select Case hDlg
           Case ghFindCmdBtn: 'POST PROCESS FIND BUTTON
             If gOldFindDlgWndHandle <> 0 Then
               FindTextHookProc = CallWindowProc(gOldFindDlgWndHandle, _
                 hDlg, uMsg, wParam, lParam)
             End If
             Call EventFindButton(strPtnFind, lMatchWord, lMatchCase, _
              lSearchUp, lSearchDn)
           Case ghCancelCmdBtn: 'PRE PROCESS CANCEL BUTTON
             Call EventCancelButton
             If gOldCancelDlgWndHandle <> 0 Then
               FindTextHookProc = CallWindowProc(gOldCancelDlgWndHandle, _
                 hDlg, uMsg, wParam, lParam)
             End If
           Case ghReplaceCmdBtn: 'POST PROCESS REPLACE BUTTON
             If gOldReplaceDlgWndHandle <> 0 Then
               FindTextHookProc = CallWindowProc(gOldReplaceDlgWndHandle, _
                 hDlg, uMsg, wParam, lParam)
             End If
             Call EventReplaceButton(strPtnFind, strPtnReplace, lMatchWord, _
              lMatchCase, lSearchUp, lSearchDn)
           Case ghReplaceAllCmdBtn: 'POST PROCESS REPLACE ALL BUTTON
             If gOldReplaceAllDlgWndHandle <> 0 Then
               FindTextHookProc = CallWindowProc(gOldReplaceAllDlgWndHandle, _
                 hDlg, uMsg, wParam, lParam)
             End If
             Call EventReplaceAllButton(strPtnFind, strPtnReplace, lMatchWord, _
              lMatchCase, lSearchUp, lSearchDn)
         End Select
      Case Else
        'Call the correct default window procedure
        Select Case hDlg
          Case ghFindCmdBtn:
            If gOldFindDlgWndHandle <> 0 Then
              FindTextHookProc = CallWindowProc(gOldFindDlgWndHandle, _
                hDlg, uMsg, wParam, lParam)
            End If
          Case ghCancelCmdBtn:
            If gOldCancelDlgWndHandle <> 0 Then
              FindTextHookProc = CallWindowProc(gOldCancelDlgWndHandle, _
                hDlg, uMsg, wParam, lParam)
            End If
          Case ghReplaceCmdBtn:
            If gOldReplaceDlgWndHandle <> 0 Then
              FindTextHookProc = CallWindowProc(gOldReplaceDlgWndHandle, _
                hDlg, uMsg, wParam, lParam)
            End If
          Case ghReplaceAllCmdBtn:
            If gOldReplaceAllDlgWndHandle <> 0 Then
              FindTextHookProc = CallWindowProc(gOldReplaceAllDlgWndHandle, _
                hDlg, uMsg, wParam, lParam)
            End If
        End Select
    End Select
  End Function
  Private Sub EventCancelButton()
    '=========================================================
    ' This SUB gets called from FindTextHookProc
    ' when Find/Replace "CANCEL" button is pressed
    '=========================================================
    Dim lngReturnValue As Long
    'UNHOOK ALL OF THE WINDOW HOOKS!!!
    If Not ghFindCmdBtn = 0 Then lngReturnValue = SetWindowLong(ghFindCmdBtn, _
      GWL_WNDPROC, gOldFindDlgWndHandle)
    If Not ghReplaceCmdBtn = 0 Then lngReturnValue = SetWindowLong(ghReplaceCmdBtn, _
      GWL_WNDPROC, gOldReplaceDlgWndHandle)
    If Not ghReplaceAllCmdBtn = 0 Then lngReturnValue = SetWindowLong(ghReplaceAllCmdBtn, _
      GWL_WNDPROC, gOldReplaceAllDlgWndHandle)
    lngReturnValue = SetWindowLong(ghCancelCmdBtn, GWL_WNDPROC, gOldCancelDlgWndHandle)
    'Cleanup the global find object
    Set gFindObj = Nothing
  End Sub
  Private Sub EventFindButton(FindString As String, MatchWord As Boolean, _
    MatchCase As Boolean, SearchUp As Boolean, SearchDn As Boolean)
    '=========================================================
    ' This SUB gets called from FindTextHookProc
    ' when Find/Replace "FIND" button is pressed
    ' gFindObj is the object we need to do stuff to...
    '=========================================================
    Dim sp As Integer        'start point of matching string
    Dim ep As Integer        'end point of matchiing string
    With gFindObj
      SetFocus .hwnd
      If SearchDn = True Or gIsDlgReplaceBox = True Then 'WE'RE DOING A FORWARD SEARCH!
        sp = InStr(IIf(.SelStart = 0, 1, .SelStart) + .SelLength, .Text, _
          IIf(MatchWord, " " + Trim$(FindString) + " ", FindString), _
          IIf(MatchCase, vbBinaryCompare, vbTextCompare))
        sp = IIf(sp = 0, -1, sp - 1)
        If sp = -1 Then
          MsgBox "Cannot find " + Chr$(34) + FindString + Chr$(34) + ".", _
            vbExclamation, "Find"
        Else
          .SelStart = sp
          .SelLength = IIf(MatchWord, Len(" " + Trim$(FindString) + " "), Len(FindString))
        End If
      Else 'WE'RE DOING A BACKWARD SEARCH
        MsgBox "I DIDNT CODE A BACKWARDS SEARCH ;-)", vbInformation, "Find"
      End If
    End With
  End Sub
  Private Sub EventReplaceAllButton(FindString As String, ReplaceString As String, _
    MatchWord As Boolean, MatchCase As Boolean, SearchUp As Boolean, SearchDn As Boolean)
    '=========================================================
    ' This SUB gets called from FindTextHookProc
    ' when Find/Replace "REPLACE ALL" button is pressed
    ' gFindObj is the object we need to do stuff to...
    '=========================================================
    MsgBox "I didn't code a REPLACE ALL Function, but this shows the event firing ;-)" + vbCrLf + _
      "Here are the variables passed into the subroutine... Happy Coding!" + vbCrLf + _
      "MatchWord=" + CStr(MatchWord) + vbCrLf + _
      "MatchCase=" + CStr(MatchCase) + vbCrLf + _
      "SearchUp=" + CStr(SearchUp) + vbCrLf + _
      "SearchDn=" + CStr(SearchDn) + vbCrLf + _
      "FindString=" + FindString + vbCrLf + _
      "ReplaceString=" + ReplaceString
  End Sub
  Private Sub EventReplaceButton(FindString As String, ReplaceString As String, _
    MatchWord As Boolean, MatchCase As Boolean, SearchUp As Boolean, SearchDn As Boolean)
    '=========================================================
    ' This SUB gets called from FindTextHookProc
    ' when Find/Replace "REPLACE" button is pressed
    ' gFindObj is the object we need to do stuff to...
    '=========================================================
    With gFindObj
      'WE'RE DOING A FORWARD SEARCH ALWAYS!
      SetFocus .hwnd
      'Replace the highlighted text, if any
      If Not .SelLength = 0 Then
        .SelText = ReplaceString
        .SelLength = 0
      End If
      'Find the next occurrence
      sp = InStr(IIf(.SelStart = 0, 1, .SelStart) + .SelLength, .Text, _
        IIf(MatchWord, " " + Trim$(FindString) + " ", FindString), _
        IIf(MatchCase, vbBinaryCompare, vbTextCompare))
      sp = IIf(sp = 0, -1, sp - 1)
      If sp = -1 Then
        MsgBox "At end of text.", vbInformation, "Find"
      Else
        .SelStart = sp
        .SelLength = IIf(MatchWord, Len(" " + Trim$(FindString) + " "), Len(FindString))
      End If
      .SetFocus
    End With
  End Sub
  Public Sub ShowFindDialog(DialogType As FR_DIALOG_TYPE, ParentObject As Object, _
    TargetObject As Object, Optional DefaultFindText, Optional DefaultReplaceText, _
    Optional DialogBoxFlags)
    '============================================================================
    ' This subroutine is a wrapper to call the FIND and FIND/REPLACE DialogBoxes
    '
    ' Arguments are:
    '
    '  DialogType     : 0=Show FindDialogBox, 1=Show ReplaceDialogBox
    '
    '  ParentObject    : Form that will be the parent of the DialogBox
    '
    '  TargetObject    : Textbox object to search/replace text
    '
    '  DefaultFindText   : OPTIONAL Initializes the "Find Text" TextBox
    '
    '  DefaultReplaceText : OPTIONAL Initialized the "Replace Text" Textbox
    '
    '  DialogBoxFlags   : OPTIONAL Turns off items in the DialogBox
    '             Values can be:
    '              FR_NOMATCHCASE Or FR_NOUPDOWN Or FR_NOWHOLEWORD
    '============================================================================
    Dim szFindString As String   'initial string to find
    Dim szReplaceString As String  'initial string to find
    Dim strFindArr() As Byte    'for API use
    Dim strReplaceArr() As Byte   'for API use
    Dim iVal As Long        'position indicator in the loop
    'Get the default strings to plug into the dialogbox, if present
    szFindString = IIf(IsMissing(DefaultFindText) = True, "", CStr(DefaultFindText)) + Chr$(0)
    ReDim strFindArr(0 To Len(szFindString) - 1)
    For iVal = 1 To Len(szFindString)
      strFindArr(iVal - 1) = Asc(Mid(szFindString, iVal, 1))
    Next iVal
    szReplaceString = IIf(IsMissing(DefaultReplaceText) = True, "", CStr(DefaultReplaceText)) + Chr$(0)
    ReDim strReplaceArr(0 To Len(szReplaceString) - 1)
    For iVal = 1 To Len(szReplaceString)
      strReplaceArr(iVal - 1) = Asc(Mid(szReplaceString, iVal, 1))
    Next iVal
    'Fill in the frText data...
    With frText
      .flags = IIf(IsMissing(DialogBoxFlags) = True, 0, DialogBoxFlags)
      .lpfnHook = 0&
      .lpTemplateName = 0&
      .lStructSize = Len(frText)
      .hwndOwner = ParentObject.hwnd
      .hInstance = App.hInstance
      .lpstrFindWhat = VarPtr(strFindArr(0))
      .wFindWhatLen = Len(szFindString)
      .lpstrReplaceWith = VarPtr(strReplaceArr(0))
      .wReplaceWithLen = Len(szReplaceString)
      .lCustData = 0
    End With
    'Set the object we're going to be doing the find/replace with
    Set gFindObj = TargetObject
    'Show the dialog box.
    If DialogType = FindDialogBox Then
      gHDlg = FindText(frText)
      gIsDlgReplaceBox = False
    Else
      gHDlg = ReplaceText(frText)
      gIsDlgReplaceBox = True
    End If
    'Set the "Search Down" radio button.
    CheckDlgButton gHDlg, GD_SEARCHUP, BST_UNCHECKED
    CheckDlgButton gHDlg, GD_SEARCHDN, BST_CHECKED
    'Get the handles of the dialog box
    ghFindCmdBtn = GetDlgItem(gHDlg, 1) 'FIND BUTTON
    ghCancelCmdBtn = GetDlgItem(gHDlg, 2) 'CANCEL BUTTON
    ghReplaceCmdBtn = GetDlgItem(gHDlg, 1024) 'REPLACE BUTTON
    ghReplaceAllCmdBtn = GetDlgItem(gHDlg, 1025) 'REPLACE ALL BUTTON
    'Hook all of the necessary default window procedures for the dialog box.
    If Not ghFindCmdBtn = 0 Then
      gOldFindDlgWndHandle = GetWindowLong(ghFindCmdBtn, GWL_WNDPROC)
      If SetWindowLong(ghFindCmdBtn, GWL_WNDPROC, AddressOf FindTextHookProc) = 0 _
        Then gOldFindDlgWndHandle = 0
    End If
    If Not ghCancelCmdBtn = 0 Then
      gOldCancelDlgWndHandle = GetWindowLong(ghCancelCmdBtn, GWL_WNDPROC)
      If SetWindowLong(ghCancelCmdBtn, GWL_WNDPROC, AddressOf FindTextHookProc) = 0 _
        Then gOldCancelDlgWndHandle = 0
    End If
    If Not ghReplaceCmdBtn = 0 Then
      gOldReplaceDlgWndHandle = GetWindowLong(ghReplaceCmdBtn, GWL_WNDPROC)
      If SetWindowLong(ghReplaceCmdBtn, GWL_WNDPROC, AddressOf FindTextHookProc) = 0 _
        Then gOldReplaceDlgWndHandle = 0
    End If
    If Not ghReplaceAllCmdBtn = 0 Then
      gOldReplaceAllDlgWndHandle = GetWindowLong(ghReplaceAllCmdBtn, GWL_WNDPROC)
      If SetWindowLong(ghReplaceAllCmdBtn, GWL_WNDPROC, AddressOf FindTextHookProc) = 0 _
        Then gOldReplaceAllDlgWndHandle = 0
    End If
  End Sub
'-------------------------------------------------------------------------------
6) Run the program and type some text into the textbox. then put the cursor in the textbox at the top of the textbox.
7) Click "Command1" and the Find Dialog box will show. Try the box out!!
8) Put the cursor in the textbox back at the beginning of the textbox and click "Command2". The Find/Replace dialog box will show... Try it out!
I have included setting the search textbox and the replace textbox in this code, so if you wanted to populate it before showing the dialogbox, call ShowFindDialog like this:
  ShowFindDialog FindDialogBox, Me, Text1, "Find This"
  ShowFindDialog ReplaceDialogBox, Me, Text1, "Find This", "Replace with this"
You can also add another optional argument to disable parts of the dialogbox... ;-)
```

