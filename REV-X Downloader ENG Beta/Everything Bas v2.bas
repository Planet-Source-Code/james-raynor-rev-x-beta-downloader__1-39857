Attribute VB_Name = "Module1"

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

'Constants
Public tempsize
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10



Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub RReg()
'This removes your program to the windows registry
Dim Reg As Object
Set Reg = CreateObject("Wscript.Shell")
Reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUNSERVICES\" & App.EXEName
End Sub

Public Sub RegRun()
' Adds your program to the registry RUN Startup
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
End Sub

Public Sub RemoveRegRun()
' Removes your program to the registry RUN Startup
Dim Reg As Object
Set Reg = CreateObject("Wscript.Shell")
Reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName
End Sub


Public Function GetFileSize(FileName) As String

    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(FileName)


    tempsize = TempStr
    
    If TempStr >= "1024" Then
        'In KB
        TempStr = CCur(TempStr / 1024) & "KB"
    Else


        If TempStr >= "1048576" Then
            'In MB
            TempStr = CCur(TempStr / (1024 * 1024)) & " KB"
        Else
            TempStr = CCur(TempStr) & "B"
        End If

    End If

    'tempsize = TempStr
    GetFileSize = TempStr
    Exit Function
Gfserror:
    GetFileSize = "0B"
    Resume
End Function


Public Function FileExists(FileName As String) As Integer

    On Error Resume Next
        x% = Len(dir$(FileName))
    If Err Or x% = 0 Then FileExists = False Else FileExists = True

End Function
Public Function GetAttrib(FileName) As String

    On Error GoTo GAError
    Dim TempStr As String
    TempStr = GetAttr(FileName)


    If TempStr = "64" Then
        TempStr = "Alias"
    End If



    If TempStr = "32" Then
        TempStr = "Archive"
    End If



    If TempStr = "16" Then
        TempStr = "Directory"
    End If



    If TempStr = "2" Then
        TempStr = "Hidden"
    End If



    If TempStr = "0" Then
        TempStr = "Normal"
    End If



    If TempStr = "1" Then
        TempStr = "ReadOnly"
    End If



    If TempStr = "4" Then
        TempStr = "System"
    End If



    If TempStr = "8" Then
        TempStr = "Volume"
    End If

    GetAttrib = TempStr
    Exit Function
GAError:
    GetAttrib = "Unknown"
    Resume
End Function



Public Sub SetHidden(FileName As String)

    On Error Resume Next
    SetAttr FileName, vbHidden
End Sub



Public Sub SetReadOnly(FileName As String)

    On Error Resume Next
    SetAttr FileName, vbReadOnly
End Sub



Public Sub SetSystem(FileName As String)

    On Error Resume Next
    SetAttr FileName, vbSystem
End Sub



Public Sub SetNormal(FileName As String)

    On Error Resume Next
    SetAttr FileName, vbNormal
End Sub



Public Function GetFileExtension(FileName As String)

    On Error Resume Next
    Dim TempStr As String
    TempStr = Right(FileName, 2)


    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(FileName, 1)
        Exit Function
    Else
        TempStr = Right(FileName, 3)


        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(FileName, 2)
            Exit Function
        Else
            TempStr = Right(FileName, 4)


            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(FileName, 3)
                Exit Function
            Else
                TempStr = Right(FileName, 5)


                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(FileName, 4)
                    Exit Function
                Else
                    GetFileExtension = "Unknown"
                End If

            End If

        End If

    End If

    
End Function



Public Function GetFileDate(FileName As String) As String

    On Error Resume Next
    GetFileDate = FileDateTime(FileName)
End Function



Public Sub DeleteFile(FileName As String)

    On Error GoTo DelError
    Kill FileName
    Exit Sub
DelError:
    Resume
End Sub



Public Sub CopyFile(Source As String, Destination As String)

    On Error GoTo CopyError
    FileCopy Source, Destination
    Exit Sub
CopyError:
    MsgBox "Error copying File"
    Resume
End Sub



Public Sub MoveFile(Source As String, Destination As String)

    On Error GoTo MoveError
    FileCopy Source, Destination
    Kill Source
    Exit Sub
MoveError:
    MsgBox "Error moving File"
    Resume
End Sub



Public Sub MakeDIR(Path As String)

    On Error GoTo DIRError
    MkDir Path
    Exit Sub
DIRError:
    MsgBox "Error creating Directory"
    Resume
End Sub



Public Sub RemoveDIR(Path As String)

    On Error GoTo DIRError2
    RmDir Path
    Exit Sub
DIRError2:
    MsgBox "Error removing Directory"
    Resume
End Sub



Public Sub CloseAllFiles()

    On Error Resume Next
    Reset
End Sub

Sub ShowBlueScrn()

' This sometimes makes your screen go blue.
' To fix it, you may have to reboot.

Shell "file:///c:\aux\aux", 1

End Sub

Sub HideCtrlAltDel() ' Hides your program from the ctrl + alt + delete task list.
Call RegisterServiceProcess(0, 1)
End Sub
Sub Show_CtrlAltDel_Show() ' Shows your program from the ctrl + alt + delete task list.
Call RegisterServiceProcess(0, 0)
End Sub
Function GetData(Text As String, Textstring1 As String, Textstring2 As String) As String


    If InStr(Text, Textstring1) = 0 Or InStr(Text, Textstring2) = 0 Then
        GetData = "STRING Not FOUND"
        Exit Function
    End If
    
    GetData = Mid$((Text), (InStr(Text, Textstring1) + Len(Textstring1)), (InStr(Text, Textstring2) - (Len(Textstring1) + InStr(Text, Textstring1))))
    
End Function

Function MSG_YesNo(TheMsg As String, TheTitle As String) As Boolean

'returns true for yes, false for no

Dim MyMsg

MyMsg = MsgBox(TheMsg, vbYesNo, TheTitle)

If MyMsg = vbYes Then
MSG_YesNo = True
Else
MSG_YesNo = False
End If

End Function

Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hWnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function

Sub CDROM_Toggle(IsOpen As Boolean)

SendMCIString "close all", False
If (App.PrevInstance = True) Then
    End
End If
fCDLoaded = False
If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    End
End If
SendMCIString "set cd time format tmsf wait", True

If IsOpen = True Then
SendMCIString "set cd door open", True
Else
SendMCIString "set cd door closed", True
End If

End Sub
Sub spinAdrive()

On Error Resume Next
Kill "a:\*.fgd"

End Sub

Sub Text_Paste(TheForm As Form)

TheForm.ActiveForm.ActiveControl.SelText = Clipboard.GetText()

End Sub
Sub Text_Cut(TheForm As Form)

Clipboard.SetText TheForm.ActiveForm.ActiveControl.SelText

TheForm.ActiveForm.ActiveControl.SelText = ""

End Sub

Sub Text_Copy(TheForm As Form)

Clipboard.SetText TheForm.ActiveForm.ActiveControl.SelText

End Sub
Function Text_UCase(TheText As String)

Let inptxt$ = TheText
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
    If NextChr$ = UCase(NextChr$) Then
    Let MyString = UCase(NextChr$)
    Else
    If NextChr$ = LCase(NextChr$) Then
    Let MyString = UCase(NextChr$)
    End If
    End If
Let NextChr$ = MyString
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_UCase = NewSent$

End Function
Function Text_LCase(TheText As String)

Let inptxt$ = TheText
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
    If NextChr$ = LCase(NextChr$) Then
    Let MyString = LCase(NextChr$)
    Else
    If NextChr$ = UCase(NextChr$) Then
    Let MyString = LCase(NextChr$)
    End If
    End If
Let NextChr$ = MyString
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_LCase = NewSent$

End Function
Function Text_Spaced(TheText As String)

Let inptxt$ = TheText
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_Spaced = NewSent$

End Function
Function Text_Elite(TheText As String)

Let inptxt$ = TheText
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChrr$ = Mid$(inptxt$, NumSpc%, 2)
If NextChrr$ = "ae" Then Let NextChrr$ = "æ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo dustepp2
If NextChrr$ = "AE" Then Let NextChrr$ = "Æ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo dustepp2
If NextChrr$ = "oe" Then Let NextChrr$ = "œ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo dustepp2
If NextChrr$ = "OE" Then Let NextChrr$ = "Œ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo dustepp2
If Crapp% > 0 Then GoTo dustepp2

If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = ")-("
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "(V)"
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "º"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
Let NewSent$ = NewSent$ + NextChr$

dustepp2:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
Text_Elite = NewSent$

End Function
Function GetBefore(TheData, TheChar) As String

Dim Data$
Dim Stored$
Dim CurChar$
Dim CurCharr$
Dim pos$
Dim Poss$
Dim Length%
Dim Lengthh%
Dim num%

Data$ = ""
Stored$ = ""
CurChar$ = ""
CurCharr$ = ""
pos$ = 0
Poss$ = 0
Length% = Len(TheData)
num% = "0"

Do

pos$ = pos$ + 1
Poss$ = Poss$ + 1

CurChar$ = (Mid$(TheData, Poss$))
CurCharr$ = (Mid$(CurChar$, 1, 1))

Loop Until CurCharr$ = TheChar

num% = num% + 1

Lengthh% = (Len(TheData) - num%)

Stored$ = (Mid$(TheData, 1, Poss$ - 1))

Data$ = Stored$
GetBefore = Data$

End Function



Function GetAfter(TheData, TheChar) As String

Dim Data$
Dim Stored$
Dim CurChar$
Dim CurCharr$
Dim pos$
Dim Poss$
Dim Length%
Dim Lengthh%
Dim num%

Data$ = ""
Stored$ = ""
CurChar$ = ""
CurCharr$ = ""
pos$ = 0
Poss$ = 0
Length% = Len(TheData)
num% = "0"

Do

pos$ = pos$ + 1
Poss$ = Poss$ + 1

CurChar$ = (Mid$(TheData, Poss$))
CurCharr$ = (Mid$(CurChar$, 1, 1))

Loop Until CurCharr$ = TheChar

Stored$ = (Mid$(TheData, Poss$ + 1))

Data$ = Stored$
GetAfter = Data$

End Function
Function Text_Backwards(TheText As String)

Let inptxt$ = TheText
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NewSent$ = NextChr$ & NewSent$
Loop
Text_Backwards = NewSent$

End Function

Sub Text_Print(TheText)

Printer.Print ""
Printer.Print TheText

End Sub


Function MSG_Input(TheMsg, TheTitle) As String

MSG_Input = InputBox(TheMsg, TheTitle)

End Function
Function ReadINI(AppName, SECTION, Key)

GetSetting (AppName), (SECTION), (Key), ("")

End Function
Function DeleteINI_Section(AppName, SECTION)

DeleteSetting (AppName), (SECTION)

End Function

Function DeleteINI_Key(AppName, SECTION, Key)

DeleteSetting (AppName), (SECTION), (Key)

End Function
Function DeleteINI(AppName, SECTION)

DeleteSetting (AppName)

End Function

Function WriteINI(AppName, SECTION, Key, value)

SaveSetting (AppName), (SECTION), (Key), (value)

End Function
Function GetCaption(hWnd)

hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$

End Function
Sub MSG_Information(TheMsg, TheTitle)

MyMsg = TheMsg
MyStyle = vbOKOnly + vbInformation + vbDefaultButton1
MyTitle = TheTitle

Message = MsgBox(MyMsg, MyStyle, MyTitle)

End Sub
Sub MSG_Question(TheMsg, TheTitle)
MyMsg = TheMsg
MyStyle = vbOKOnly + vbQuestion + vbDefaultButton1
MyTitle = TheTitle

Message = MsgBox(MyMsg, MyStyle, MyTitle)
End Sub
Sub MSG_Critical(TheMsg, TheTitle)

MyMsg = TheMsg
MyStyle = vbOKOnly + vbCritical + vbDefaultButton1
MyTitle = TheTitle

Message = MsgBox(MyMsg, MyStyle, MyTitle)

End Sub
Sub MSG(TheMsg, TheTitle)

MyMsg = TheMsg
MyStyle = vbOKOnly
MyTitle = TheTitle

Message = MsgBox(MyMsg, MyStyle, MyTitle)

End Sub

Sub MSG_Warning(TheMsg, TheTitle)
MyMsg = TheMsg
MyStyle = vbOKOnly + vbExclamation + vbDefaultButton1
MyTitle = TheTitle

Message = MsgBox(MyMsg, MyStyle, MyTitle)
End Sub
Sub File_mkdir(TheDiR)

On Error Resume Next

MkDir (TheDiR)

End Sub

Sub File_rmdir(TheDiR)

On Error Resume Next

RmDir (TheDiR)

End Sub

Sub File_KillFile(TheFile)

On Error Resume Next

Kill (TheFile)

End Sub

Sub GotoSite(TheSite)

Dim Action

If (GetBefore(TheSite, "/")) <> "http:" Then
Action = Shell("start.exe http://" + TheSite, 1)
Else
Action = Shell("start.exe " + TheSite, 1)
End If

End Sub

Sub File_Open(TheFile)

On Error Resume Next

Dim Action
Action = Shell(TheFile, 1)

End Sub

Sub Load_Notepad()

Dim Action
Action = Shell("c:\windows\notepad.exe", 1)

End Sub

Sub Load_CDPlayer()

Dim Action
Action = Shell("c:\windows\cdplayer.exe", 1)

End Sub

Sub Load_Calculator()

Dim Action
Action = Shell("c:\windows\calc.exe", 1)

End Sub

Sub Load_Explorer()

Dim Action
Action = Shell("c:\windows\explorer.exe", 1)

End Sub

Sub Load_Solitaire()

Dim Action
Action = Shell("c:\windows\sol.exe", 1)

End Sub

Sub Load_PaintBrush()

Dim Action
Action = Shell("c:\windows\pbrush.exe", 1)

End Sub


Sub OnTop(TheForm As Form, bOnTop As Boolean)

Dim SetOnTop

If bOnTop = True Then
    SetOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
Else
    SetOnTop = SetWindowPos(TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End If

End Sub

Sub FormOnTop(TheForm As Form)

Dim SetOnTop

    SetOnTop = SetWindowPos(TheForm.hWnd, HWND_TOP, 0, 0, 0, 0, Flags)

End Sub

Sub Center(TheObject)

TheObject.Left = (Screen.Width - TheObject.Width) \ 2
TheObject.Top = (Screen.Height - TheObject.Height) \ 2

End Sub
Sub CenterTop(TheObject)

TheObject.Left = (Screen.Width - TheObject.Width) / 2
TheObject.Top = (Screen.Height - TheObject.Height) / (Screen.Height)

End Sub

Sub Pause(interval)

current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop

End Sub
Sub Wait(interval)

current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop

End Sub

Sub Hold(interval)

current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop

End Sub

Public Sub ExecuteFile(FilePath As String)
'Execute a file
On Error GoTo error
ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & (FilePath), vbMaximizedFocus)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

