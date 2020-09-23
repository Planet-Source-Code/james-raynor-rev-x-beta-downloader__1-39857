Attribute VB_Name = "Split"
Option Explicit

Type FileSection
    Bytes() As Byte
    FileLen As Long
End Type
Type SectionedFile
    Files() As FileSection
    NumberOfFiles As Long
End Type
Type FileInfo
    OrigProjSize As Long
    OrigFileName As String
    FileCount As Integer
    FileStartNum As Long
End Type
Type CommReturn
    FileName As String
    Extention As String
    FilePath As String
End Type
 
Public Function Save_Load_File(ShowSave As Boolean, ComDlgCnt As CommonDialog, Filter As String, Flags As Long, DialogTitle As String, Optional FilterIndex As Long) As CommReturn
    On Error Resume Next
    ComDlgCnt.FileName = ""
    ComDlgCnt.Filter = Filter
    ComDlgCnt.Flags = Flags
    ComDlgCnt.FilterIndex = FilterIndex
    ComDlgCnt.DialogTitle = DialogTitle
    If ShowSave Then
        ComDlgCnt.ShowSave
        If Err = cdlCancel Then Exit Function
    Else
        ComDlgCnt.ShowOpen
        If Err = cdlCancel Then Exit Function
    End If
    Save_Load_File.FileName = RetFileName(ComDlgCnt.FileName)
    Save_Load_File.Extention = ReturnExtention(ComDlgCnt.FileName, False)
    Save_Load_File.FilePath = FilePath(ComDlgCnt.FileName)
End Function
Public Function ReturnExtention(FileName As String, ReturnFilename As Boolean) As String
    Dim Buffer1 As String, m_LngLoop As Long, StartPos As Long
    Buffer1 = FileName
    For m_LngLoop = 1 To Len(Buffer1)
        If Mid(Buffer1, m_LngLoop, 1) = "." Then
            StartPos = m_LngLoop
        End If
    Next m_LngLoop
    If StartPos = 0 Then ReturnExtention = ""
    If ReturnFilename = True Then
        ReturnExtention = Mid(Buffer1, 1, StartPos - 1)
    Else
        ReturnExtention = Mid(Buffer1, StartPos + 1)
    End If
End Function
Sub SplitDirName(DirName As String, Lines() As String)
'SplitDirName
'Created By Allen
    If DirName = "" Then Exit Sub
    Dim Text As String, CurNum As Long, TotalNum As Long, CurPos As Long
    Text = DirName
    CurNum = 1
    CurPos = 1
    TotalNum = GetCount(Text, "\")
    ReDim Lines(1 To TotalNum)
    Do Until CurNum = TotalNum + 1
        Lines(CurNum) = Mid(Text, 1, InStr(CurPos, Text, "\") - 1)
        Text = Mid(Text, Len(Lines(CurNum)) + 2)
        CurNum = CurNum + 1
    Loop
End Sub
Public Function GetCount(Text As String, Search As String)
    Dim CCnt As Long, m_LngLoop As Long
    For m_LngLoop = 1 To Len(Text)
        If Mid(Text, m_LngLoop, Len(Search)) = Search Then
            CCnt = CCnt + 1
        End If
    Next
    GetCount = CCnt
End Function
 
Public Function FilePath(FileName As String) As String
    Dim XText As String, DFileName As String, m_LngLoop As Long, DLines() As String
    XText = FileName
    If Not Right(XText, 1) = "\" Then XText = XText & "\"
    SplitDirName CStr(XText), DLines()
    For m_LngLoop = 1 To UBound(DLines) - 1
        DFileName = DFileName & DLines(m_LngLoop) & "\"
    Next
    FilePath = DFileName
End Function
Public Function SplitFile(SplitFileName As String, BeginningNumber As Long, ReturnErrorDes As String, Optional Split As Long = 1439865) As Boolean
    Dim SaveName As String
    Dim fnum As Integer
    
    SplitFile = True 'Assume Success
    On Error GoTo CleanUp
    Dim CurrentFile As SectionedFile, m_lngNumFil As Long, m_LngLoop As Long, FilesLen As Long
    FilesLen = FileLen(SplitFileName)
    If FilesLen <= Split + 1 Then
        SplitFile = False 'If the File
        ' Name is Smaller than the Split Ratio then
        ' The Function Doesnt Need Called So it Fails.
        ReturnErrorDes = "File Is Too Small"
        Exit Function
    End If
    
    fnum = FreeFile
    Open SplitFileName For Binary As fnum
        If CInt(FilesLen / Split) >= _
        FilesLen / Split Or CInt(FilesLen / Split) _
        = FilesLen / Split Then
            m_lngNumFil = CInt(FilesLen _
            / Split)  ' If VB heightened(or if they _
            were equal) the length of the file _
            divided by the total Split ratio then _
            nothing needs To Do anything.
        ElseIf CInt(FilesLen / Split) <= _
        FilesLen / Split Then
            m_lngNumFil = CInt(FilesLen / _
            Split) + 1 ' If VB Lowered The _
            Length Of the File Divided by the Total _
            Split Ratio then it Will Need To Correct _
            it.
        End If
        ReDim CurrentFile.Files(1 To m_lngNumFil)
        For m_LngLoop = 1 To m_lngNumFil - 1
            ReDim CurrentFile.Files(m_LngLoop) _
                .Bytes(1 To Split) 'Re-Define(Re _
                Dimention) the Number Of Bytes Per _
                File
            CurrentFile.Files(m_LngLoop) _
                .FileLen = UBound(CurrentFile.Files _
                (m_LngLoop).Bytes) 'Just For Reference
        Next
        For m_LngLoop = 1 To m_lngNumFil
            Get #fnum, , CurrentFile.Files(m_LngLoop) _
            .Bytes
        Next
        ReDim CurrentFile.Files(m_lngNumFil) _
            .Bytes(1 To FilesLen - ((m_lngNumFil _
            - 1) * Split)) 'ReDefine the Number of _
            bytes for the last file since in many cases _
            it will not be at the Split ratio.
        CurrentFile.NumberOfFiles = m_lngNumFil
        Get #fnum, , CurrentFile.Files(m_lngNumFil) _
        .Bytes
        CurrentFile.Files(m_lngNumFil) _
        .FileLen = UBound(CurrentFile.Files _
        (m_lngNumFil).Bytes)
    Close #fnum 'Close File
    For m_LngLoop = 1 To CurrentFile.NumberOfFiles _
    'Save What We Have Done Into Seperate Files
        SaveName = SplitFileName & "." & Format(BeginningNumber - 1 + m_LngLoop, _
        "00000000#")
        fnum = FreeFile
        Open SaveName For Binary As fnum
            Put #fnum, 1, CurrentFile.Files(m_LngLoop)
        Close #fnum
    Next
    Dim FileInfoFile As FileInfo
    FileInfoFile.FileCount = m_lngNumFil
    FileInfoFile.OrigFileName = SplitFileName
    FileInfoFile.OrigProjSize = FileLen(SplitFileName)
    FileInfoFile.FileStartNum = BeginningNumber
    SaveName = SplitFileName & ".tpl"
    fnum = FreeFile
    Open SaveName For Binary As #fnum
        Put #fnum, , FileInfoFile
    Close #fnum
    Exit Function
CleanUp:
    ReturnErrorDes = Err.Description
    SplitFile = False
    '©Copyright Allen Clark Copeland Jr. 1998
End Function
 
Public Function ReassembleFile(TemplateFileName As String, Optional UseOldFilename As Boolean = True, Optional OutPutName = "C:\Filename.Extention") As Boolean
    Dim FileInfo As FileInfo, OutName As String, _
    File As SectionedFile, m_LngLoop As Long, OpenName
    Dim fnum As Integer
    
    ReassembleFile = True 'Assume Success
    fnum = FreeFile
    Open TemplateFileName For Binary As #fnum
        Get #fnum, , FileInfo 'Get Information on the         Previously Saved File(s)
    Close #fnum
    If UseOldFilename Then
        OutName = FileInfo.OrigFileName
    Else
        OutName = OutPutName
    End If
    ReDim File.Files(1 To FileInfo.FileCount)
    For m_LngLoop = 1 To FileInfo.FileCount
        OpenName = FileInfo.OrigFileName & "." & _
        Format((FileInfo.FileStartNum - 1 + _
        m_LngLoop), "00000000#")
        fnum = FreeFile
        Open OpenName For Binary As #fnum
            Get #fnum, 1, File.Files(m_LngLoop)
        Close #fnum
    Next
    fnum = FreeFile
    Open OutName For Binary As #fnum
        For m_LngLoop = 1 To FileInfo.FileCount
            Put #fnum, , File.Files(m_LngLoop).Bytes
        Next
    Close #fnum
End Function
 
Public Function RetFileName(Text As String)
    Dim XText As String, DLines() As String
    XText = Text
    If Not Right(XText, 1) = "\" Then XText = XText & "\"
    SplitDirName CStr(XText), DLines()
    RetFileName = DLines(UBound(DLines))
End Function


