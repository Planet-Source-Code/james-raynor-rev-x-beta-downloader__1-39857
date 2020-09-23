Attribute VB_Name = "modCompression"
Public Const cmpMin = 1
Public Const cmpLow = 3
Public Const cmpMedium = 5
Public Const cmpHigh = 7
Public Const cmpMax = 9

'====================================================================================
'====================================================================================

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
            
Public Declare Function compress Lib "zlib.dll" _
            (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) _
                                                                            As Long
Public Declare Function compress2 Lib "zlib.dll" _
            (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, _
                ByVal level As Long) _
                                                                            As Long
Public Declare Function uncompress Lib "zlib.dll" _
            (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) _
                                                                            As Long
                                                                            
'====================================================================================
'====================================================================================

Public Function CompressByteArray _
            (arrData() As Byte, ByVal iCompressionLevel As Integer) _
                                                                            As Long
    Dim lBufferSize As Long
    Dim arrBytes() As Byte
    
    'Allocate memory for byte array
    lBufferSize = UBound(arrData) + 1
    lBufferSize = lBufferSize + (lBufferSize * 0.01) + 12
    ReDim arrBytes(lBufferSize)
    
    'Compress byte array
    CompressByteArray = compress2(arrBytes(0), lBufferSize, arrData(0), _
                                UBound(arrData) + 1, iCompressionLevel)
    
    'Truncate to compressed size
    ReDim Preserve arrData(lBufferSize - 1)
    CopyMemory arrData(0), arrBytes(0), lBufferSize
    
End Function

Public Function DecompressByteArray _
            (arrData() As Byte, lOriginalSize As Long) _
                                                                            As Long
    Dim lBufferSize As Long
    Dim arrByteArray() As Byte
    
    'Allocate memory for byte array
    lBufferSize = lOriginalSize
    lBufferSize = lBufferSize + (lBufferSize * 0.01) + 12
    ReDim arrBytes(lBufferSize)
    
    'Decompress byte array
    DecompressByteArray = uncompress(arrBytes(0), lBufferSize, arrData(0), _
                                    UBound(arrData) + 1)
    
    'Truncate to compressed size
    ReDim Preserve arrData(lBufferSize - 1)
    CopyMemory arrData(0), arrBytes(0), lBufferSize

End Function

Public Function CompressFile _
            (ByVal sInFile As String, sOutFile As String, ByVal iCompressionLevel) _
                                                                            As Long
    Dim FH As Integer
    Dim arrBytes() As Byte
    Dim lFileLen As Long
    
    lFileLen = FileLen(sInFile)
    
    'Allocate memory for byte array
    ReDim arrBytes(lFileLen - 1)
    
    'Read byte array from input file
    FH = FreeFile
    Open sInFile For Binary Access Read As #FH
        Get #FH, , arrBytes()
    Close #FH
        
    'Compress byte array
    CompressFile = CompressByteArray(arrBytes(), iCompressionLevel)
    
    'Kill any file in place of output file
    If Dir(sOutFile) <> "" Then Kill sOutFile
    
    'Create output file
    FH = FreeFile
    Open sOutFile For Binary Access Write As #FH
        Put #FH, , lFileLen 'must store the length of the original file
        Put #FH, , arrBytes()
    Close #FH
    
    Erase arrBytes

End Function

Public Function DecompressFile _
            (sInFile As String, sOutFile As String) _
                                                                            As Long
    Dim FH As Integer
    Dim arrBytes() As Byte
    Dim lFileLen As Long
    
    If Dir(sInFile) = "" Then
        Err.Description = sInFile & " could not be found."
        Exit Function
    End If

    'Allocate memory for byte array
    ReDim arrBytes(FileLen(sInFile) - 1)
    
    'Read byte array from input file
    FH = FreeFile
    Open sInFile For Binary Access Read As #FH
        Get #FH, , lFileLen 'the original (uncompressed) file's length
        Get #FH, , arrBytes()
    Close #FH
    
    'Decompress byte array
    DecompressFile = DecompressByteArray(arrBytes(), lFileLen)
    
    'Kill any file in place of output file
    If Dir(sOutFile) <> "" Then Kill (sOutFile)
    
    'Create output file
    FH = FreeFile
    Open sOutFile For Binary Access Write As #FH
        Put #FH, , arrBytes()
    Close #FH
    
    Erase arrBytes

End Function
