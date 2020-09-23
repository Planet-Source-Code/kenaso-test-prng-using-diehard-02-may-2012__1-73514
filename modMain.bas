Attribute VB_Name = "modMain"
' ***************************************************************************
' Routine:   modMain
'
' Purpose:   Example of creating PRNG (PseudoRandom Number Generator) binary
'            test files for testings with Diehard or ENT software.  See
'            "Using Diehard.pdf" for more information about using Diehard.
'
' ---------------------------------
' Randomness testing software
' ---------------------------------
' Diehard by George Marsaglia
' http://stat.fsu.edu/pub/diehard/
'
' ENT - A Pseudorandom Number Sequence Test Program
' http://www.fourmilab.ch/random/
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Created module
' 10-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote ElapsedTime() routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const KB_32     As Long = &H8000&     ' 32768
  Private Const KB_64     As Long = &H10000     ' 65536
  Private Const MIN_LONG  As Long = &H80000000  ' -2147483648
  Private Const MAX_LONG  As Long = &H7FFFFFFF  '  2147483647
  Private Const FILE_SIZE As Long = &HAF0000    '  11468800  Max output file size

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


' ****************************************************************************
' This will build the binary input file (Approx 11mb) needed when using
' Diehard or ENT for randomness testing.
'
' Test file contains 2,867,200 32-bit long integers (11,468,800 bytes)
'
' ****************************************************************************
Public Sub Main()
    
    ' Use 8.3 file naming standards
    ' To run, press F5
    ' Only takes a few seconds to complete
    
    Dim strPath  As String
    Dim strTime1 As String
    Dim strTime2 As String
    Dim objFSO   As Scripting.FileSystemObject
    
    Screen.MousePointer = vbHourglass   ' set cursor to show hourglass
    
    ' Prepare target path by appending backslashes where needed
    strPath = QualifyPath(App.Path) & QualifyPath("Test_Files")
        
    Set objFSO = New Scripting.FileSystemObject
        
    ' See if there is a temp folder availalble
    If Not objFSO.FolderExists(strPath) Then
        objFSO.CreateFolder strPath
    End If
        
    Set objFSO = Nothing
    
    strTime1 = Test1(strPath & "Crypto.bin")   ' MS CryptoAPI random number generator
    strTime2 = Test2(strPath & "VB_RND.bin")   ' Alfred Hellmüller's enhanced algorithm
    Screen.MousePointer = vbDefault            ' reset cursor to normal
    
    MsgBox "Located in " & strPath & Space$(5) & vbNewLine & vbNewLine & _
           "Crypto.bin" & Space$(4) & strTime1 & vbNewLine & _
           "VB_Rnd.bin" & Space$(4) & strTime2, vbOKOnly, "Create Test files"
           
End Sub

Private Function Test1(ByVal strPath As String) As String
    
    Dim strElapsed As String
    Dim hFile      As Integer
    Dim lngLoop    As Long
    Dim lngStart   As Long
    Dim lngPointer As Long
    Dim lngByteCnt As Long
    Dim abytData() As Byte
    Dim objPrng    As cPrng

    Set objPrng = New cPrng   ' Instantiate class object
    
    lngPointer = 1    ' output file pointer position
    Erase abytData()  ' empty arrays
    
    ' Verify receiving file is empty
    hFile = FreeFile                   ' capture first free file handle
    Open strPath For Output As #hFile  ' Create an empty file
    Close #hFile                       ' close file handle

    hFile = FreeFile                                 ' capture first free file handle
    Open strPath For Binary Access Write As #hFile   ' re-open file in binary mode

    '-----------------------------------------------------------------------------
    lngStart = GetTickCount()  ' starting time
    
    With objPrng
        ' Generate some random data
        For lngLoop = 1 To 175
    
            abytData() = .BuildRndData(KB_64, ePRNG_BYTE_ARRAY)  ' Create random data
            ReDim Preserve abytData(KB_64 - 1)                   ' Size byte array (0-65535)
            Put #hFile, lngPointer, abytData()                   ' write byte array to output file
            lngPointer = lngPointer + UBound(abytData)           ' update pointer within output file
            Erase abytData()                                     ' empty temp array
            
        Next lngLoop
        
        abytData() = .BuildRndData(KB_64, ePRNG_BYTE_ARRAY)      ' Create random data
    End With
        
    lngByteCnt = FILE_SIZE - lngPointer  ' Calc byte count
    ReDim Preserve abytData(lngByteCnt)  ' Size byte array
    Put #hFile, lngPointer, abytData()   ' write byte array to output file
    
    strElapsed = ElapsedTime(GetTickCount() - lngStart)  ' Finish time
    '-----------------------------------------------------------------------------
    
    Close #hFile           ' close output file
    Erase abytData()       ' Always empty arrays when not needed
    Set objPrng = Nothing  ' Free object from memory
        
    Test1 = strElapsed  ' Return elapsed time
    
End Function

Private Function Test2(ByVal strPath As String) As String
    
    Dim strElapsed As String
    Dim hFile      As Integer
    Dim lngIdx     As Long
    Dim lngLoop    As Long
    Dim lngStart   As Long
    Dim lngByteCnt As Long
    Dim lngPointer As Long
    Dim alngData() As Long
    Dim abytData() As Byte
    
    lngPointer = 1             ' output file pointer position
    lngByteCnt = (KB_32 * 4)   ' Calc byte count
    
    Rnd (-1)                                     ' Reset VB random number generator
    Randomize GetTickCount - (GetTickCount \ 4)  ' Reseed VB random number generator
    
    ' Verify receiving file is empty
    hFile = FreeFile                   ' capture first free file handle
    Open strPath For Output As #hFile  ' Create an empty file
    Close #hFile                       ' close file handle

    hFile = FreeFile                                 ' capture first free file handle
    Open strPath For Binary Access Write As #hFile   ' re-open file in binary mode

    '-----------------------------------------------------------------------------
    lngStart = GetTickCount()   ' starting time
    
    ' Generate some random data
    For lngLoop = 1 To 87
        
        Erase alngData()  ' Start with empty arrays
        Erase abytData()
        
        ReDim alngData(KB_32)    ' Size arrays to desired limit
        ReDim abytData(lngByteCnt)
        
        ' Create random long integer numbers
        For lngIdx = 0 To KB_32 - 1
            alngData(lngIdx) = GetRndValue(MIN_LONG, MAX_LONG)  ' Full 32 Bit range
        Next lngIdx
                
        CopyMemory abytData(0), alngData(0), lngByteCnt    ' Convert long array to byte array
        ReDim Preserve abytData(lngByteCnt - 1)            ' resize byte array (0-65535)
        Put #hFile, lngPointer, abytData()                 ' write byte array to output file
        lngPointer = lngPointer + (UBound(alngData) * 4)   ' update pointer within output file
        
        Rnd (-1)                           ' Reset VB random number generator
        Randomize Abs(CDbl(alngData(0)))   ' Reseed VB random number generator
        
    Next lngLoop
    
    Erase alngData()   ' empty temp arrays
    Erase abytData()
    
    lngByteCnt = FILE_SIZE - lngPointer + 1   ' Calc number of bytes left
    ReDim alngData((lngByteCnt \ 4) + 4)      ' Size long integer array
    ReDim abytData(lngByteCnt)                ' Size byte array

    ' Generate last group of long integers
    For lngIdx = 0 To UBound(alngData) - 1
        alngData(lngIdx) = GetRndValue(MIN_LONG, MAX_LONG)  ' Full 32 Bit range
    Next lngIdx

    CopyMemory abytData(0), alngData(0), lngByteCnt       ' Convert long array data to byte array
    ReDim Preserve abytData(lngByteCnt - 1)               ' Resize byte array
    Put #hFile, lngPointer, abytData()                    ' Write byte array to output file

    strElapsed = ElapsedTime(GetTickCount() - lngStart)   ' Finish time
    '-----------------------------------------------------------------------------
    
    Close #hFile      ' close output file
    Erase alngData()  ' Always empty arrays when not needed
    Erase abytData()
    
    Test2 = strElapsed  ' Return elapsed time
    
End Function

Private Static Function GetRndValue(ByVal MIN As Single, _
                                    ByVal MAX As Single) As Long
    
    ' This algorithm written by Alfred Hellmüller
    '
    ' This algorithm was not written to replace the one written
    ' by MS, as shown below, but for creating random data in general.
    '
    ' The data produced by this routine will pass some of the
    ' Diehard tests but not all.
    
    GetRndValue = CLng(Int(Rnd() * (MAX - MIN + 1)) + MIN) Xor _
                  CLng(Int(Rnd() * (512 - -512 + 1)) + -512)

    '======================================================================
    ' The algorithm below was written by MS for producing values
    ' in a given range. This will build a binary file very quickly.
    ' However, Diehard will usually abort before finishing the
    ' first test (Birthday Spacings) and ENT shows very poor
    ' results. This just goes to show that numbers may look
    ' random when in reality they are not truely random.  This
    ' is still a great algorithm for selecting values between two
    ' known values using Visual Basic's RND function.
    '
    ' Visual Basic Language Reference Rnd Function
    ' http://msdn2.microsoft.com/en-us/library/f7s023d2(VS.71).aspx
    '
    ' GetRndValue = CLng(Int(Rnd() * (Max - Min + 1)) + Min)
    '======================================================================
                  
End Function

' ***************************************************************************
' Routine:       ElapsedTime
'
' Description:   Formats time display
'
' Reference:     Karl E. Peterson, http://vb.mvps.org/
'
' Parameters:    lngMilliseconds - Time in milliseconds
'
' Returns:       Formatted output
'                Ex:  12:34:56.789  <- 12 hours 34 minutes 56 seconds 789 thousandths
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function ElapsedTime(ByVal lngMilliseconds As Long) As String

    Dim lngDays As Long
    
    Const ONE_DAY As Long = 86400000   ' Number of milliseconds in a day
    
    ElapsedTime = vbNullString                 ' Verify output string is empty
    lngDays = Fix(lngMilliseconds / ONE_DAY)   ' Calculate number of days
        
    ' See if one or more days has passed
    If lngDays > 0 Then
        ElapsedTime = CStr(lngDays) & " day(s)  "                 ' Start loading output string
        lngMilliseconds = lngMilliseconds - (ONE_DAY * lngDays)   ' Calculate number of milliseconds left
    End If

    ' Continue formatting output string as HH:MM:SS
    ElapsedTime = ElapsedTime & Format$(DateAdd("s", (lngMilliseconds \ 1000), #12:00:00 AM#), "HH:MM:SS")
    lngMilliseconds = lngMilliseconds - ((lngMilliseconds \ 1000) * 1000)   ' Calc number of milliseconds left
    
    ' Append thousandths to output string
    ElapsedTime = ElapsedTime & "." & Format$(lngMilliseconds, "000")
   
End Function

' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Private Function QualifyPath(ByVal strPath As String, _
                    Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If
    
End Function


