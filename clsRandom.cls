VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsRandom - a cryptographic random number generator
'                (Pseudo Random Number Generator using Microsoft's CryptoAPI)
'
' Description:   CryptGenRandom gets its randomness, also known as
'                entropy, from many sources in Windows, including
'                the following:
'
'                - The current process ID (GetCurrentProcessID).
'                - The current thread ID (GetCurrentThreadID).
'                - The ticks since boot (GetTickCount).
'                - The current time (GetLocalTime).
'                - Various high-precision performance counters
'                      (QueryPerformanceCounter).
'                - A Message Digest 4 (MD4) hash of the user’s environment
'                      block, which includes username, computer name, and
'                      search path. MD4 is a hashing algorithm that creates
'                      a 128-bit message digest (16 bytes) from input data
'                      to verify data integrity.
'                - High-precision internal CPU counters, such as RDTSC,
'                      RDMSR, RDPMC.
'                - Low-level system information, such as idle time, kernel
'                      time, interrupt times, commit limit, page read count,
'                      cache read count, nonpaged pool allocations, alignment
'                      fixup count, operating system lookaside information.
'                - [Optional] User defined data as extra seed data. I created
'                      a routine named CreateExtraSeed() to generate a unique
'                      hex data string as my optional data.  A good example
'                      is used in BuildRndData() routine.
'
'                Such information is added to a buffer, which is hashed using
'                MD4 and used as the key to modify the user-provided buffer
'                using RC4.  (Refer to the CryptGenRandom() documentation in
'                the Platform SDK)  The result is a cryptographic random
'                number value.
'
' References:    Randomize Statement Doesn't Re-initialize Rnd Function
'                "To re-initialize the random-number generator, use the Rnd
'                function with a value of -1 and then use the Randomize
'                statement with the value you want to use as the seed value
'                for the Rnd function."
'                http://support.microsoft.com/default.aspx?scid=kb;en-us;120587
'
'                VBA's Pseudo Random Number Generator
'                http://www.noesis.net.au/prng.php
'
'                Visual Basic Language Reference Rnd Function
'                http://msdn2.microsoft.com/en-us/library/f7s023d2(VS.71).aspx
'
'                Mark Hutchinson article:
'                An Examination of Visual Basic's Random Number Generation
'                http://www.15seconds.com/issue/051110.htm
'
'                INFO: How Visual Basic Generates Pseudo-Random Numbers for
'                the RND Function
'                http://support.microsoft.com/kb/231847/en-us
'
'                RND and RANDOMIZE Alternatives for Generating Random Numbers
'                http://support.microsoft.com/kb/28150/EN-US/
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
' 10-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated CreateExtraSeed() in determining the toggle method.
' 25-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote GetRndValue() routine.  Thanks to Alfred Hellmüller
'              for seeing the shortcomings of using the Visual Basic RND()
'              function.
' 03-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Removed some obsolete data.
' 12-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added additional range testing in GetRndValue() routine.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Modified logic for obtaining a valid value in GetRndValue()
'                routine.
'              - Added static variable in CreateExtraSeed() routine to hold
'                a carryover data string.
'              - Fixed bug in BuildWithinRange() routine in loading a hex
'                array and testing for valid return formats.
' 17-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Added EmptyCollection() routine to properly empty a
'                collection object.
'              - Updated CombSort() routine.
'              - Updated documentation in ReshuffleData() and NonRepeatingNbrs()
'                routines.
' 14-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated CreateExtraSeed() and RndSeed() routines.
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected a potential overflow in alternate calculation in
'              CreateExtraSeed() routine.
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for bringing to my attention the
'              need to update GetProviderHandle() routine.
'              - Rewrote GetProviderHandle() routine to test for availability
'                of Advanced Encryption Standard (AES) hash functionality.
'              - Added HashSelection() routine.
'              - Updated CreateHash() routine to access SHA2 hash family.
' 03-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Updated CreateHash() and HashSelection() routines.
'              - Updated documentation in this module and associated text
'                files.
' 08-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote ReshuffleData() and CreateExtraSeed() routines.
'              - Updated CombSort(), RndSeed() and RemoveDupes() routines.
'              - Added ReverseArrayData() routine.  Called by CombSort().
'              - Removed some obsolete code.
' 24-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Modifed and documented ConvertDataToHex(), HashSelection(),
'                CreateHash(), BuildRndData() routines
'              - Added boolean property AES_Ready()
' 10-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Updated selection of data to seed VB random number generator
'                in CreateExtraSeed() routine.
'              - Updated evaluation of input data in NonRepeatingNbrs()
'                routine.
' 02-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Added boolean flag parameter to RndSeed() routine.
'              - Updated CreateExtraSeed() routine.  Added reference
'                to API GetTickCount()
' 18-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote RndSeed() routine.
' 25-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Added property CompareMethod() to determine type of data
'                comparison.
'              - Updated RemoveDupes() routine to use CompareMethod()
'                property.
' 20-Oct-2011  Kenneth Ives kenaso@tx.rr.com
'              Increased maximum number of mixing iterations in ReshuffleData()
'              routine.
' 30-Dec-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated CreateExtraSeed() routine.
'              - Bug fix in GetRndValue() routine.
'              - Added optional parameter, blnCreateExtraSeed, to BuildRndData()
'                and BuildWithinRange() routines.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              Deleted RemoveTrailingNulls() routine from this module. 
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME            As String = "clsRandom"
  Private Const ZEROES_8               As String = "00000000"
  Private Const KB_4                   As Long = &H1000&            '  4096
  Private Const KB_64                  As Long = &H10000            '  65536
  Private Const MAX_INT                As Long = &H7FFF&            '  32767
  Private Const MAX_LONG               As Long = &H7FFFFFFF         '  2147483647
  Private Const MIN_LONG               As Long = &H80000000         ' -2147483648
  Private Const GB_4                   As Double = (2# ^ 32)        '  4294967296  (== 4.2 Gig)
  Private Const MAX_DWORD              As Double = (2# ^ 32) - 1    '  4294967295  (unsigned long int)
  Private Const DBL_LOW                As Double = 0.000000000001
  Private Const DBL_HIGH               As Double = (1.999999999998 / MAX_DWORD)  ' 0.000000000465661287415694
  
  ' Microsoft Hash constants
  Private Const HP_HASHVAL             As Long = 2
  Private Const ALG_CLASS_HASH         As Long = &H8000&     ' Used by all hashes (32768)
  Private Const ALG_TYPE_ANY           As Long = 0           ' Used by all hashes
  Private Const ALG_SID_MD2            As Long = 1           ' 16 byte hashed return length
  Private Const ALG_SID_MD4            As Long = 2           ' 16 byte hashed return length
  Private Const ALG_SID_MD5            As Long = 3           ' 16 byte hashed return length
  Private Const ALG_SID_SHA1           As Long = 4           ' 20 byte hashed return length
  Private Const ALG_SID_SHA_256        As Long = 12          ' 32 byte hashed return length
  Private Const ALG_SID_SHA_384        As Long = 13          ' 48 byte hashed return length
  Private Const ALG_SID_SHA_512        As Long = 14          ' 64 byte hashed return length
  Private Const CALG_MD2               As Long = &H8001&     ' ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2 (32769)
  Private Const CALG_MD4               As Long = &H8002&     ' ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4 (32770)
  Private Const CALG_MD5               As Long = &H8003&     ' ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5 (32771)
  Private Const CALG_SHA1              As Long = &H8004&     ' ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1 (32772)
  Private Const CALG_SHA_256           As Long = &H800C&     ' ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256 (32780)
  Private Const CALG_SHA_384           As Long = &H800D&     ' ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384 (32781)
  Private Const CALG_SHA_512           As Long = &H800E&     ' ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512 (32782)
  ' Microsoft Provider type constants
  Private Const PROV_RSA_FULL          As Long = 1           ' Provider type ID
  Private Const PROV_RSA_AES           As Long = 24          ' AES Provider type ID
  Private Const CRYPT_NEWKEYSET        As Long = &H8         ' For creating a generic provider handle
  Private Const CRYPT_VERIFYCONTEXT    As Long = &HF0000000  ' -268435456
  ' Verify provider names
  ' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography\Defaults\Provider\
  Private Const MS_DEFAULT_PROV        As String = "Microsoft Base Cryptographic Provider v1.0"
  Private Const MS_ENHANCED_PROV       As String = "Microsoft Enhanced Cryptographic Provider v1.0"
  Private Const MS_ENH_RSA_AES_PROV    As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"
  Private Const MS_ENH_RSA_AES_PROV_XP As String = "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"

' ***************************************************************************
' Type Structures
' ***************************************************************************
  Private Type CARRYOVER_DATA
      Value1 As Long
      Value2 As Long
  End Type

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumPRNG_ReturnFormat
      ePRNG_ASCII           ' 0
      ePRNG_HEX             ' 1
      ePRNG_HEX_ARRAY       ' 2
      ePRNG_BYTE_ARRAY      ' 3
      ePRNG_LONG_ARRAY      ' 4
      ePRNG_DBL_ARRAY       ' 5
  End Enum

  Public Enum enumPRNG_HashAlgorithm
      ePRNG_MD2             ' 0
      ePRNG_MD4             ' 1
      ePRNG_MD5             ' 2
      ePRNG_SHA1            ' 3
      ePRNG_SHA256          ' 4
      ePRNG_SHA384          ' 5
      ePRNG_SHA512          ' 6
  End Enum

  Public Enum enumPRNG_Compare
      ePRNG_CaseSensitive   ' 0 - vbBinaryCompare
      ePRNG_IgnoreCase      ' 1 - vbTextCompare
  End Enum

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The QueryPerformanceCounter function retrieves the current value of the
  ' high-resolution performance counter, if one exists.
  Private Declare Function QueryPerformanceCounter Lib "Kernel32" _
          (curCounter As Currency) As Boolean

  ' The QueryPerformanceFrequency function retrieves the frequency of the
  ' high-resolution performance counter, if one exists. The frequency
  ' cannot change while the system is running.
  Private Declare Function QueryPerformanceFrequency Lib "Kernel32" _
          (curFrequency As Currency) As Boolean

  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Private Declare Function GetTickCount Lib "Kernel32" () As Long
  
  ' The CryptCreateHash function initiates the hashing of a stream of
  ' data. It creates and returns to the calling application a handle
  ' to a CSP hash object. This handle is used in subsequent calls to
  ' CryptHashData and CryptHashSessionKey to hash session keys and
  ' other streams of data.
  Private Declare Function CryptCreateHash Lib "advapi32.dll" _
          (ByVal hProv As Long, ByVal algid As Long, _
          ByVal hkey As Long, ByVal dwFlags As Long, _
          ByRef phHash As Long) As Long

  ' The CryptHashData function adds data to a specified hash object.
  ' This function and CryptHashSessionKey can be called multiple
  ' times to compute the hash of long or discontinuous data streams.
  Private Declare Function CryptHashData Lib "advapi32.dll" _
          (ByVal hhash As Long, ByVal pbData As String, _
          ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

  ' The CryptGetHashParam function retrieves data that governs the
  ' operations of a hash object. The actual hash value can be
  ' retrieved by using this function.
  Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
          (ByVal hhash As Long, ByVal dwParam As Long, _
          ByVal pbData As String, pdwDataLen As Long, _
          ByVal dwFlags As Long) As Long

  'The CryptDestroyHash function destroys the hash object referenced
  ' by the hHash parameter. After a hash object has been destroyed,
  ' it can no longer be used.  The destruction of hash objects after
  ' their use is finished is recommended for sedblity reasons.
  Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
          (ByVal hhash As Long) As Long

  ' The CryptAcquireContext function is used to acquire a handle to a
  ' particular key container within a particular cryptographic service
  ' provider (CSP). This returned handle can then be used to make
  ' calls to the selected CSP.  This function performs two operations.
  ' It first attempts to find a CSP with the characteristics described
  ' in the dwProvType and pszProvider parameters. If the CSP is found,
  ' the function attempts to find a key container within the CSP
  ' matching the name specified by the pszContainer parameter. With the
  ' appropriate setting of dwFlags, this function can also create and
  ' destroy key containers. If function succeeds, return value is nonzero.
  Private Declare Function CryptAcquireContext Lib "advapi32.dll" _
          Alias "CryptAcquireContextA" (ByRef phProv As Long, _
          ByVal pszContainer As String, ByVal pszProvider As String, _
          ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

  ' The CryptReleaseContext function releases the handle of a
  ' cryptographic service provider (CSP) and a key container. At each
  ' call to this function, the reference count on the CSP is reduced
  ' by one. When the reference count reaches zero, the context is fully
  ' released and it can no longer be used by any function in the application.
  ' An application calls this function after finishing the use of the CSP.
  ' After this function is called, the released CSP handle is no longer
  ' valid. This function does not destroy key containers or key pairs.
  Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
          (ByVal hProv As Long, ByVal dwFlags As Long) As Long

  ' The data produced by this function is cryptographically random.  The
  ' data is far more random than the data generated by the typical random
  ' number generator such as the one shipped with your C or VB compiler.
  ' If the function fails, the return value is zero.
  Private Declare Function CryptGenRandom Lib "advapi32.dll" _
          (ByVal hProv As Long, ByVal dwLen As Long, _
          ByVal pbBuffer As String) As Long

  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' ZeroMemory is used for clearing the contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long)
'                    |  |     |----- Variable subname
'                    - --- --------------
' Naming standard:   m lng ProviderHandle
' Variable name:     mlngProviderHandle
'
' ***************************************************************************
  Private mcurSeed           As Currency
  Private mcurCounter        As Currency
  Private mcurFrequency      As Currency
  Private mlngProviderHandle As Long              ' CryptoAPI provider handle
  Private mblnAESProvider    As Boolean           ' Flag denoting AES strong provider
  Private mlngCompareMethod  As enumPRNG_Compare  ' Comparison method
  Private mtypCarryOver      As CARRYOVER_DATA    ' Carryover values

' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

'Public Property Get Version() As String
'    Version = DLL_NAME & ".dll (tm) v" & CStr(App.Major) & "." & _
'              CStr(App.Minor) & "." & CStr(App.Revision) & _
'              " Copyright (c) 2004-2012  Kenneth Ives  kenaso@tx.rr.com"
'End Property
 
Public Property Get AES_Ready() As Boolean
    
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.AES_Ready
    
    ' Test for a provider handle
    If mlngProviderHandle < 1 Then
        GetProviderHandle        ' Get provider handle
    End If
    
    AES_Ready = mblnAESProvider  ' Is advanced hashing available?
    
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property

Public Property Let CompareMethod(ByVal lngCompareMethod As enumPRNG_Compare)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.   Syntax: X.CompareMethod = ePRNG_CaseSensitive
    mlngCompareMethod = lngCompareMethod
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       BuildRndData
'
' Description:   Create random generated data based on a range (0-255)
'
' Parameters:    lngDataLength - Length of data string
'                lngReturnFormat - OPTIONAL - Return data in a type of
'                    format.  Default - ePRNG_BYTE_ARRAY
'                blnCreateExtraSeed - OPTIONAL - Flag designates if extra
'                    seed values should be created.  Default - TRUE
'
' Returns:       Specific length of random generated data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 04-SEP-2006  Kenneth Ives  kenaso@tx.rr.com
'              Added sixteen extra bytes when generating data to allow for
'              conversion.
' 30-Dec-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added optional parameter so user can create an optional
'              seed when generating random data.
' ***************************************************************************
Public Function BuildRndData(ByVal lngDataLength As Long, _
                    Optional ByVal lngReturnFormat As enumPRNG_ReturnFormat = ePRNG_BYTE_ARRAY, _
                    Optional ByVal blnCreateExtraSeed As Boolean = True) As Variant

    Dim lngIdx     As Long     ' array index
    Dim lngIndex   As Long
    Dim lngValue   As Long     ' temp hold area
    Dim lngMaxSize As Long     ' max array size
    Dim alngData() As Long
    Dim abytRnd()  As Byte
    Dim abytData() As Byte
    Dim adblData() As Double
    Dim strRndData As String   ' output random data string
    Dim astrData() As String
    
    Const ROUTINE_NAME As String = "BuildRndData"
    Const EXTRA_BYTES  As Long = 16
    
    On Error GoTo BuildRndData_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' A return length greater than 0 is required
    If lngDataLength < 1 Then
        InfoMsg "Return data length must be greater than zero." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    Erase abytRnd()  ' Always start with empty arrays
    Erase abytData()
    Erase alngData()
    Erase adblData()
    Erase astrData()
    
    BuildRndData = Empty           ' Empty return value
    ReDim avntData(lngDataLength)  ' Size work array

    Select Case lngReturnFormat
           
           Case ePRNG_ASCII, ePRNG_HEX, ePRNG_HEX_ARRAY, ePRNG_BYTE_ARRAY
                lngMaxSize = lngDataLength + EXTRA_BYTES
                
           Case ePRNG_LONG_ARRAY, ePRNG_DBL_ARRAY
                ' 4 bytes = 1 long Integer
                lngMaxSize = (lngDataLength * 4&) + EXTRA_BYTES
    End Select

    ' If no provider handle then obtain
    ' one for this version of Windows
    If mlngProviderHandle < 1 Then
        GetProviderHandle
    End If

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo BuildRndData_CleanUp
    End If

    If blnCreateExtraSeed Then
        ' Create hex data string (10-128 characters)
        ' using an optional seed with an appended
        ' string of blank spaces equal to requested
        ' return length.
        lngValue = Val(Int(Rnd() * (128 - 10 + 1)) + 10)
        strRndData = CreateExtraSeed(lngValue) & Space$(lngMaxSize)
    Else
        ' String of blank spaces equal to requested
        ' return length
        strRndData = Space$(lngMaxSize)
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo BuildRndData_CleanUp
    End If

    ' Make API call to generate random data string
    If CryptGenRandom(mlngProviderHandle, Len(strRndData), strRndData) = 0 Then
        gblnStopProcessing = True  ' An error occurred
    Else
        ' Random data will always begin in the
        ' left-most position of returned string
        strRndData = Left$(strRndData, lngMaxSize)  ' Capture random data
    
        Select Case lngReturnFormat
               
               Case ePRNG_ASCII
                    strRndData = Left$(strRndData, lngDataLength)
                    BuildRndData = strRndData
    
               Case ePRNG_HEX
                    strRndData = ConvertDataToHex(lngDataLength, strRndData, True)
                    BuildRndData = strRndData
    
               Case ePRNG_HEX_ARRAY
                    astrData() = ConvertDataToHex(lngDataLength, strRndData, False)
                    ReDim Preserve astrData(lngDataLength)
                    BuildRndData = astrData()
    
               Case ePRNG_BYTE_ARRAY
                    abytData() = StringToByteArray(strRndData)
                    ReDim Preserve abytData(lngDataLength)
                    BuildRndData = abytData()
    
               Case ePRNG_LONG_ARRAY
                    ' return a long integer value of
                    ' -2147483648 to 2147483647
                    
                    abytRnd() = StringToByteArray(strRndData)  ' Convert string data to byte array
                    lngMaxSize = UBound(abytRnd)               ' capture array size
                    lngValue = lngMaxSize Mod 4                ' must be in multiples of 4
                    
                    ' If not a multiple of 4 then
                    ' resize the array accordingly
                    If lngValue <> 0 Then
                        ReDim Preserve abytRnd(lngMaxSize - lngValue)
                    End If
                    
                    ReDim alngData(lngDataLength)  ' resize output array
                    lngIdx = 0
    
                    For lngIndex = 0 To (UBound(abytRnd) - 1) Step 4
    
                        CopyMemory lngValue, abytRnd(lngIndex), 4&   ' create 1 long integer
                        alngData(lngIdx) = lngValue                  ' Save to output array
                        lngIdx = lngIdx + 1                          ' increment index
    
                        If lngIdx = lngDataLength Then
                            Exit For    ' exit For..Next loop
                        End If
    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
    
                    ReDim Preserve alngData(lngDataLength)
                    BuildRndData = alngData()
    
               Case ePRNG_DBL_ARRAY
                    ' Return a decimal position of 14 places.
                    ' This was chosen so that exponentials
                    ' would not be returned.
                    ' Ex: -6.62493605785922E-02
                    '
                    ' These are the type of values desired
                    ' -0.99999999999999 to 0.99999999999999
                    
                    abytRnd() = StringToByteArray(strRndData)   ' Convert string data to byte array
                    lngMaxSize = UBound(abytRnd)                ' capture array size
                    lngValue = lngMaxSize Mod 4                 ' must be in multiples of 4
                    lngIdx = 0
                    
                    ' If not a multiple of 4 then
                    ' resize the array accordingly
                    If lngValue <> 0 Then
                        ReDim Preserve abytRnd(lngMaxSize - lngValue)
                    End If
                    
                    ReDim adblData(lngDataLength)  ' resize output array
                    
                    For lngIndex = 0 To (UBound(abytRnd) - 1) Step 4
    
                        CopyMemory lngValue, abytRnd(lngIndex), 4&   ' create 1 long integer
                        adblData(lngIdx) = LongToDouble(lngValue)    ' convert to a double
                        lngIdx = lngIdx + 1                          ' increment index
    
                        If lngIdx = lngDataLength Then
                            Exit For    ' exit For..Next loop
                        End If
    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
    
                    ReDim Preserve adblData(lngDataLength)
                    BuildRndData = adblData()
    
               Case Else  ' return string of generated data
                    strRndData = Left$(strRndData, lngDataLength)
                    BuildRndData = strRndData
        End Select
    End If
        
BuildRndData_CleanUp:
    Erase abytRnd()   ' Always empty arrays when not needed
    Erase abytData()
    Erase alngData()
    Erase adblData()
    Erase astrData()
    strRndData = vbNullString
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

BuildRndData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume BuildRndData_CleanUp

End Function

' ***************************************************************************
' Routine:       BuildWithinRange
'
' Description:   Create random generated data based on a range (0-255)
'
' Parameters:    lngDataLength - Length of data string
'                lngLow - [OPTIONAL] lowest value for the return string
'                    DEFAULT = 0
'                lngHigh - [OPTIONAL] highest value for the return string
'                    Default = 255
'                lngReturnFormat - [OPTIONAL] Return data in a type of
'                    format.  Default - ePRNG_BYTE_ARRAY
'                blnCreateExtraSeed - OPTIONAL - Flag designates if extra
'                    seed values should be created.  Default - TRUE
'
' Returns:       Random generated data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 04-SEP-2006  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in the logic of collecting the data and added
'              additional error checking.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in loading return hex array.  Updated return
'              format test.
' 30-Dec-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added optional parameter so user can create an optional
'              seed when generating random data.
' ***************************************************************************
Public Function BuildWithinRange(ByVal lngDataLength As Long, _
                        Optional ByVal lngLow As Long = 0, _
                        Optional ByVal lngHigh As Long = 255, _
                        Optional ByVal lngReturnFormat As enumPRNG_ReturnFormat = ePRNG_BYTE_ARRAY, _
                        Optional ByVal blnCreateExtraSeed As Boolean = True) As Variant

    Dim lngIdx     As Long
    Dim lngMax     As Long
    Dim lngIndex   As Long
    Dim strRndData As String
    Dim astrData() As String
    Dim abytData() As Byte
    Dim abytRnd()  As Byte

    Const ROUTINE_NAME  As String = "BuildWithinRange"

    On Error GoTo BuildWithinRange_Error

    BuildWithinRange = Empty   ' Empty return value
    
    '----------------------------------------------------- Evaluate incoming data
    ' Check data return format
    Select Case lngReturnFormat
           Case ePRNG_LONG_ARRAY, ePRNG_DBL_ARRAY
                InfoMsg "Long Integer and Double precision arrays " & vbNewLine & _
                        "are not valid return formats." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                GoTo BuildWithinRange_CleanUp
    End Select
    
    ' If using full ASCII range (0-255) then
    ' there is no reason to go any further
    If lngLow = 0 And lngHigh = 255 Then
        BuildWithinRange = BuildRndData(lngDataLength, lngReturnFormat)
        GoTo BuildWithinRange_CleanUp
    End If

    ' A length greater than 0 is required
    If lngDataLength < 1 Then
        InfoMsg "The length of the return data must be" & vbNewLine & _
                "greater than zero." & vbNewLine & vbNewLine & _
                "Return length - " & CStr(lngDataLength) & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    ' check range limits
    Select Case lngLow
           Case Is < 0
                InfoMsg "The minimum range value must be a positive" & vbNewLine & _
                        "value of 0-254." & vbNewLine & vbNewLine & _
                        "Current Min - " & CStr(lngLow) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function

           Case Is > 254
                InfoMsg "The minimum range value must be a positive" & vbNewLine & _
                        "value of 0-254." & vbNewLine & vbNewLine & _
                        "Current Min - " & CStr(lngLow) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function

           Case Is >= lngHigh
                InfoMsg "The minimum range value must be less than" & vbNewLine & _
                        "the maximum range value." & vbNewLine & vbNewLine & _
                        "Current Min - " & CStr(lngLow) & vbNewLine & _
                        "Current Max - " & CStr(lngHigh) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function
    End Select

    Select Case lngHigh
           Case Is < 1
                InfoMsg "The maximum range value must be a positive" & vbNewLine & _
                        "value of 1-255." & vbNewLine & vbNewLine & _
                        "Current Max - " & CStr(lngHigh) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function

           Case Is > 255
                InfoMsg "The maximum range value must be a positive" & vbNewLine & _
                        "value of 1-255." & vbNewLine & vbNewLine & _
                        "Current Max - " & CStr(lngHigh) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function
    End Select
    '----------------------------------------------------- End of evaluation

    lngIdx = 0
    lngIndex = MAX_INT           ' preset to cause get rnd data
    lngMax = lngDataLength + 4   ' add buffer of extra chars
    ReDim abytData(lngMax)       ' Resize temp work array

    ' Build string of random data
    Do
        If lngIndex >= (KB_4 - 1) Then
            
            lngIndex = 0                                          ' reset pointer
            Erase abytRnd()                                       ' empty temp array
            abytRnd() = BuildRndData(KB_4, , blnCreateExtraSeed)  ' create some random data
        
            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit Do  ' exit Do..Loop
            End If
        End If

        ' does current value fall within requested range
        Select Case abytRnd(lngIndex)
               
               Case lngLow To lngHigh   ' valid data
                    abytData(lngIdx) = abytRnd(lngIndex)
                    lngIdx = lngIdx + 1

                    ' Has enough data has been collected?
                    If lngIdx = lngMax Then
                        Exit Do  ' exit Do..Loop
                    End If
        End Select
    
        lngIndex = lngIndex + 1  ' Increment index pointer
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit Do  ' exit Do..Loop
        End If

    Loop

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo BuildWithinRange_CleanUp
    End If

    Select Case lngReturnFormat

           Case ePRNG_ASCII      ' return as raw string data
                strRndData = ByteArrayToString(abytData())
                BuildWithinRange = Left$(strRndData, lngDataLength)

           Case ePRNG_HEX        ' return as hex string
                strRndData = ByteArrayToString(abytData())
                strRndData = ConvertDataToHex(lngDataLength, strRndData, True)
                BuildWithinRange = Left$(strRndData, lngDataLength)

           Case ePRNG_HEX_ARRAY  ' return as hex array
                strRndData = ByteArrayToString(abytData())
                astrData() = ConvertDataToHex(lngDataLength, strRndData, False)
                ReDim Preserve astrData(lngDataLength)
                BuildWithinRange = astrData()

           Case ePRNG_BYTE_ARRAY  ' return as byte array
                ReDim Preserve abytData(lngDataLength)
                BuildWithinRange = abytData()

           Case Else  ' return as byte array
                ReDim Preserve abytData(lngDataLength)
                BuildWithinRange = abytData()
    End Select

BuildWithinRange_CleanUp:
    Erase abytRnd()   ' Always empty arrays when not needed
    Erase abytData()
    Erase astrData()
    strRndData = vbNullString
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

BuildWithinRange_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume BuildWithinRange_CleanUp

End Function

' ***************************************************************************
' Procedure:     RemoveDupes
'
' Description:   First the data will be sorted in ascending order.  This
'                routine removes all duplicates based on user selection
'                of case sensitivity.  Number of duplicates removed are
'                returned.
'
'                Example:  RemoveDupes abytTemp(), lngDupes
'
'                          Remove duplicates from array comparing the data
'                          as case sensitive and return number of duplicates
'                          that were in the array.
'
'                                 Raw        Sorted
'                                 Data     (no dupes)
'                              --------     --------
'                                 105         100
'                                 107         101
'                                 101         105
'                                 107         106
'                                 106         107
'                                 109         109
'                                 106         110
'                                 110
'                                 100      lngDupeCnt = 3
'                                 106
'
' Parameters:    avntData - an array of data to be evaluated
'                lngDupeCnt - [OPTIONAL] Returns the number of duplicates
'                     removed. User can access this variable if desired.
'
' Returns:       An array of data with no duplicate values.  If requested,
'                number of items removed from original array.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 07-JAN-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 15-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug that dropped trailing values
' 08-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Force all incoming data to go thru sorting process
' 25-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated to use CompareMethod() property
' ***************************************************************************
Public Function RemoveDupes(ByRef avntData As Variant, _
                   Optional ByRef lngDupeCnt As Long = 0) As Boolean

    Dim lngIdx      As Long
    Dim lngLow      As Long    ' lowest number of elements in the array
    Dim lngHigh     As Long    ' highest number of elements in the array
    Dim lngItem     As Long
    Dim lngIndex    As Long
    Dim strCurrItem As String
    Dim strPrevItem As String
    Dim avntTemp    As Variant

    Const ROUTINE_NAME As String = "RemoveDupes"

    On Error GoTo RemoveDupes_Error

    avntTemp = Empty     ' Always start with empty variants
    lngIdx = 0
    lngDupeCnt = 0
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Test incoming data
    If Not CBool(IsArrayInitialized(avntData)) Then
        InfoMsg "Incoming array has not been properly initialized." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    lngLow = LBound(avntData)
    lngHigh = UBound(avntData)
    ReDim avntTemp(lngLow To lngHigh)
    
    ' Copy data to temp array
    For lngIndex = lngLow To lngHigh - 1
        avntTemp(lngIndex) = avntData(lngIndex)
    Next lngIndex
        
    CombSort avntTemp    ' Sort data in ascending order
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo RemoveDupes_CleanUp
    End If

    ' loop backwards thru data array and remove
    ' any duplicates by emptying array element
    For lngItem = (lngHigh - 1) To (lngLow + 1) Step -1
        
        strCurrItem = Trim$(CStr(avntTemp(lngItem)))
        strPrevItem = Trim$(CStr(avntTemp(lngItem - 1)))

        If StrComp(strCurrItem, strPrevItem, mlngCompareMethod) = 0 Then
            avntTemp(lngItem) = vbNullString   ' Empty array element
            lngDupeCnt = lngDupeCnt + 1        ' Increment duplicate counter
        End If
    
    Next lngItem
    
    Erase avntData                     ' Empty return array
    ReDim avntData(lngLow To lngHigh)  ' resize return data array to max
    
    ' reload return array
    For lngIndex = lngLow To lngHigh - 1
        
        ' If there is data in array element
        ' then copy data to return array
        If Len(Trim$(avntTemp(lngIndex))) > 0 Then
            avntData(lngIdx) = avntTemp(lngIndex)
            lngIdx = lngIdx + 1
        End If

    Next lngIndex
    
    ' Resize original array to actual
    ' number of elements to be returned
    ReDim Preserve avntData(lngLow To lngIdx)

RemoveDupes_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        RemoveDupes = False  ' Set flag for bad finish
    Else
        RemoveDupes = True   ' Set flag for successful finish
    End If

    avntTemp = Empty  ' Always empty variants when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

RemoveDupes_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume RemoveDupes_CleanUp
    
End Function

' ***************************************************************************
' Routine:       ReshuffleData
'
' Description:   An array of data passed to this routine will be rearranged.
'                Incoming array must have a minimum of 3 elements. The array
'                must be zero based.  During the shuffling process, no items
'                will be duplicated or removed.  This is a short coming in
'                some reshuffling routines.
'
' Parameters:    avntData - Array of data to be mixed (Unknown data type)
'                lngMixCount - Optional - Number of mixing iterations to 
'                           rearrange array data.  Default = 25
'
' Returns:       A mixed array of data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-May-2010  Kenneth Ives kenaso@tx.rr.com
'              Rewrote routine.
' 20-Oct-2011  Kenneth Ives kenaso@tx.rr.com
'              - Increased maximum number of mixing iterations
'              - Updated documentation
' ***************************************************************************
Public Sub ReshuffleData(ByRef avntData As Variant, _
                Optional ByVal lngMixCount As Long = 25)

    Dim lngMin    As Long      ' smallest available subscript of an array
    Dim lngMax    As Long      ' largest available subscript of an array
    Dim lngLoop   As Long      ' Loop counter
    Dim lngIndex  As Long      ' Loop counter
    Dim lngTmpIdx As Long      ' Calculated index for swapping
    Dim lngIdxPtr As Long      ' Reseed index pointer
    Dim blnReseed As Boolean   ' Flag to designate additional seeding
    
    Const MIN          As Long = 25
    Const MAX          As Long = 200
    Const ROUTINE_NAME As String = "ReshuffleData"

    On Error GoTo ReshuffleData_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
        
    '----------------------------------------------------- Evaluate incoming data
    ' Verify number of looping actions
    ' is within an acceptable range
    Select Case lngMixCount
           Case Is < MIN: lngMixCount = MIN   ' Set to minimum
           Case Is > MAX: lngMixCount = MAX   ' Set to maximum
    End Select
    
    ' Test incoming array
    If CBool(IsArrayInitialized(avntData)) Then
        lngMin = LBound(avntData)      ' Lowest index value
        lngMax = UBound(avntData) - 1  ' Highest index value
    Else
        InfoMsg "Incoming array has not been properly prepared." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Sub
    End If
    
    ' Verfiy number of array elements
    If (lngMax < 2) Then
        InfoMsg "Invalid amount of data in array." & _
                vbNewLine & "Element count:  " & CStr(lngMax) & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME

        gblnStopProcessing = True
        Exit Sub
    End If
    '----------------------------------------------------- End of evaluation
    
    RndSeed            ' Reseed VB random number generator
    blnReseed = False  ' Preset reseed flag to FALSE
    lngIdxPtr = 0      ' Preset loop indicator to zero
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
        
    ' Determine if additional reseeding of VB random number
    ' generator will be needed. Read article by Mark Hutchinson.
    ' An Examination of Visual Basic's Random Number Generation
    ' http://www.15seconds.com/issue/051110.htm
    If (lngMixCount * lngMax) > KB_64 Then
        
        For lngIndex = 2 To lngMixCount
            
            ' Is this the loop that exceeds 65,536 bytes
            If (lngIndex * lngMax) > KB_64 Then
                blnReseed = True           ' Set flag to TRUE
                lngIdxPtr = lngIndex - 1   ' Save previous index pointer
                Exit For                   ' exit For..Next loop
            End If
        Next lngIndex
    End If
    
    ' Rearrange data elements within an array.  The more elegant
    ' and faster shuffling algorithm is also known as the "Knuth
    ' Shuffle", popularized by Donald Knuth, in his book,
    ' "The Art of Computer Programming".
    '
    ' "Knuth Shuffle" randomly picks an element from the array
    ' and swaps it with the current element.  The maximum count
    ' is then decremented.  Fourteen or more loops will create
    ' a good mix of data.  Default here is twenty-five loops.
    '
    ' Shuffling - Shuffle a deck of cards - Knuth Shuffle
    ' http://tekpool.wordpress.com/2006/10/06/shuffling-shuffle-a-deck-of-cards-knuth-shuffle/
    For lngLoop = 1 To lngMixCount
        
        ' Step backwards thru array and swap data
        For lngIndex = lngMax To lngMin Step -1
        
            ' Create random index
            lngTmpIdx = Val(Int(Rnd() * (lngIndex - lngMin + 1)) + lngMin)
            
            ' Swap data if not same index pointer
            If lngIndex <> lngTmpIdx Then
                SwapData avntData(lngIndex), avntData(lngTmpIdx)
            End If
            
        Next lngIndex
            
        ' Will there be multiple reseeding?
        If blnReseed Then
            
            ' Use modulus arithmetic to determine if it
            ' is time to reseed VB random number generator
            If lngLoop Mod lngIdxPtr = 0 Then
                RndSeed
            End If
        End If
                     
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For   ' exit For..Next loop
        End If

    Next lngLoop
    
ReshuffleData_CleanUp:
    On Error GoTo 0    ' Nullify this error trap
    Exit Sub

ReshuffleData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ReshuffleData_CleanUp

End Sub

' ***************************************************************************
' Routine:       NonRepeatingNbrs
'
' Description:   With this routine you can generate a series of
'                non-repeating numbers.  An array will be loaded starting
'                with the base number (lngMinValue) requested up to the
'                maximum value requested (lngMaxValue).  You can also enter
'                the incremental step between the minimum and maximum value.
'                This data is then passed to another routine ReshuffleData()
'                to be throughly rearranged.  When it is returned, the
'                requested number of elements (lngReturnQty) from the mixed
'                data are transferred sequentially to the return array
'                (alngMixed()).
'
' Parameters:    lngReturnQty - Total count of numbers to be returned
'                lngMinValue - Lowest value in selection array
'                lngMaxValue - Highest value in selection array
'                lngStep - [OPTIONAL] Increment the values by this positive
'                          amount.  Default=1
'                blnSortData - [OPTIONAL] Flag to sort the returned data.
'                          True - Sort data before it is returned
'                          False - Do not sort data
'                          Default=TRUE
'
' Syntax:        Lottery example.
'                    Return 6 numbers
'                    Lowest possible value = 1
'                    Highest possible value = 54
'                    Incremental step = 1 (Default)
'                    Sort in Ascending order (Default)
'
'                    x() = NonRepeatingNbrs(6, 1, 54)
'
' Returns:       An array of random generated numbers.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-MAY-2001  Kenneth Ives kenaso@tx.rr.com
'              Created routine
' 22-MAY-2001  Kenneth Ives kenaso@tx.rr.com
'              Corrected logic in creation of output array.
' 01-JUN-2001  Kenneth Ives kenaso@tx.rr.com
'              Corrected logic in creation of input array.
' 05-MAR-2002  Kenneth Ives kenaso@tx.rr.com
'              Inserted error trapping and redisgned the way the arrays are
'              sized and filled.
' 08-JAN-2006  Kenneth Ives kenaso@tx.rr.com
'              Added lngStep parameter to control the incremental steps between
'              the min and max values.
' 12-MAR-2006  Kenneth Ives kenaso@tx.rr.com
'              Fixed a bug in mixing
' 27-Aug-2008  Kenneth Ives kenaso@tx.rr.com
'              Fixed a bug in range testing
' 17-Mar-2009  Kenneth Ives kenaso@tx.rr.com
'              Reformatted code and added documentation.
' 10-Dec-2010  Kenneth Ives kenaso@tx.rr.com
'              Updated evaluation of input data
' 13-Sep-2010  Kenneth Ives kenaso@tx.rr.com
'              Randomly determing number of mixing iterations
' 13-Sep-2011  Kenneth Ives kenaso@tx.rr.com
'              Formatted meaningful error message
' ***************************************************************************
Public Function NonRepeatingNbrs(ByVal lngReturnQty As Long, _
                                 ByVal lngMinValue As Long, _
                                 ByVal lngMaxValue As Long, _
                        Optional ByVal lngStep As Long = 1, _
                        Optional ByVal blnSortData As Boolean = True) As Long()
    
    Dim strMsg          As String
    Dim strMsgFmt       As String
    Dim lngIdx          As Long
    Dim lngIndex        As Long
    Dim lngValue        As Long
    Dim lngMixCnt       As Long
    Dim lngItemCnt      As Long
    Dim alngMixed()     As Long
    Dim alngSelection() As Long
        
    Const ROUTINE_NAME  As String = "NonRepeatingNbrs"
    Const MIN           As Long = 25
    Const MAX           As Long = 200
    
    On Error GoTo NonRepeatingNbrs_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    lngIdx = 0
    lngItemCnt = 0
    strMsgFmt = "!" & String$(65, "@")   ' left justified string
    
    strMsg = Format$("Returned items:  " & Format$(lngReturnQty, "#,##0"), strMsgFmt) & vbNewLine & _
             Format$("Minimum value:   " & Format$(lngMinValue, "#,##0"), strMsgFmt) & vbNewLine & _
             Format$("Maximum value:   " & Format$(lngMaxValue, "#,##0"), strMsgFmt)
    
    '----------------------------------------------------- Evaluate incoming data
    ' Reset to default if not within acceptable range
    If (lngStep <= 0) Or _
       (lngStep >= lngMaxValue) Then
       
        lngStep = 1
    End If
    
    If lngReturnQty <= 0 Then
        InfoMsg "Need to request at least one value to be returned." & _
                vbNewLine & vbNewLine & strMsg & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    If lngMinValue >= lngMaxValue Then
        InfoMsg "Minimum value cannot exceed or equal maximum value." & _
                vbNewLine & vbNewLine & strMsg & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    ' Count number of items to be
    ' put into the selection array
    For lngIndex = lngMinValue To lngMaxValue Step lngStep
        lngItemCnt = lngItemCnt + 1
    Next lngIndex
    
    If lngReturnQty > lngItemCnt Then
        InfoMsg "Return quantity exceeds number of items produced." & _
                vbNewLine & vbNewLine & strMsg & _
                vbNewLine & Format$("Items produced:  " & Format$(lngItemCnt, "#,##0"), strMsgFmt) & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    strMsgFmt = vbNullString   ' empty message strings
    strMsg = vbNullString
    '----------------------------------------------------- End of evaluation
    
    ' Resize arrays based on values
    ' passed to this routine.
    ReDim alngSelection(lngItemCnt)   ' number of elements to pick from
    ReDim alngMixed(lngReturnQty)     ' number of elements to be returned
    lngValue = lngMinValue
    
    ' Preload selection array.
    ' These are the values to choose from.
    For lngIndex = 0 To (lngItemCnt - 1)
        alngSelection(lngIndex) = lngValue
        lngValue = lngValue + lngStep
    Next lngIndex
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo NonRepeatingNbrs_CleanUp
    End If

    ' Mix numeric data array
    lngMixCnt = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Calc number of mixing loops
    ReshuffleData alngSelection(), lngMixCnt              ' Mix data array
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo NonRepeatingNbrs_CleanUp
    End If
    
    ' Perform verification checks prior to
    ' transfering data from array that was
    ' just mixed into return array
    For lngIndex = 0 To (lngItemCnt - 1)
    
        ' Current element must be within MIN and MAX range
        If (alngSelection(lngIndex) >= lngMinValue) And _
           (alngSelection(lngIndex) <= lngMaxValue) Then
                
            ' Copy data to return array
            alngMixed(lngIdx) = alngSelection(lngIndex)
            lngIdx = lngIdx + 1
                
        End If
        
        ' Has enough data been collected?
        If lngIdx = lngReturnQty Then
            Exit For    ' exit For..Next loop
        End If
        
    Next lngIndex
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo NonRepeatingNbrs_CleanUp
    End If
    
    ' Verify return quantity
    ReDim Preserve alngMixed(lngReturnQty)
    
    ' If more than one number is selected and
    ' sorted data is requested, then sort data
    If lngReturnQty > 1 And blnSortData Then
        CombSort alngMixed()
    End If
  
NonRepeatingNbrs_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim alngMixed(0)   ' Return one zero
    End If

    NonRepeatingNbrs = alngMixed()
    
    Erase alngMixed()     ' Always empty arrays when not needed
    Erase alngSelection()

    On Error GoTo 0       ' Nullify this error trap
    Exit Function

NonRepeatingNbrs_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume NonRepeatingNbrs_CleanUp
    
End Function

' ***************************************************************************
' Procedure:     CombSort
'
' Description:   CombSort is faster than all but QuickSort and close to
'                it.  On the other hand, the code is much simpler than
'                QuickSort and can be easily customized for any array type.
'                The CombSort was first published by Richard Box and Stephen
'                Lacey in the April 1991 issue of Byte magazine. They found
'                that using a sequence for the gaps that decreased by a shrink
'                factor of 1.3 gave the best results. Trial and error produced
'                the refinement that an eventual gap size of 11, (rather than
'                9 or 10) gave a more efficient sequence of gaps below that
'                point.
'
'                Every programmer knows that QuickSort is the best in-memory
'                sort in the world. Look inside the sort function of almost
'                any library and you'll find QuickSort. Another thing that
'                every programmer knows is that BubbleSort, the "hello world"
'                of sorts, is one of the slowest in the world. But what most
'                programmers don't know is that a simple modification to
'                bubble sort turns it into CombSort, a remarkably simple sort
'                that's nearly as fast as QuickSort.
'
'                CombSort starts out comparing items that are far apart. Then
'                it makes the gap smaller and does it again. In the
'                algorithm's last passes the gap is 1, making it act
'                identical to BubbleSort. That makes it easy to see that this
'                algorithm is correct, since we know that BubbleSort is
'                correct and this algorithm always turns into BubbleSort.
'
' Parameters:    avntData - An array of unsorted data (byte, long, string)
'                blnAscending - Optional - Flag denoting if sorting is to be
'                     in ascending or descending order.
'                     Default = Ascending
'
' Returns:       A sorted array is TRUE if the sort was succesful.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 08-JAN-2006  Kenneth Ives kenaso@tx.rr.com
'              Added error trapping
' 10-Feb-2008  Kenneth Ives kenaso@tx.rr.com
'              Made an assumption of the swapped flag being initialized to
'              FALSE.  Now, I manually set the flag prior to sorting.
' 01-Dec-2008  Kenneth Ives kenaso@tx.rr.com
'              Preventing a possible memory leak by emptying a variant when
'              it is not needed.
' 08-May-2010  Kenneth Ives kenaso@tx.rr.com
'              Added parameter denoting if sort should be in ascending or
'              descending order.
' ***************************************************************************
Public Function CombSort(ByRef avntData As Variant, _
                Optional ByVal blnAscending As Boolean = True) As Boolean
       
    ' Called by RemoveDupes()
    '           NonRepeatingNbrs()
    
    Dim lngGap     As Long
    Dim lngLow     As Long
    Dim lngHigh    As Long
    Dim lngIndex   As Long
    Dim vntHold    As Variant
    Dim blnSwapped As Boolean
    
    Const ROUTINE_NAME As String = "CombSort"
      
    On Error GoTo CombSort_Error

    CombSort = False  ' Preset to FALSE
    vntHold = Empty   ' Always start with empty variants
    
    If CBool(IsArrayInitialized(avntData)) Then
        
        lngLow = LBound(avntData)       ' Lowest number of elements
        lngHigh = UBound(avntData) - 1  ' Highest number of elements
        lngGap = lngHigh - lngLow + 1   ' Total element count
        blnSwapped = False              ' Set swapped flag to FALSE
    
        ' too much overhead if only
        ' two items in the array
        If lngHigh <= 2 Then
            
            If blnAscending Then
                If (avntData(0) > avntData(1)) Then
                    SwapData avntData(0), avntData(1)
                End If
            End If
            
        Else
        
            ' Loop until the GAP is less than one
            ' or nothing is swapped
            Do While (lngGap > 1) Or blnSwapped
                
                ' divide the Gap by 1.3
                lngGap = (10 * lngGap) \ 13
                
                ' The most efficient series of a
                ' final GAP values starts with 11
                If (lngGap = 9 Or lngGap = 10) Then
                    lngGap = 11
                End If
                
                ' Preset flag to FALSE thus designating
                ' nothing has been swapped
                blnSwapped = False
                
                ' Compare and possibly swap values
                ' and pointers separated by GAP
                For lngIndex = lngLow To (lngHigh - lngGap)
                    
                    ' Capture current data for comparison
                    vntHold = avntData(lngIndex)   
                    
                    ' Is hold data greater than current array data?
                    If (vntHold > avntData(lngIndex + lngGap)) Then
                        
                        ' Move data to new position within array
                        avntData(lngIndex) = avntData(lngIndex + lngGap)
                        avntData(lngIndex + lngGap) = vntHold
                        
                        vntHold = Empty    ' Empty holding area
                        blnSwapped = True  ' Reset swap flag
                        
                    End If
                   
                Next lngIndex
            
                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    Exit Do
                End If
    
            Loop
        
            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                GoTo CombSort_CleanUp
            End If
        
            ' If descending order is requested
            ' then reverse data within array
            If Not blnAscending Then
                
                Do While lngLow < lngHigh
                    
                    SwapData avntData(lngLow), avntData(lngHigh)  ' swap data elements
                    lngLow = lngLow + 1                           ' increment low index
                    lngHigh = lngHigh - 1                         ' decrement high index
                
                Loop
            End If
        End If
    End If
    
CombSort_CleanUp:
    vntHold = Empty   ' Always empty variants when not needed

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        CombSort = False   ' Bad finish
    else
        CombSort = True    ' Successful completion
    End If
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CombSort_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CombSort_CleanUp

End Function

' ***************************************************************************
' Routine:       CreateHash
'
' Description:   Generate a one-way hash string from a string of data using
'                CryptoAPI. These are the algorithms available:
'
'                     MD2  MD4  MD5  SHA-1  SHA-256  SHA-384  SHA-512
'
'                Special note:  SHA-512/224 and SHA-512/256 have not yet
'                been implemented into the Microsoft crypto suite of hashes.
'
'                SHA-2 support on MS Windows
'                Paraphrasing:  Regarding SHA-224 support, SHA-224 offers
'                less security than SHA-256 but takes the same amount of
'                resources.  Also SHA-224 is not generally used by protocols
'                and applications.  The National Security Agency (NSA)
'                Suite B standards also does not include it.  Microsoft
'                has no plans to add it to future versions of their
'                Cryptographic Service Providers (CSP).
'                http://blogs.msdn.com/b/alejacma/archive/2009/01/23/sha-2-support-on-windows-xp.aspx
'
'                Hashes are extremely useful for determining whether a
'                transmission or file has been altered.  No two hashes are
'                alike unless the string matches perfectly, whether binary
'                data or a text string.
'
'                Be aware that if you choose to not convert the raw hashed data
'                to hex, then hashes may not store the text correctly because
'                of the possible existence of non printable characters in
'                the stream.  It is recommended you store unconverted data in
'                a byte array.
'
' Reference:     How to encrypt a string in Visual Basic 6.0 and how to decrypt
'                the string in Visual Basic .NET or in Visual Basic 2005
'                http://support.microsoft.com/kb/821762
'
' WARNING:       MD4, MD5, RIPEMD Algorithms have been compromised at the rump
'                session of Crypto 2004.  It was announced that Xiaoyun Wang,
'                Dengguo Feng, Xuejia Lai and Hongbo Yu found collisions for
'                MD4, MD5, RIPEMD, and the 128-bit version of HAVAL.
'                http://eprint.iacr.org/2004/199.pdf
'
'                Feb-2005:  SHA-1 has been compromised.  Recommended that you
'                do not use for password or document authentication.
'                http://www.schneier.com/blog/archives/2005/02/sha1_broken.html
'                http://csrc.nist.gov/groups/ST/toolkit/documents/shs/NISTHashComments-final.pdf
'
'                March 15, 2006:  The SHA-2 family of hash functions
'                (i.e., SHA-224, SHA-256, SHA-384 and SHA-512) may be used
'                by Federal agencies for all applications using sedble hash
'                algorithms. Federal agencies should stop using SHA-1 for
'                digital signatures, digital time stamping and other
'                applications that require collision resistance as soon as
'                practical, and must use the SHA-2 family of hash functions
'                for these applications after 2010. After 2010, Federal
'                agencies may use SHA-1 only for the following applications:
'                    - hash-based message authentication codes (HMACs)
'                    - key derivation functions (KDFs)
'                    - random number generators (RNGs)
'                Regardless of use, NIST encourages application and protocol
'                designers to use the SHA-2 family of hash functions for all
'                new applications and protocols.
'                http://csrc.nist.gov/groups/ST/hash/policy.html
'
' Parameters:    strInput - string of data to be hashed.
'                lngHashSelection - [OPTIONAL] Type of hash algorithm to use.
'                       Default = ePRNG_SHA512
'                blnReturnHexString - [OPTIONAL] Format for the hashed data
'                       TRUE - Return data in hex string [Default]
'                       FALSE - Return data in hex array
'
' Returns:       Hex string of characters or a string array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 25-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated variable names
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to access SHA2 family of hash algorithms.
' 03-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Fix a bug.  Passed hash selection was always zero.
'              - Changed code structure to reduce number of GoTo statements.
'              - Changed default hash selection to strongest possible hash
'                for this version of Windows.
'              - Allow for data to be returned as string or hex array
'              - Updated documentation.
' 24-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation.
' ***************************************************************************
Public Function CreateHash(ByVal strInput As String, _
                  Optional ByVal lngHashSelection As enumPRNG_HashAlgorithm = ePRNG_SHA512, _
                  Optional ByVal blnReturnHexString As Boolean = True) As Variant

    ' Called by CreateExtraSeed()
    
    Dim lngHashAlgo   As Long
    Dim lngHashObject As Long
    Dim lngHashLength As Long
    Dim strHash       As String
    Dim astrHash()    As String

    Const ROUTINE_NAME As String = "CreateHash"

    On Error GoTo CreateHash_Error

    strHash = vbNullString   ' Verify receiving area is empty
    lngHashAlgo = 0          ' Initialize hash parameters
    lngHashLength = 0
    
    ' Test for a provider handle
    If mlngProviderHandle < 1 Then
        GetProviderHandle   ' Get provider handle
    End If

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo CreateHash_CleanUp
    End If

    ' Determine which hash algorithm
    ' to use and its return length
    HashSelection lngHashSelection, lngHashAlgo, lngHashLength
    
    strHash = Space$(lngHashLength)  ' Preload with blank spaces
    lngHashObject = 0                ' Initialize hash object
    
    ' Step 1.  Create hash object.  The CryptCreateHash function
    '          initiates the hashing of a stream of data.  It
    '          creates and returns to the calling application a
    '          handle to a CSP hash object.  This handle is used
    '          in subsequent calls to CryptHashData to hash
    '          session keys and other streams of data.  Failure
    '          is a return value of zero.
    If CryptCreateHash(mlngProviderHandle, lngHashAlgo, _
                       ByVal 0&, ByVal 0&, lngHashObject) <> 0 Then
        
        ' Step 2.  Hash in data and create a unique handle.
        '          The CryptHashData function adds data to
        '          a specified hash object.  Failure is a
        '          return value of zero.
        If CryptHashData(lngHashObject, strInput, Len(strInput), ByVal 0&) <> 0 Then
                    
            ' Step 3.  The CryptGetHashParam function retrieves
            '          data that governs the operations of a hash
            '          object. Actual hash value can be captured
            '          by using this function.  Failure is a return
            '          value of zero.
            If CryptGetHashParam(lngHashObject, HP_HASHVAL, strHash, _
                                 lngHashLength, ByVal 0&) <> 0 Then
                
                ' Capture raw hashed data
                strHash = Left$(strHash, lngHashLength)
                
                ' Format data to a hex format and return as string
                ' or in an array.
                If blnReturnHexString Then
                    ' convert hashed data to hex string
                    strHash = ConvertDataToHex(lngHashLength, strHash, True)
                Else
                    ' convert hashed data to string array
                    astrHash() = ConvertDataToHex(lngHashLength, strHash, False)
                End If
                    
            Else
                InfoMsg "Failed to retrieve hashed data." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
            End If
        Else
            InfoMsg "Failed to hash data." & _
                    vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
            gblnStopProcessing = True
        End If
    Else
        InfoMsg "Failed to create hash object key." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
    End If

CreateHash_CleanUp:
    ' CryptDestroyHash function releases
    ' handle the hKey parameter references
    If lngHashObject <> 0 Then
        CryptDestroyHash lngHashObject
        lngHashObject = 0
    End If

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    If blnReturnHexString Then
        CreateHash = strHash     ' Return hex string
    Else
        CreateHash = astrHash()  ' Return hex string array
    End If
        
    Erase astrHash()             ' Always empty arrays when not needed
    strHash = vbNullString       ' Verify holding area is empty
    
    On Error GoTo 0              ' Nullify this error trap
    Exit Function

CreateHash_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CreateHash_CleanUp

End Function

' ***************************************************************************
' Routine:       GetRndValue
'
' Description:   Generate a random long integer between two input values
'                using CryptoAPI random number generator.
'
'                Equivalent to:    x = Int(Rnd() * (Max - Min + 1)) + Min
'
'                Visual Basic Language Reference Rnd Function
'                http://msdn2.microsoft.com/en-us/library/f7s023d2(VS.71).aspx
'
' Parameters:    dblLow - Lowest possible value
'                dblHigh - highest possible value
'
' Returns:       long integer value between low to high input inclusive
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-SEP-1998  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 05-MAR-2005  Kenneth Ives  kenaso@tx.rr.com
'              Updated the seed algorithm
' 25-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote GetRndValue() routine.  Thanks to Alfred Hellmüller
'              for seeing the shortcomings of using the Visual Basic RND()
'              function.
' 12-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added additional range testing for long integers only.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified logic for obtaining a valid value
' 10-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Verify input parameters fall within Long Integer range
' 30-Dec-2011  Kenneth Ives  kenaso@tx.rr.com
'              Bug fix - if low were negative and high were zero
' ***************************************************************************
Public Function GetRndValue(ByVal dblLow As Double, _
                            ByVal dblHigh As Double) As Long

    Dim blnGood    As Boolean   ' Flag - found good return value
    Dim blnBothNeg As Boolean   ' Flag - both inputs are negative
    Dim blnBothPos As Boolean   ' Flag - both inputs are positive
    Dim lngIndex   As Long      ' Array index pointer
    Dim dblValue   As Double    ' Return value
    Dim adblData() As Double    ' Array of random values

    blnGood = False     ' Preset flags to FALSE
    blnBothNeg = False
    blnBothPos = False
         
    dblLow = Fix(dblLow)    ' Drop decimal positions
    dblHigh = Fix(dblHigh)
    
    ' Verify parameters fall within long integer range
    ' -2147483648 to 2147483647
    '
    ' If input for low value exceeds long integer
    ' minimum then use long integer minimum
    If dblLow < MIN_LONG Then
        dblLow = CDbl(MIN_LONG)
    End If
    
    ' If input for upper value exceeds long integer
    ' maximum then use long integer maximum
    If dblHigh > MAX_LONG Then
        dblHigh = CDbl(MAX_LONG)
    End If
    
    ' If low is greater than or equal high
    If dblLow >= dblHigh Then
        GetRndValue = dblLow   ' Return low value
        Exit Function
    End If

    ' If only one digit difference
    If dblHigh - dblLow = 1 Then
        GetRndValue = dblLow   ' Return low value
        Exit Function
    End If
    
    ' See if both values are positive
    If dblLow >= 0 And dblHigh > 0 Then
        
        blnBothPos = True  ' Set flag to TRUE
        
    ' See if both values are negative
    ElseIf dblLow < 0 And dblHigh <= 0 Then
        
        blnBothNeg = True         ' Set flag to TRUE
        dblLow = Abs(dblLow)      ' make both input values positive
        dblHigh = Abs(dblHigh)
        SwapData dblLow, dblHigh  ' swap data
    End If
    
    ' Fill array with random generated double values.
    ' More than enough data to work with.  Normally
    ' just the first element will be used.
    adblData() = BuildRndData(5, ePRNG_DBL_ARRAY, False)
    
    ' Loop thru random array and process
    ' each element until end of array
    For lngIndex = 0 To UBound(adblData) - 1
        
        ' Calculate a non-decimal value
        dblValue = CDbl(Int(adblData(lngIndex) * (dblLow + (dblHigh - dblLow))))
        
        ' If both inputs are positive or both
        ' are negative then make calculated
        ' result positive. This reduces the
        ' number of calculations.
        If blnBothPos Or blnBothNeg Then
            dblValue = Abs(dblValue)
        End If
        
        ' Is value greater than or equal to low range and
        ' is value less than or equal to upper range ?
        If (dblValue >= dblLow) And _
           (dblValue <= dblHigh) Then
             
            ' If both input values were negative
            If blnBothNeg Then
            
                ' If value greater than zero
                ' make return value negative
                If dblValue > 0 Then
                    dblValue = -dblValue
                End If
            
            End If
        
            blnGood = True  ' Found a valid value
            Exit For        ' exit For..Next loop

        End If
            
    Next lngIndex
    
    If blnGood Then
        GetRndValue = CLng(dblValue)   ' Return random value
    Else
        ' See if both input values were negative
        If blnBothNeg Then
            ' Low and high values were previously swapped.
            ' If high value greater than zero make return
            ' value negative and insert into low value.
            If dblHigh > 0 Then
                dblLow = -dblHigh
            End If
        End If

        GetRndValue = CLng(dblLow)    ' Safety - should not get to here
    End If

End Function

' ***************************************************************************
' Routine:       LongToDouble
'
' Description:   Convert a long integer to a double precision number.
'                Returns a decimal position of 14 places. This was
'                chosen so that exponentials would not be returned.
'                If you are working with large numbers then your return
'                values will probably be in string format so as to get
'                the full value of your calculations.
'
'                    Ex: 2147483647  ->  -2.32830732662872E-10
'
'                Desired return value is within this range:
'
'                    -0.99999999999999 to 0.99999999999999
'
'                    Ex: 2147483647  ->  -0.00000000023283
'
'                By changing the return value from Double to String and
'                Uncommenting out the line of code indicated below, you
'                will get an extended return value.
'
'                    Ex: 2147483647  ->  -0.000000000232830732662872
'
' Parameters:    lngValue - Long integer to be converted
'
' Returns:       Double precision value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function LongToDouble(ByVal lngValue As Long) As Double

    ' Called by BuildRndData()
    
    Dim dblTemp As Double

    If lngValue < 0 Then
        dblTemp = lngValue + GB_4
    Else
        dblTemp = lngValue
    End If
 
    LongToDouble = FormatNumber(DBL_LOW + ((dblTemp * DBL_HIGH) - 1#), 14)
    
    ' Uncomment to create a string return value
    'LongToDouble = CStr(CDec(DBL_LOW + ((dblTemp * DBL_HIGH) - 1#)))

End Function

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer

    On Error GoTo IsArrayInitialized_Exit
    
    IsArrayInitialized = 0  ' preset to FALSE
    
    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then
        
        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&
    
    End If
    
IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       EmptyCollection
'
' Description:   Properly empty and deactivate a collection
'
' Parameters:    colData - Collection to be processed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub EmptyCollection(ByRef colData As Collection)

    ' Has collection been deactivated?
    If colData Is Nothing Then
        Exit Sub
    End If
    
    ' Is the collection empty?
    Do While colData.Count > 0
        
        ' Parse backwards thru collection and delete data.
        ' Backwards parsing prevents a collection from
        ' having to reindex itself after each data removal.
        colData.Remove colData.Count
    Loop
    
    ' Free collection object from memory
    Set colData = Nothing
    
End Sub

' ***************************************************************************
' Routine:       SwapData
'
' Description:   Swap data with each other.  I wrote this function since
'                BASIC stopped having its own SWAP function.  I use this
'                for swapping strings, type structures, numbers with
'                decimal values, etc.
'
' Parameters:    vntValue1 - Incoming data to be swapped with Value2
'                vntValue2 - Incoming data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function SwapData(ByRef vntData1 As Variant, _
                         ByRef vntData2 As Variant)

    ' Called by ReshuffleData()
    '           CombSort()
    '           GetRndValue()
    
    Dim vntHold As Variant

    vntHold = Empty   ' Always start with empty variants

    vntHold = vntData1
    vntData1 = vntData2
    vntData2 = vntHold

    vntHold = Empty   ' Always empty variants when not needed

End Function

' ***************************************************************************
' Routine:       SwapLong
'
' Description:   Swap data (byte, integer, or long) with each other without
'                using a temp.  Using Math to do this takes a substantial
'                amount more processing than Logic Gates. The logic gates
'                foundation is in all processors. Which you could argue that
'                so is Math, but running a math process with a temp variable
'                in fact uses a considerable amount more processor cycles
'                than logic gates.
'
'                Using 3 Xor's over 3 Basic Equations greatly speeds up the
'                application. (especially when dealing with millions of
'                equations/Logic Gates)
'
'                This routine works with byte, integer and long values.
'                Change the parameter data type accordingly.
'
' Parameters:    AA - data to be swapped with Value2
'                BB - data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Mar-2003  Ulli
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=43903&lngWId=1
' ***************************************************************************
Public Sub SwapLong(ByRef AA As Long, _
                    ByRef BB As Long)

    ' Swap long integer values
    
    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub

Public Sub SwapInt(ByRef AA As Integer, _
                   ByRef BB As Integer)

    ' Swap short integer values
    
    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub

Public Sub SwapByte(ByRef AA As Byte, _
                    ByRef BB As Byte)

    ' Swap byte values
    
    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub

' ***************************************************************************
' Routine:       ByteArrayToString
'
' Description:   Converts a byte array to string data
'
' Parameters:    abytData - array of bytes
'
' Returns:       Data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function ByteArrayToString(ByRef abytData() As Byte) As String

    ' Called by BuildWithinRange()
    
    ByteArrayToString = StrConv(abytData(), vbUnicode)

End Function

' ***************************************************************************
' Routine:       StringToByteArray
'
' Description:   Converts string data to a byte array
'
' Parameters:    strData - Data string to be converted
'
' Returns:       byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function StringToByteArray(ByVal strData As String) As Byte()

    ' Called by BuildRndData()
    
     StringToByteArray = StrConv(strData, vbFromUnicode)

End Function

' ***************************************************************************
' Procedure:     RndSeed
'
' Description:   Reseeds VB random number generator with a unique value.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function RndSeed()

    ' Called by ReshuffleData()
    
    CreateExtraSeed   ' Reseed VB random number generator only
    
End Function

' ***************************************************************************
' Procedure:     CreateExtraSeed
'
' Description:   Creates a unique string of hex data using CryptoAPI hash
'                functions.  Also, randomly select a starting position in
'                hashed data string to capture two eight byte strings of
'                data.  These will be converted into long integers for new
'                carryover values.
'
' Parameters:    lngReturnLen - Optional - Length of return string
'                    (0-128 bytes)   Default = 0
'
' Returns:       Hex data string after hashing
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Dec-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated data collection logic
' ***************************************************************************
Public Function CreateExtraSeed(Optional ByVal lngReturnLen As Long = 0) As String
        
    ' Called by BuildRndData()
    '           RndSeed()
    
    Dim lngData      As Long
    Dim strHex       As String
    Dim strHash      As String
    Dim strCarryOver As String
    
    On Error Resume Next
        
    ' Process carryover values
    With mtypCarryOver
        
        ' Load carryover values with current time
        ' in milliseconds, if either half is zero
        If .Value1 = 0 Or .Value2 = 0 Then
           
            lngData = GetTickCount()                                        ' Capture milliseconds
            QueryPerformanceCounter mcurCounter                             ' Capture Hi-Performance milliseonds
            mcurSeed = Fix(((mcurCounter / mcurFrequency) * 10000) * 1000)  ' Calculate seed value
    
            .Value1 = CLng(mcurSeed)             ' Hi-Performance seed
            .Value2 = (lngData - (lngData \ 4))  ' Subtract approx 25%
        End If
    
        ' Format carryover string by converting
        ' numeric values to hex
        '
        '                Numeric  to   Hex
        ' Ex:  Value1  1234567890    499602D2
        '      Value2     1234567    0012D687
        '
        '      strCarryOver = 499602D20012D687
        '
        strCarryOver = Right$(ZEROES_8 & Hex$(.Value1), 8) & _
                       Right$(ZEROES_8 & Hex$(.Value2), 8)
    End With
    
    ' Perform hash using strongest algorithm
    ' allowed for this version of Windows
    '
    '           |--------SHA-1 (40 chars)--------------|
    ' strHash = 12E001FCF95F1290A40085159B908E25579186AF
    '   or
    ' strHash = 019862B34470730CA9...C1D3ADF1932D53ADB45
    '           |--------SHA-512 (128 chars)-----------|
    '
    strHash = CreateHash(strCarryOver)

    ' See if there was a hashing error
    If Len(Trim$(strHash)) = 0 Then

        ZeroMemory mtypCarryOver, Len(mtypCarryOver)   ' Empty carryover type structure    
        lngData = GetTickCount                         ' Capture number of milliseconds
        Rnd (-1)                                       ' Reset VB Random Number Generator
        Randomize (lngData - (lngData \ 10))           ' Reseed VB Random Number Generator
        GoTo CreateExtraSeed_CleanUp                   ' Exit this routine

    End If

    ' Reseed VB random number generator by randomly
    ' selecting an eight character data string and
    ' converting it to a positive value.
    strHex = Mid$(strHash, Int(Rnd() * (Len(strHash) - 7)) + 1, 8)   ' Capture eight characters
    Rnd (-1)                                                         ' Reset VB RNG
    Randomize Abs(CDbl("&H" & strHex))                               ' Reseed VB RNG

    ' Update carryover data by capturing two random eight
    ' byte strings and converting them into long integers
    With mtypCarryOver
        .Value1 = Val("&H" & Mid$(strHash, Int(Rnd() * (Len(strHash) - 7)) + 1, 8))
        .Value2 = Val("&H" & Mid$(strHash, Int(Rnd() * (Len(strHash) - 7)) + 1, 8))
    End With
        
    ' Determine amount of data to be returned
    Select Case lngReturnLen
           Case Is < 1:             CreateExtraSeed = vbNullString   ' No data to return
           Case Is >= Len(strHash): CreateExtraSeed = strHash        ' Return complete hex string

           ' Hex string is longer than return length. Randomly
           ' select a starting position within hex string and
           ' return requested amount of data.
           Case Else
                CreateExtraSeed = Mid$(strHash, Int(Rnd() * (Len(strHash) - (lngReturnLen - 1))) + 1, lngReturnLen)
    End Select

CreateExtraSeed_CleanUp:
    strHex = vbNullString        ' Empty variables
    strHash = vbNullString
    strCarryOver = vbNullString
    On Error GoTo 0              ' Nullify this error trap
    
End Function



' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:       ConvertDataToHex
'
' Description:   Convert hashed data into 2-char hex string format or
'                hex string array.
'
' Parameters:    lngReturnLength - Length of returned data
'                strHashedData - incoming hash data string
'                blnReturnString - Return data in string or array format
'
' Returns:       Either string or array of data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-DEC-2000  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 24-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modifed and documented
' ***************************************************************************
Private Function ConvertDataToHex(ByVal lngReturnLength As Long, _
                                  ByVal strHashedData As String, _
                                  ByVal blnReturnString As Boolean) As Variant

    ' Called by BuildRndData()
    '           BuildWithinRange()
    '           CreateHash()
    
    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim strOutput  As String
    Dim astrHash() As String

    Const ROUTINE_NAME As String = "ConvertDataToHex"

    On Error GoTo ConvertDataToHex_Error

    lngLength = Len(strHashedData)  ' capture length of incoming data

    If blnReturnString Then
        
        ' Return hex string
        strOutput = Space$(lngReturnLength * 2 + 4)   ' preload output string
        lngPointer = 1                                ' initialize output pointer

        ' parse the data and convert each character to hex
        For lngIndex = 1 To lngLength

            Mid$(strOutput, lngPointer, 2) = Right$("00" & Hex$(Asc(Mid$(strHashedData, lngIndex, 1))), 2)
            lngPointer = lngPointer + 2

            ' See if the limits have been reached
            If lngIndex = lngReturnLength Then
                Exit For    ' exit For..Next loop
            End If

        Next lngIndex

        ConvertDataToHex = Trim$(Left$(strOutput, lngReturnLength * 2)) ' return hex string
    
    Else
        ' Return hex array
        ReDim astrHash(lngReturnLength)   ' Size output array
        lngIndex = 0                      ' initialize array index

        ' parse incoming data and convert
        ' to decimal value
        For lngPointer = 1 To lngLength

            astrHash(lngIndex) = Right$("00" & Hex$(Asc(Mid$(strHashedData, lngPointer, 1))), 2)
            lngIndex = lngIndex + 1

            ' See if the limits have been reached
            If lngIndex = lngReturnLength Then
                Exit For    ' exit For..Next loop
            End If

        Next lngPointer
    
        ConvertDataToHex = astrHash()  ' return hex array
    
    End If

ConvertDataToHex_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ConvertDataToHex_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ConvertDataToHex = Empty
    gblnStopProcessing = True
    Resume ConvertDataToHex_CleanUp

End Function

' ***************************************************************************
' Procedure:     HashSelection
'
' Description:   User passes a numeric value, usually 0-6, and this routine
'                will determine the hash selection.  No data needs to be
'                passed to the HashLength parameter as this is output only.
'                If the AES provider is not available then the parameter
'                passed will be calculated to use a lesser hash algorithm.
'
' Parameters:    lngHashSelection - Input - Selected hash algorithm
'                lngHashAlgo  - Output - Hash algorithm selection
'                lngHashLength - Output - hash internal return length
'
' Returns:       Five digit code representing appropriate hash selection
'                along with the hashed value return byte length.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 04-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added range testing and updated documentation
' ***************************************************************************
Private Sub HashSelection(ByVal lngHashSelection As enumPRNG_HashAlgorithm, _
                          ByRef lngHashAlgo As Long, _
                          ByRef lngHashLength As Long)
                 
    ' Called by CreateHash()
    
    Select Case mblnAESProvider
           
           Case True   ' Strong (AES) provider available
                
                Select Case lngHashSelection
                       Case ePRNG_MD2:    lngHashAlgo = CALG_MD2:     lngHashLength = 16    ' 0
                       Case ePRNG_MD4:    lngHashAlgo = CALG_MD4:     lngHashLength = 16    ' 1
                       Case ePRNG_MD5:    lngHashAlgo = CALG_MD5:     lngHashLength = 16    ' 2
                       Case ePRNG_SHA1:   lngHashAlgo = CALG_SHA1:    lngHashLength = 20    ' 3
                       Case ePRNG_SHA256: lngHashAlgo = CALG_SHA_256: lngHashLength = 32    ' 4
                       Case ePRNG_SHA384: lngHashAlgo = CALG_SHA_384: lngHashLength = 48    ' 5
                       Case ePRNG_SHA512: lngHashAlgo = CALG_SHA_512: lngHashLength = 64    ' 6
                End Select
    
           Case False  ' Default or Enhanced provider available
                
                DoEvents
                If lngHashSelection > 3 Then
                    lngHashSelection = ePRNG_SHA1
                End If
                
                Select Case lngHashSelection
                       Case ePRNG_MD2:  lngHashAlgo = CALG_MD2:  lngHashLength = 16   ' 0
                       Case ePRNG_MD4:  lngHashAlgo = CALG_MD4:  lngHashLength = 16   ' 1
                       Case ePRNG_MD5:  lngHashAlgo = CALG_MD5:  lngHashLength = 16   ' 2
                       Case ePRNG_SHA1: lngHashAlgo = CALG_SHA1: lngHashLength = 20   ' 3
                End Select
    End Select
    
End Sub

' ***************************************************************************
' Routine:       GetProviderHandle
'
' Description:   This is a generic routine to capture the MS provider
'                handle to see if any enhanced encryption capabilities are
'                available.
'
'                The Advanced Encryption Standard (AES) cipher algorithm
'                and SHA2 family of hash algorithms were not available until
'                Windows XP with SP3 installed.  The container provider for
'                Windows XP SP3 was named the same as later versions except
'                with the word prototype appended.  (ex:  "(Prototype)")
'                See constant MS_ENH_RSA_AES_PROV_XP in Declare section of
'                this module.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Feb-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for bringing to my attention the
'              need to update this routine.  Rewrote this routine to test
'              for availability of SHA2 hash functionality
' 04-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              Changed code structure to reduce number of GoTo statements.
' ***************************************************************************
Private Sub GetProviderHandle()

    ' Called by BuildRndData()
    '           CreateHash()
    '           AES_Ready() [Property]
    
    Dim strContainer As String   ' Container value
    Dim strProvider  As String   ' Name of provider

    Const ROUTINE_NAME As String = "GetProviderHandle"
    
    On Error GoTo GetProviderHandle_Error

    ReleaseProvider   ' Release any hold on provider handle
    
    ' Set search parameters
    strContainer = vbNullChar
    strProvider = MS_ENH_RSA_AES_PROV_XP & vbNullChar

    ' Attempt to acquire a handle to strong key container
    ' for Windows XP SP3. The "Prototype" is only in this
    ' version of Windows. Windows XP SP2 or earlier versions
    ' normally do not have this functionality.
    ' Failure is a return value of zero.
    If CryptAcquireContext(mlngProviderHandle, _
                           ByVal strContainer, ByVal strProvider, _
                           PROV_RSA_AES, CRYPT_VERIFYCONTEXT) <> 0 Then
        
        mblnAESProvider = True   ' Set property value (AES capability is available)
        
    Else
        ' Update provider search parameter
        strProvider = MS_ENH_RSA_AES_PROV & vbNullChar
    
        ' Attempt to acquire a handle to strong key container
        ' for Windows Vista or newer.  Failure is a return
        ' value of zero.
        If CryptAcquireContext(mlngProviderHandle, _
                               ByVal strContainer, ByVal strProvider, _
                               PROV_RSA_AES, CRYPT_VERIFYCONTEXT) <> 0 Then
            
            mblnAESProvider = True   ' Set property value (AES capability is available)
            
        Else
            ' Update provider search parameter
            strProvider = MS_ENHANCED_PROV & vbNullChar
            
            ' Attempt to acquire a handle to enhanced key container.
            ' Failure is a return value of zero.
            If CryptAcquireContext(mlngProviderHandle, _
                                   ByVal strContainer, ByVal strProvider, _
                                   PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
                
                ' Did not find enhanced provider.
                ' Update provider search parameter.
                strProvider = MS_DEFAULT_PROV & vbNullChar
                
                ' Attempt to acquire a handle to default key container
                ' Failure is a return value of zero.
                If CryptAcquireContext(mlngProviderHandle, _
                                       ByVal strContainer, ByVal strProvider, _
                                       PROV_RSA_FULL, ByVal 0&) = 0 Then
            
                    InfoMsg "An unknown error occurred while retrieving" & vbNewLine & _
                            "CryptoAPI provider handle." & _
                            vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    
                    ReleaseProvider            ' Verify provider handle is clear
                    gblnStopProcessing = True  ' Time to stop processing
                    
                End If
            End If
        End If
    End If
        
GetProviderHandle_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetProviderHandle_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ReleaseProvider                   ' Release provider handle
    gblnStopProcessing = True         ' Time to stop processing
    Resume GetProviderHandle_CleanUp

End Sub

Private Sub ReleaseProvider()

    ' Called by GetProviderHandle()
    '           Class_Terminate()
    
    If mlngProviderHandle <> 0 Then
        CryptReleaseContext mlngProviderHandle, 0&
    End If
    
    mblnAESProvider = False  ' Set property value (AES capability not available)
    mlngProviderHandle = 0   ' Numeric value depicting provider handle
   
End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objPrng = new cPrng
                
    ZeroMemory mtypCarryOver, Len(mtypCarryOver)  ' Empty carryover type structure
    StopProcessing = False                        ' Preset property to FALSE
    CompareMethod = ePRNG_CaseSensitive           ' Preset comparison method
    
    QueryPerformanceFrequency mcurFrequency   ' Capture Hi-Performance frequency counter
    mcurFrequency = mcurFrequency * 10000     ' Format without decimals
    QueryPerformanceCounter mcurCounter       ' Capture Hi-Performance milliseconds
    
    ' Calculate seed value
    mcurSeed = (((mcurCounter / mcurFrequency) * 10000) * 1000) - CCur(Timer)
    
    Rnd (-1)                  ' Reset VB Random Number Generator
    Randomize Fix(mcurSeed)   ' Reseed VB Random Number Generator
    
End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objPrng = Nothing
            
    ReleaseProvider                               ' Release provider handle (no longer needed)
    ZeroMemory mtypCarryOver, Len(mtypCarryOver)  ' Empty carryover type structure
    
End Sub

