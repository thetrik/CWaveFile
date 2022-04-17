VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWaveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CWaveFile.cls - class for working with WAVE-PCM files
' // By The trick, 2021
' // ver. 1.0
' //

Option Explicit
Option Base 0

Public Enum eChannelsMask
    CM_0 = 1
    CM_1 = 2
    CM_2 = 4
    CM_3 = 8
    CM_4 = 16
    CM_5 = 32
    CM_6 = 64
    CM_7 = 128
    CM_LR = 3
    CM_ALL = 255
End Enum

Private Const MODULE_NAME   As String = "CWaveFile"

Private Const FOURCC_MEM                          As Long = &H204D454D
Private Const MMIO_CREATERIFF                     As Long = &H20
Private Const MMIO_DIRTY                          As Long = &H10000000
Private Const MMIO_CREATE                         As Long = &H1000
Private Const MMIO_WRITE                          As Long = &H1
Private Const MMIO_READWRITE                      As Long = &H2
Private Const WAVE_FORMAT_PCM                     As Long = 1
Private Const WAVE_FORMAT_EXTENSIBLE              As Long = -2
Private Const WAVE_FORMAT_IEEE_FLOAT              As Long = 3
Private Const SEEK_SET                            As Long = 0
Private Const MMIO_FINDCHUNK                      As Long = &H10
Private Const MMIO_FINDRIFF                       As Long = &H20
Private Const GMEM_FIXED                          As Long = &H0
Private Const FADF_AUTO                           As Long = 1
Private Const FOURCC_RIFF                         As Long = &H46464952
Private Const FOURCC_WAVE                         As Long = &H45564157
Private Const FOURCC_FMT                          As Long = &H20746D66
Private Const FOURCC_DATA                         As Long = &H61746164
Private Const KSDATAFORMAT_SUBTYPE_PCM_STR        As String = "00000001-0000-0010-8000-00AA00389B71"
Private Const KSDATAFORMAT_SUBTYPE_IEEE_FLOAT_STR As String = "00000003-0000-0010-8000-00aa00389b71"
Private Const SND_MEMORY                          As Long = &H4         '  lpszSoundName points to a memory file
Private Const SND_SYNC                            As Long = &H0         '  play synchronously (default)

Private Type UUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements           As Long
    lLBound             As Long
End Type

Private Type SAFEARRAY
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    Bounds              As SAFEARRAYBOUND
End Type

Private Type MMCKINFO
    ckid                As Long
    ckSize              As Long
    fccType             As Long
    dwDataOffset        As Long
    dwFlags             As Long
End Type

Private Type MMIOINFO
    dwFlags             As Long
    fccIOProc           As Long
    pIOProc             As Long
    wErrorRet           As Long
    htask               As Long
    cchBuffer           As Long
    pchBuffer           As Long
    pchNext             As Long
    pchEndRead          As Long
    pchEndWrite         As Long
    lBufOffset          As Long
    lDiskOffset         As Long
    adwInfo(4)          As Long
    dwReserved1         As Long
    dwReserved2         As Long
    hmmio               As Long
End Type

Private Type WAVEFORMATEXTENSIBLE
    wFormatTag          As Integer
    nChannels           As Integer
    nSamplesPerSec      As Long
    nAvgBytesPerSec     As Long
    nBlockAlign         As Integer
    wBitsPerSample      As Integer
    cbSize              As Integer
    wValidBitsPerSample As Integer
    dwChannelMask       As Long
    guidSubFormat       As UUID
End Type

Private Type tTestWaveFile
    lRIFFID             As Long
    lRIFFSize           As Long
    lDataTypeID         As Long
    lFmtCunkID          As Long
    lFmtChunkSize       As Long
    wFormatTag          As Integer
    wNumOfChannels      As Integer
    nSamplesPerSec      As Long
    nAvgBytesPerSec     As Long
    nBlockAlign         As Integer
    wBitsPerSample      As Integer
    lDataChunkID        As Long
    lDataChunkSize      As Long
End Type

Private Type tTestWaveExFile
    lRIFFID             As Long
    lRIFFSize           As Long
    lDataTypeID         As Long
    lFmtCunkID          As Long
    lFmtChunkSize       As Long
    wFormatTag          As Integer
    wNumOfChannels      As Integer
    nSamplesPerSec      As Long
    nAvgBytesPerSec     As Long
    nBlockAlign         As Integer
    wBitsPerSample      As Integer
    cbSize              As Integer
    wValidBits          As Integer
    dwChannelMask       As Long
    guidSubFormat       As UUID
    lDataChunkID        As Long
    lDataChunkSize      As Long
End Type

Private Declare Function mmioClose Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         Optional ByVal uFlags As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" _
                         Alias "mmioOpenW" ( _
                         ByVal szFileName As Long, _
                         ByRef lpmmioinfo As Any, _
                         ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" _
                         Alias "mmioStringToFOURCCA" ( _
                         ByVal sz As String, _
                         ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef lpck As MMCKINFO, _
                         ByVal uFlags As Long) As Long
Private Declare Function mmioCreateChunk Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef lpck As MMCKINFO, _
                         ByVal uFlags As Long) As Long
Private Declare Function mmioWrite Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef pch As Any, _
                         ByVal cch As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef lpck As MMCKINFO, _
                         ByRef lpckParent As Any, _
                         ByVal uFlags As Long) As Long
Private Declare Function mmioSeek Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByVal lOffset As Long, _
                         ByVal iOrigin As Long) As Long
Private Declare Function memcpy Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function mmioRead Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef pch As Any, _
                         ByVal cch As Long) As Long
Private Declare Function mmioGetInfo Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef lpmmioinfo As Any, _
                         ByVal wFlags As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
                         ByVal wFlags As Long, _
                         ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" _
                         Alias "VarPtr" ( _
                         ByRef pArr() As Any) As Long
Private Declare Function UuidFromString Lib "rpcrt4" _
                         Alias "UuidFromStringW" ( _
                         ByVal pStringUuid As Long, _
                         ByRef pUuid As UUID) As Long
Private Declare Function IsEqualGUID Lib "ole32" ( _
                         ByRef rguid1 As UUID, _
                         ByRef rguid2 As UUID) As Long
Private Declare Function PlaySound Lib "winmm.dll" _
                         Alias "PlaySoundW" ( _
                         ByRef lpszName As Any, _
                         ByVal hModule As Long, _
                         ByVal dwFlags As Long) As Long

Private Declare Sub PutArr Lib "msvbvm60" _
                    Alias "PutMem4" ( _
                    ByRef pArr() As Any, _
                    Optional ByVal pSA As Long = 0)
Private Declare Sub GetMem4 Lib "msvbvm60" ( _
                    ByRef pSrc As Any, _
                    ByRef pDst As Any)
Private Declare Sub MoveArray Lib "msvbvm60" _
                    Alias "__vbaAryMove" ( _
                    ByRef pDestination() As Any, _
                    ByRef pSource() As Any)
                    
Private m_fSamples()    As Single
Private m_lSamples      As Long
Private m_lChannels     As Long
Private m_lSampleRate   As Long

Public Property Get SampleRate() As Long
    SampleRate = m_lSampleRate
End Property

Public Property Let SampleRate( _
                    ByVal lValue As Long)
                    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "SampleRate_put"
    
    If lValue <= 0 Or lValue > 1000000 Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    m_lSampleRate = lValue
    
End Property

Public Property Get SamplesCount() As Long
    SamplesCount = m_lSamples
End Property

Public Property Let SamplesCount( _
                    ByVal lValue As Long)
                    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "SamplesCount_put"
    
    If lValue < 0 Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    If lValue > 0 Then
    
        If m_lChannels = 0 Then
            m_lChannels = 1
        End If
    
        ReDim Preserve m_fSamples(m_lChannels - 1, lValue - 1)
     
    Else
        Erase m_fSamples
    End If
    
    m_lSamples = lValue
    
End Property

Public Property Get Channels() As Long
    Channels = m_lChannels
End Property

Public Property Get Duration() As Double
    If m_lSampleRate > 0 Then
        Duration = m_lSamples / m_lSampleRate
    End If
End Property

' // Play all as 16-bit
Public Sub Play( _
           ByVal eChannelsMask As eChannelsMask, _
           ByVal lStartSample As Long, _
           ByVal lCountOfSamples As Long)
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "Play"
    
    Dim tFileHeader As tTestWaveExFile
    Dim iFileData() As Integer
    Dim lIndex      As Long
    Dim bChUsage(7) As Boolean
    Dim lChIndex    As Long
    Dim lDstIndex   As Long
    Dim fValue      As Single
    
    If lStartSample < 0 Or lCountOfSamples < 0 Or lStartSample + lCountOfSamples > m_lSamples Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    For lIndex = 0 To 7
        bChUsage(lIndex) = eChannelsMask And 1
        eChannelsMask = eChannelsMask \ 2
    Next
    
    With tFileHeader
    
        .lRIFFID = FOURCC_RIFF
        .lRIFFSize = Len(tFileHeader) - 8 + lCountOfSamples * 2 * m_lChannels
        .lDataTypeID = FOURCC_WAVE
        .lFmtCunkID = FOURCC_FMT
        .lFmtChunkSize = 40
        .wFormatTag = WAVE_FORMAT_EXTENSIBLE
        .wNumOfChannels = m_lChannels
        .wBitsPerSample = 16
        .nSamplesPerSec = m_lSampleRate
        .nBlockAlign = (.wBitsPerSample \ 8) * .wNumOfChannels
        .nAvgBytesPerSec = .nBlockAlign * .nSamplesPerSec
        .cbSize = 22
        .wValidBits = 16
        
        If m_lChannels = 1 Then
            .dwChannelMask = 4  ' // SPEAKER_FRONT_CENTER
        Else
            .dwChannelMask = (2 ^ m_lChannels) - 1
        End If
        
        .guidSubFormat = KSDATAFORMAT_SUBTYPE_PCM
        .lDataChunkID = FOURCC_DATA
        .lDataChunkSize = lCountOfSamples * 2 * m_lChannels
        
    End With
    
    ReDim iFileData(Len(tFileHeader) \ 2 + lCountOfSamples * m_lChannels - 1)
    
    memcpy iFileData(0), tFileHeader, Len(tFileHeader)
    
    lDstIndex = Len(tFileHeader) \ 2
    
    For lIndex = 0 To lCountOfSamples - 1
        For lChIndex = 0 To m_lChannels - 1
        
            If bChUsage(lChIndex) Then
            
                fValue = m_fSamples(lChIndex, lIndex)
                
                If fValue > 1 Then
                    fValue = 1
                ElseIf fValue < -1 Then
                    fValue = -1
                End If
                
                iFileData(lDstIndex) = fValue * 32767
                
            End If
            
            lDstIndex = lDstIndex + 1
            
        Next
    Next

    If PlaySound(iFileData(0), 0, SND_MEMORY Or SND_SYNC) = 0 Then
        Err.Raise 7, FULL_PROC_NAME
    End If
    
End Sub

' // Mix channels and play them as 16-bit mono
Public Sub PlayChannels( _
           ByVal eChannelsMask As eChannelsMask, _
           ByVal lStartSample As Long, _
           ByVal lCountOfSamples As Long)
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "PlayChannels"
    
    Dim tFileHeader     As tTestWaveFile
    Dim iFileData()     As Integer
    Dim lIndex          As Long
    Dim lChannels(7)    As Long
    Dim lChannelIndex   As Long
    Dim lMixCount       As Long
    Dim fValue          As Single
    
    If lStartSample < 0 Or lCountOfSamples < 0 Or lStartSample + lCountOfSamples > m_lSamples Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    For lIndex = 0 To m_lChannels - 1
    
        If (eChannelsMask And 1) Then
            lChannels(lChannelIndex) = lIndex
            lChannelIndex = lChannelIndex + 1
        End If
        
        eChannelsMask = eChannelsMask \ 2
        
    Next

    lMixCount = lChannelIndex

    With tFileHeader
        .lRIFFID = FOURCC_RIFF
        .lRIFFSize = Len(tFileHeader) - 8 + lCountOfSamples * 2
        .lDataTypeID = FOURCC_WAVE
        .lFmtCunkID = FOURCC_FMT
        .lFmtChunkSize = &H10
        .wFormatTag = WAVE_FORMAT_PCM
        .wNumOfChannels = 1
        .wBitsPerSample = 16
        .nSamplesPerSec = m_lSampleRate
        .nBlockAlign = .wBitsPerSample \ 8 * .wNumOfChannels
        .nAvgBytesPerSec = .nBlockAlign * .nSamplesPerSec
        .lDataChunkID = FOURCC_DATA
        .lDataChunkSize = lCountOfSamples * 2
    End With
    
    ReDim iFileData(Len(tFileHeader) \ 2 + lCountOfSamples - 1)
    
    memcpy iFileData(0), tFileHeader, Len(tFileHeader)
    
    If lMixCount > 0 Then

        For lIndex = 0 To lCountOfSamples - 1
            
            fValue = 0
            
            For lChannelIndex = 0 To lMixCount - 1
                fValue = fValue + m_fSamples(lChannels(lChannelIndex), lIndex + lStartSample)
            Next

            fValue = fValue / lMixCount
            
            If fValue > 1 Then
                fValue = 1
            ElseIf fValue < -1 Then
                fValue = -1
            End If
            
            iFileData(lIndex + Len(tFileHeader) \ 2) = fValue * 32767
            
        Next

    End If
    
    If PlaySound(iFileData(0), 0, SND_MEMORY Or SND_SYNC) = 0 Then
        Err.Raise 7, FULL_PROC_NAME
    End If
    
End Sub

Public Sub AddChannel( _
           Optional ByVal lIndex As Long = -1)
           
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "AddChannel"
    
    Dim fResult()   As Single
    Dim lSmpIndex   As Long
    Dim lInChIndex  As Long
    Dim lOutChIndex As Long
    
    If lIndex >= 8 Or lIndex < -1 Then
        Err.Raise 5, FULL_PROC_NAME
    ElseIf m_lChannels >= 8 Then
        Err.Raise 9, FULL_PROC_NAME
    End If

    If lIndex = -1 Then
        lIndex = m_lChannels
    End If
    
    If m_lSamples > 0 Then
        
        ReDim fResult(m_lChannels, m_lSamples - 1)
        
        For lSmpIndex = 0 To m_lSamples - 1
            
            lInChIndex = 0
            
            For lOutChIndex = 0 To m_lChannels
                If lOutChIndex <> lIndex Then
                    fResult(lOutChIndex, lSmpIndex) = m_fSamples(lInChIndex, lSmpIndex)
                    lInChIndex = lInChIndex + 1
                End If
            Next
            
        Next
        
        MoveArray m_fSamples, fResult
        
    End If

    m_lChannels = m_lChannels + 1
    
End Sub

Public Sub RemoveChannel( _
           ByVal lIndex As Long)
           
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "RemoveChannel"
    
    Dim fResult()   As Single
    Dim lSmpIndex   As Long
    Dim lInChIndex  As Long
    Dim lOutChIndex As Long
    
    If lIndex < 0 Or lIndex >= m_lChannels Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    If m_lSamples > 0 Then
        
        If m_lChannels > 1 Then
            
            ReDim fResult(m_lChannels - 2, m_lSamples - 1)
            
            For lSmpIndex = 0 To m_lSamples - 1
                
                lInChIndex = 0
                
                For lOutChIndex = 0 To m_lChannels - 2
                
                    If lOutChIndex = lIndex Then
                        lInChIndex = lInChIndex + 1
                    End If
                    
                    fResult(lOutChIndex, lSmpIndex) = m_fSamples(lInChIndex, lSmpIndex)
                        
                    lInChIndex = lInChIndex + 1
                    
                Next
                
            Next
            
            MoveArray m_fSamples, fResult
            
        Else
            Erase m_fSamples
            m_lSamples = 0
        End If
        
    End If

    m_lChannels = m_lChannels - 1
    
End Sub

' // Get specified channel data
Public Property Get Channel( _
                    ByVal lChannelIndex As Long, _
                    ByVal lStartSample As Long, _
                    ByVal lCountOfSamples As Long) As Single()
                    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "Channel_get"
    
    Dim fRet()  As Single
    Dim lIndex  As Long
    
    If lStartSample < 0 Or lCountOfSamples < 0 Or lStartSample + lCountOfSamples > m_lSamples Or _
        lChannelIndex < 0 Or lChannelIndex >= m_lChannels Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    If lCountOfSamples Then
        
        ReDim fRet(lCountOfSamples - 1)
        
        For lIndex = 0 To lCountOfSamples - 1
            fRet(lIndex) = m_fSamples(lChannelIndex, lIndex + lStartSample)
        Next
        
        Channel = fRet
        
    End If
    
End Property

' // Mix all the channels using channel mask (bitset corresponds to mixed channels)
Public Property Get Mix( _
                    ByVal eChannelsMask As eChannelsMask, _
                    ByVal lStartSample As Long, _
                    ByVal lCountOfSamples As Long) As Single()
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "Mix_get"
    
    Dim fRet()          As Single
    Dim lIndex          As Long
    Dim lChannels(7)    As Long
    Dim lChannelIndex   As Long
    Dim lMixCount       As Long
    
    If lStartSample < 0 Or lCountOfSamples < 0 Or lStartSample + lCountOfSamples > m_lSamples Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    If lCountOfSamples Then
        
        For lIndex = 0 To m_lChannels - 1
        
            If (eChannelsMask And 1) Then
                lChannels(lChannelIndex) = lIndex
                lChannelIndex = lChannelIndex + 1
            End If
            
            eChannelsMask = eChannelsMask \ 2
            
        Next
        
        ReDim fRet(lCountOfSamples - 1)
        
        lMixCount = lChannelIndex
        
        If lMixCount > 0 Then

            For lIndex = 0 To lCountOfSamples - 1
            
                For lChannelIndex = 0 To lMixCount - 1
                    fRet(lIndex) = fRet(lIndex) + m_fSamples(lChannels(lChannelIndex), lIndex + lStartSample)
                Next
                
                If lMixCount > 1 Then
                    fRet(lIndex) = fRet(lIndex) / lMixCount
                End If
                
            Next

        End If
        
        Mix = fRet
        
    End If
    
End Property

' // Set the specified channel data
Public Property Let Channel( _
                    ByVal lChannelIndex As Long, _
                    ByVal lStartSample As Long, _
                    ByVal lCountOfSamples As Long, _
                    ByRef fSamples() As Single)
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "Channel_put"
    
    Dim lIndex  As Long
    Dim pSA     As Long
    
    If lStartSample < 0 Or lCountOfSamples < 0 Or lStartSample + lCountOfSamples > m_lSamples Or _
        lChannelIndex < 0 Or lChannelIndex >= m_lChannels Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    If lCountOfSamples Then
    
        GetMem4 ByVal ArrPtr(fSamples), pSA
        
        If pSA = 0 Then
            Err.Raise 5, FULL_PROC_NAME
        ElseIf UBound(fSamples) < lCountOfSamples - 1 Then
            Err.Raise 5, FULL_PROC_NAME
        End If
        
        For lIndex = 0 To lCountOfSamples - 1
            m_fSamples(lChannelIndex, lIndex + lStartSample) = fSamples(lIndex)
        Next
        
    End If
    
End Property

Public Sub InitNew( _
           ByVal lNumberOfChannels As Long, _
           ByVal lNumberOfSamples As Long, _
           ByVal lSampleRate As Long)
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "InitNew"
    
    If lNumberOfChannels <= 0 Or lNumberOfChannels > 8 Or lNumberOfSamples < 0 Or _
        lSampleRate <= 0 Or lSampleRate > 1000000 Then
        Err.Raise 5, FULL_PROC_NAME
    End If
    
    If lNumberOfSamples Then
        ReDim m_fSamples(lNumberOfChannels - 1, lNumberOfSamples - 1)
    Else
        Erase m_fSamples
    End If
     
    m_lChannels = lNumberOfChannels
    m_lSamples = lNumberOfSamples
    m_lSampleRate = lSampleRate
    
End Sub

Public Sub Save( _
           ByRef sFileName As String, _
           ByVal lBitsPerSample As Long, _
           Optional ByVal bUseFloatFor32bit As Boolean)
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "Save"
    
    Dim hMMFile As OLE_HANDLE
    
    hMMFile = mmioOpen(StrPtr(sFileName), ByVal 0&, MMIO_READWRITE Or MMIO_CREATE)
    If hMMFile = 0 Then
        Err.Raise 7, FULL_PROC_NAME, "mmioOpen failed"
    End If
    
    On Error GoTo error_handler

    SaveInternal hMMFile, lBitsPerSample, bUseFloatFor32bit
    
    mmioClose hMMFile
    
    Exit Sub
    
error_handler:
    
    If hMMFile Then
        mmioClose hMMFile
    End If
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Public Function SaveToMemory( _
                ByVal lBitsPerSample As Long, _
                Optional ByVal bUseFloatFor32bit As Boolean) As Byte()
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "SaveToMemory"
    
    Dim hMMFile As OLE_HANDLE
    Dim tIoInfo As MMIOINFO
    Dim bOut()  As Byte
    
    With tIoInfo
        .fccIOProc = FOURCC_MEM
        .cchBuffer = &H10000
        .adwInfo(0) = &H80000
    End With
    
    hMMFile = mmioOpen(0, tIoInfo, MMIO_READWRITE Or MMIO_CREATE)
    If hMMFile = 0 Then
        Err.Raise 7, FULL_PROC_NAME, "mmioOpen failed"
    End If
    
    On Error GoTo error_handler

    SaveInternal hMMFile, lBitsPerSample, bUseFloatFor32bit
    
    On Error GoTo 0
    
    If mmioGetInfo(hMMFile, tIoInfo, 0) Then
        mmioClose hMMFile
        Err.Raise 7, FULL_PROC_NAME, "Unable to get stream info"
    End If
    
    If tIoInfo.pchNext > tIoInfo.pchBuffer Then
    
        ReDim bOut(tIoInfo.pchNext - tIoInfo.pchBuffer - 1)
    
        memcpy bOut(0), ByVal tIoInfo.pchBuffer, UBound(bOut) + 1

    End If
    
    mmioClose hMMFile
    
    SaveToMemory = bOut
    
    Exit Function
    
error_handler:
    
    If hMMFile Then
        mmioClose hMMFile
    End If
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Public Sub Load( _
           ByRef sFileName As String)
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "Load"
    
    Dim hMMFile As OLE_HANDLE
    
    hMMFile = mmioOpen(StrPtr(sFileName), ByVal 0&, MMIO_READWRITE)
    If hMMFile = 0 Then
        Err.Raise 7, FULL_PROC_NAME, "mmioOpen failed"
    End If
    
    On Error GoTo error_handler

    LoadInternal hMMFile
    
    mmioClose hMMFile
    
    Exit Sub
    
error_handler:
    
    If hMMFile Then
        mmioClose hMMFile
    End If
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Public Sub LoadFromMemory( _
           ByVal pData As Long, _
           ByVal lSize As Long)
    
    Const FULL_PROC_NAME As String = MODULE_NAME & "::" & "LoadFromMemory"
    
    Dim hMMFile As OLE_HANDLE
    Dim tIoInfo As MMIOINFO
    
    With tIoInfo
        .fccIOProc = FOURCC_MEM
        .cchBuffer = lSize
        .pchBuffer = pData
    End With
    
    hMMFile = mmioOpen(0, tIoInfo, MMIO_READWRITE)
    If hMMFile = 0 Then
        Err.Raise 7, FULL_PROC_NAME, "mmioOpen failed"
    End If
    
    On Error GoTo error_handler

    LoadInternal hMMFile
    
    mmioClose hMMFile
    
    Exit Sub
    
error_handler:
    
    If hMMFile Then
        mmioClose hMMFile
    End If
    
    Err.Raise Err.Number, FULL_PROC_NAME, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Private Sub SaveInternal( _
            ByVal hMMFile As OLE_HANDLE, _
            ByVal lBitsPerSample As Long, _
            ByVal bFloat As Boolean)
    Dim tckRIFF     As MMCKINFO
    Dim tckWAVE     As MMCKINFO
    Dim tckFMT      As MMCKINFO
    Dim tckDATA     As MMCKINFO
    Dim tFMT        As WAVEFORMATEXTENSIBLE
    Dim lFmtSize    As Long
    Dim lSmpIndex   As Long
    Dim lChIndex    As Long
    Dim lDstIndex   As Long
    Dim bArr()      As Byte
    Dim iArr()      As Integer
    Dim lArr()      As Long
    Dim fValue      As Single
    Dim lTemp       As Long
    Dim pData       As Long
    Dim lDataSize   As Long
    
    If (lBitsPerSample <> 8 And lBitsPerSample <> 16 And lBitsPerSample <> 24 And lBitsPerSample <> 32) Or _
        (lBitsPerSample <> 32 And bFloat) Then
        Err.Raise 5
    ElseIf m_lChannels = 0 Then
        Err.Raise 321
    End If
    
    tckRIFF.fccType = mmioStringToFOURCC("WAVE", 0)
    
    If mmioCreateChunk(hMMFile, tckRIFF, MMIO_CREATERIFF) Then
        Err.Raise 7, , "Unable to create RIFF chunk"
    End If
    
    tckFMT.ckid = mmioStringToFOURCC("fmt", 0)
    
    With tFMT
        
        .nChannels = m_lChannels
        .nSamplesPerSec = m_lSampleRate
        .wBitsPerSample = lBitsPerSample
        .nBlockAlign = (lBitsPerSample \ 8) * m_lChannels
        .nAvgBytesPerSec = .nBlockAlign * m_lSampleRate
        
        If m_lChannels <= 2 Then
            ' // WAVEFORMATEX
            
            If bFloat Then
                .wFormatTag = WAVE_FORMAT_IEEE_FLOAT
            Else
                .wFormatTag = WAVE_FORMAT_PCM
            End If
            
            lFmtSize = 16
            
        Else
            ' // WAVEFORMATEXTENSIBLE
            
            .cbSize = 22
            .wFormatTag = WAVE_FORMAT_EXTENSIBLE
            .dwChannelMask = (2 ^ m_lChannels) - 1
            .wValidBitsPerSample = lBitsPerSample
            
            If bFloat Then
                .guidSubFormat = KSDATAFORMAT_SUBTYPE_IEEE_FLOAT
            Else
                .guidSubFormat = KSDATAFORMAT_SUBTYPE_PCM
            End If
            
            lFmtSize = 40
            
        End If
        
    End With
    
    If mmioCreateChunk(hMMFile, tckFMT, 0) Then
        Err.Raise 7, , "Unable to create fmt chunk"
    End If
    
    If mmioWrite(hMMFile, tFMT, lFmtSize) = -1 Then
        Err.Raise 7, , "Unable to write fmt chunk"
    End If
    
    mmioAscend hMMFile, tckFMT, 0
    
    tckDATA.ckid = mmioStringToFOURCC("data", 0)
    
    If mmioCreateChunk(hMMFile, tckDATA, 0) Then
        Err.Raise 7, , "Unable to create data chunk"
    End If
    
    If m_lSamples Then
        
        Select Case lBitsPerSample
        Case 8
        
            ReDim bArr(m_lChannels - 1, m_lSamples - 1)
             
            For lSmpIndex = 0 To m_lSamples - 1
                For lChIndex = 0 To m_lChannels - 1
                    
                    fValue = m_fSamples(lChIndex, lSmpIndex)
                    
                    If fValue > 1 Then
                        fValue = 1
                    ElseIf fValue < -1 Then
                        fValue = -1
                    End If
                    
                    bArr(lChIndex, lSmpIndex) = fValue * 127 + 127
                    
                Next
            Next
            
            pData = VarPtr(bArr(0, 0))
            lDataSize = m_lChannels * m_lSamples
             
        Case 16
        
            ReDim iArr(m_lChannels - 1, m_lSamples - 1)
             
            For lSmpIndex = 0 To m_lSamples - 1
                For lChIndex = 0 To m_lChannels - 1
                    
                    fValue = m_fSamples(lChIndex, lSmpIndex)
                    
                    If fValue > 1 Then
                        fValue = 1
                    ElseIf fValue < -1 Then
                        fValue = -1
                    End If
                    
                    iArr(lChIndex, lSmpIndex) = fValue * 32767
                    
                Next
            Next
            
            pData = VarPtr(iArr(0, 0))
            lDataSize = m_lChannels * m_lSamples * 2
            
        Case 24
        
            ReDim bArr(m_lChannels * m_lSamples * 3 - 1)
             
            For lSmpIndex = 0 To m_lSamples - 1
                For lChIndex = 0 To m_lChannels - 1
                    
                    fValue = m_fSamples(lChIndex, lSmpIndex)
                    
                    If fValue > 1 Then
                        fValue = 1
                    ElseIf fValue < -1 Then
                        fValue = -1
                    End If
                    
                    lTemp = fValue * &H7FFFFF
                                        
                    bArr(lDstIndex) = lTemp And &HFF&
                    bArr(lDstIndex + 1) = (lTemp And &HFF00&) \ &H100
                    bArr(lDstIndex + 2) = (lTemp And &HFF0000) \ &H10000
                    
                    lDstIndex = lDstIndex + 3
                    
                Next
            Next
            
            pData = VarPtr(bArr(0))
            lDataSize = m_lChannels * m_lSamples * 3
            
        Case 32
        
            If bFloat Then
                pData = VarPtr(m_fSamples(0, 0))
            Else
            
                ReDim lArr(m_lChannels - 1, m_lSamples - 1)
                 
                For lSmpIndex = 0 To m_lSamples - 1
                    For lChIndex = 0 To m_lChannels - 1
                        
                        fValue = m_fSamples(lChIndex, lSmpIndex)
                        
                        If fValue > 1 Then
                            fValue = 1
                        ElseIf fValue < -1 Then
                            fValue = -1
                        End If
                        
                        lArr(lChIndex, lSmpIndex) = fValue * 2147483647
                        
                    Next
                Next
                
                pData = VarPtr(lArr(0, 0))

            End If
            
            lDataSize = m_lChannels * m_lSamples * 4
            
        End Select
            
        If mmioWrite(hMMFile, ByVal pData, lDataSize) = -1 Then
            Err.Raise 7, , "Unable to write data chunk"
        End If
    
    End If
    
    mmioAscend hMMFile, tckDATA, 0
    mmioAscend hMMFile, tckRIFF, 0

End Sub

Private Sub LoadInternal( _
            ByVal hMMFile As OLE_HANDLE)
    Dim tckRIFF     As MMCKINFO
    Dim tckWAVE     As MMCKINFO
    Dim tckFMT      As MMCKINFO
    Dim tckDATA     As MMCKINFO
    Dim tFMT        As WAVEFORMATEXTENSIBLE
    Dim lInIndex    As Long
    Dim lOutIndex   As Long
    Dim lOutSamples As Long
    Dim fSample     As Single
    Dim pRawBytes   As Long
    Dim bArr()      As Byte
    Dim iArr()      As Integer
    Dim lArr()      As Long
    Dim tArrDesc    As SAFEARRAY
    Dim pSafeArray  As Long
    Dim lTemp       As Long
    Dim lChIndex    As Long
    Dim lDivisor    As Long
    Dim bIsFloat    As Boolean
    
    tckWAVE.fccType = mmioStringToFOURCC("WAVE", 0)

    If mmioDescend(hMMFile, tckWAVE, ByVal 0&, MMIO_FINDRIFF) Then
        Err.Raise 321, , "Isn't valid file"
    End If
    
    tckFMT.ckid = mmioStringToFOURCC("fmt", 0)
    
    If mmioDescend(hMMFile, tckFMT, tckWAVE, MMIO_FINDCHUNK) Then
        Err.Raise 321, , "Format chunk not found"
    End If
    
    If tckFMT.ckSize < 0 Then
        Err.Raise 321, , "Invalid format"
    End If
    
    ReDim bFMT(tckFMT.ckSize - 1)
    
    If mmioRead(hMMFile, bFMT(0), tckFMT.ckSize) = -1 Then
        Err.Raise 7, , "Can't read format"
    End If
    
    mmioAscend hMMFile, tckFMT, 0
    
    tckDATA.ckid = mmioStringToFOURCC("data", 0)

    If mmioDescend(hMMFile, tckDATA, tckWAVE, MMIO_FINDCHUNK) Then
        Err.Raise 321, , "Wave data isn't found"
    End If
    
    If tckDATA.ckSize <= 0 Then
        Err.Raise 321, , "Invalid data size"
    End If
    
    If tckFMT.ckSize > Len(tFMT) Then
        tckFMT.ckSize = Len(tFMT)
    End If
    
    memcpy tFMT, bFMT(0), tckFMT.ckSize
    
    If ((tFMT.wFormatTag <> WAVE_FORMAT_PCM And tFMT.wFormatTag <> WAVE_FORMAT_EXTENSIBLE And _
        tFMT.wFormatTag <> WAVE_FORMAT_IEEE_FLOAT) Or _
        tFMT.nChannels > 8 Or tFMT.nChannels <= 0 Or tFMT.nBlockAlign <> tFMT.wBitsPerSample * tFMT.nChannels \ 8) Or _
        Not (tFMT.wBitsPerSample = 8 Or tFMT.wBitsPerSample = 16 Or tFMT.wBitsPerSample = 24 Or tFMT.wBitsPerSample = 32) Then
        Err.Raise 321, , "Unsupported format"
    End If
    
    If tFMT.wFormatTag = WAVE_FORMAT_EXTENSIBLE Then
        
        If tFMT.wValidBitsPerSample <= 0 Or tFMT.wValidBitsPerSample > tFMT.wBitsPerSample Then
            Err.Raise 321, , "Invalid file format"
        End If
        
        If tFMT.cbSize <> 22 Then
            Err.Raise 321, , "Unsupported format"
        ElseIf IsEqualGUID(tFMT.guidSubFormat, KSDATAFORMAT_SUBTYPE_PCM) = 0 Then
        
            If IsEqualGUID(tFMT.guidSubFormat, KSDATAFORMAT_SUBTYPE_IEEE_FLOAT) = 0 Then
                Err.Raise 321, , "Unsupported format"
            End If
            
            bIsFloat = True
            
        End If
        
    Else
    
        If tFMT.cbSize Then
            Err.Raise 321, , "Invalid file format"
        End If
        
        tFMT.wValidBitsPerSample = tFMT.wBitsPerSample
        
        bIsFloat = tFMT.wFormatTag = WAVE_FORMAT_IEEE_FLOAT
        
    End If
    
    pRawBytes = GlobalAlloc(GMEM_FIXED, ((tckDATA.ckSize + 3) \ 4) * 4)
    If pRawBytes = 0 Then
        Err.Raise 7, , "GlobalAlloc failed"
    End If
    
    If mmioRead(hMMFile, ByVal pRawBytes, tckDATA.ckSize) = -1 Then
        GlobalFree pRawBytes
        Err.Raise 7, , "Unable to read wave data"
    End If

    lOutSamples = tckDATA.ckSize \ tFMT.nBlockAlign
    
    If lOutSamples > 0 Then
                
        ReDim m_fSamples(tFMT.nChannels - 1, lOutSamples - 1)
        
        pSafeArray = VarPtr(tArrDesc)

        tArrDesc.cDims = 1
        tArrDesc.fFeatures = FADF_AUTO
        tArrDesc.pvData = pRawBytes
        
        If tFMT.wValidBitsPerSample Then
            lDivisor = (2 ^ (tFMT.wValidBitsPerSample - 1)) - 1
        Else
            lDivisor = 1
        End If
        
        Select Case tFMT.wBitsPerSample
        Case 8
            
            tArrDesc.cbElements = 1
            tArrDesc.Bounds.cElements = lOutSamples * tFMT.nChannels

            PutArr bArr, pSafeArray
            
            For lOutIndex = 0 To lOutSamples - 1
                For lChIndex = 0 To tFMT.nChannels - 1
                    m_fSamples(lChIndex, lOutIndex) = (CLng(bArr(lInIndex)) - 128&) / lDivisor
                    lInIndex = lInIndex + 1
                Next
            Next
            
            PutArr bArr, 0
            
        Case 16
        
            tArrDesc.cbElements = 2
            tArrDesc.Bounds.cElements = lOutSamples * tFMT.nChannels
            
            PutArr iArr, pSafeArray
            
            For lOutIndex = 0 To lOutSamples - 1
                For lChIndex = 0 To tFMT.nChannels - 1
                    m_fSamples(lChIndex, lOutIndex) = iArr(lInIndex) / lDivisor
                    lInIndex = lInIndex + 1
                Next
            Next
            
            PutArr iArr, 0
            
        Case 24
        
            tArrDesc.cbElements = 1
            tArrDesc.Bounds.cElements = lOutSamples * 3 * tFMT.nChannels
            
            PutArr bArr, pSafeArray

            For lOutIndex = 0 To lOutSamples - 1
                For lChIndex = 0 To tFMT.nChannels - 1
                    
                    lTemp = CLng(bArr(lInIndex * 3)) Or _
                             (bArr(lInIndex * 3 + 1) * &H100&) Or _
                             (bArr(lInIndex * 3 + 2)) * &H10000 Or _
                             -(bArr(lInIndex * 3 + 2) And &H80) * &H20000
                    m_fSamples(lChIndex, lOutIndex) = lTemp / lDivisor
                    lInIndex = lInIndex + 1
                    
                Next
            Next
            
            PutArr bArr, 0
            
        Case 32
            
            tArrDesc.cbElements = 4
            tArrDesc.Bounds.cElements = lOutSamples * tFMT.nChannels
            
            If Not bIsFloat Then
                
                PutArr lArr, pSafeArray
                
                For lOutIndex = 0 To lOutSamples - 1
                    For lChIndex = 0 To tFMT.nChannels - 1
                        m_fSamples(lChIndex, lOutIndex) = lArr(lInIndex) / lDivisor
                        lInIndex = lInIndex + 1
                    Next
                Next
                
                PutArr lArr, 0
                
            Else
                memcpy m_fSamples(0, 0), ByVal pRawBytes, lOutSamples * tFMT.nChannels * 4
            End If

        End Select
    Else
        Erase m_fSamples
    End If
    
    GlobalFree pRawBytes
    
    m_lSamples = lOutSamples
    m_lSampleRate = tFMT.nSamplesPerSec
    m_lChannels = tFMT.nChannels
    
End Sub

Private Function KSDATAFORMAT_SUBTYPE_PCM() As UUID
    Static s_tCache As UUID
    
    If s_tCache.Data1 = 0 Then
        UuidFromString StrPtr(KSDATAFORMAT_SUBTYPE_PCM_STR), s_tCache
    End If
    
    KSDATAFORMAT_SUBTYPE_PCM = s_tCache
    
End Function

Private Function KSDATAFORMAT_SUBTYPE_IEEE_FLOAT() As UUID
    Static s_tCache As UUID
    
    If s_tCache.Data1 = 0 Then
        UuidFromString StrPtr(KSDATAFORMAT_SUBTYPE_IEEE_FLOAT_STR), s_tCache
    End If
    
    KSDATAFORMAT_SUBTYPE_IEEE_FLOAT = s_tCache
    
End Function
