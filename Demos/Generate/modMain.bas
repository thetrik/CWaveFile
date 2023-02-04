Attribute VB_Name = "modMain"
' //
' // Pure tone generate/play/save using CWaveFile class
' //

Option Explicit

Const PI As Double = 3.14159265358979

Sub Main()
    Dim cFile       As CWaveFile
    Dim fSamples()  As Single
    Dim lIndex      As Long
    Dim lSampleRate As Long
    Dim dDelta      As Double
    
    Set cFile = New CWaveFile
    
    lSampleRate = 22050
    
    ' // Initialize sound with 22050 Hz 2 seconds
    cFile.InitNew 1, lSampleRate * 2, lSampleRate
    
    ' // Generate 1000 Hz sine wave
    ReDim fSamples(lSampleRate * 2 - 1)
    
    dDelta = 1000 / lSampleRate * PI * 2
    
    For lIndex = 0 To UBound(fSamples)
        fSamples(lIndex) = Sin(lIndex * dDelta)
    Next
    
    ' // Set data
    cFile.Channel(0, 0, UBound(fSamples) + 1) = fSamples
    
    ' // Play
    cFile.Play CM_ALL, 0, cFile.SamplesCount
    
    ' // Save
    cFile.Save App.Path & "\test.wav", 8
    
End Sub
