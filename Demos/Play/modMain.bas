Attribute VB_Name = "modMain"
' //
' // Play a file
' //

Option Explicit

Sub Main()
    Dim cFile   As CWaveFile
    
    Set cFile = New CWaveFile
    
    cFile.Load App.Path & "\file.wav"
    
    ' // You can specify channels to play using OR
    cFile.Play CM_0 Or CM_1, 0, cFile.SamplesCount
    
    
End Sub
