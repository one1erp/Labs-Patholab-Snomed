Attribute VB_Name = "ClueInit"
Option Explicit
Dim xClueApp As Object 'ClueBrowser5.Application
Dim xClueBrw As Object 'ClueBrowser
Dim xClueRef As Object
Dim ControlId As Long



Private Sub ClueInit()
    
Dim Tpause As Double
'Initializes the browser and reference engines
    On Error Resume Next
    Tpause = Timer
    If xClueApp Is Nothing Then
        Set xClueApp = CreateObject("ClueBrowser5.Application")
        xClueApp.ClueBrowser.Hide
'wait here for 2 seconds for the browser to load
'a pause is essential as other calls at this stage may lead to
'incomplete initialisation
        While Abs(Timer - Tpause) < 2
            DoEvents
            DoEvents
            DoEvents
            DoEvents
        Wend
    End If
'wait here for up to 5 seconds for the browser to report that it is available
    While Not xClueApp.Available And Abs(Timer - Tpause) < 5
        DoEvents
        DoEvents
        DoEvents
        DoEvents
    Wend
'handle failure - most likely due to no licence or bad installation or no data
    If Not xClueApp.Available Then
        MsgBox "Cannot access Clue API"
        Exit Sub
    End If
'
    Set xClueBrw = xClueApp.Browser
    xClueBrw.Hide
    Set xClueRef = xClueApp.Reference
'show tool bar and add ok button
    xClueBrw.AddImage "Cancel", "C:\Program Files\cic\CLUE Browser\images\btn_stop.ico"
    xClueBrw.AddImage "Select", "C:\Program Files\cic\CLUE Browser\images\btn_tick.ico"
    xClueBrw.ToolbarVisible = True
    xClueBrw.UserButton key:="OK", Caption:="Select", Tag:="Ok", Style:=0, Image:="Select"
    xClueBrw.UserButton key:="Cancel", Caption:="Cancel", Tag:="Cancel", Style:=0, Image:="Cancel"
End Sub
Public Sub Quit()
    On Error Resume Next
    xClueApp.Quit
    DoEvents
    DoEvents
    Set xClueApp = Nothing
    Set xClueBrw = Nothing
    Set xClueRef = Nothing
End Sub
Public Property Get ClueBrw() As Object 'ClueBrowser
    If xClueBrw Is Nothing Then ClueInit
    Set ClueBrw = xClueBrw
End Property
Public Property Get ClueRef() As Object 'ClueReference
    If xClueRef Is Nothing Then ClueInit
    Set ClueRef = xClueRef
End Property

Public Property Get ClueApp() As Object 'ClueBrowser5.Application
    If xClueApp Is Nothing Then ClueInit
    Set ClueApp = xClueApp
End Property


Public Property Get UserControlId() As Long
    UserControlId = ControlId
End Property

Public Property Let UserControlId(UserControl As Long)
    ControlId = UserControl
End Property

