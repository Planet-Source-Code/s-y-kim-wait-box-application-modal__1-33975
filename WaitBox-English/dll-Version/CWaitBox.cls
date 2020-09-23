VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWaitBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Events
Public Event StartProgress()
Public Event EndProgress()
Public Event WindowStateChange(ByVal State As FormWindowStateConstants)

'Declare the instance of frmWait
Private WithEvents moWaitBox As frmWait
Attribute moWaitBox.VB_VarHelpID = -1

'Local variables for time infos
Private m_lngValue As Long
Private m_lngMinValue As Long
Private m_lngMaxValue As Long
Private m_bErrorFlag As Boolean
Private m_lngPercent As Long
Private m_StartTime As Date
Private m_EndTime As Date
Private m_Elapsed As Date
Private m_Remaining As Date
Private m_bAlreadyStarted  As Boolean
Private m_EndEstimated As Date

Public Property Get StartTime() As String
    'Format the start time
    StartTime = Format(m_StartTime, "hh:mm:ss")
    'You can use the following code
    'StartTime = Format(m_StartTime, "Long Time")
End Property

Public Property Get EndTime() As String
    'Fromat the end time
    EndTime = Format(m_EndTime, "hh:mm:ss")
End Property

Public Property Get EndEstimatedTime() As String
    'Format the estimated end time
    EndEstimatedTime = Format(m_EndEstimated, "hh:mm:ss")
    'You can use the following code
    'EndEstimatedTime = Format(m_EndEstimated, "Long Time")
End Property

Public Property Get ElapsedTime() As String
    'Format the elapased time
    ElapsedTime = Format(m_Elapsed, "hh:mm:ss")
End Property

Public Property Get RemainingTime() As String
    'Format the remaining time
    RemainingTime = Format(m_Remaining, "hh:mm:ss")
End Property


Public Property Get Percent() As Long
    'Calculate the current progress percentage
    Percent = (m_lngValue / (m_lngMaxValue - m_lngMinValue)) * 100
End Property

Public Property Get Caption() As String
    'Return the current caption of WaitBox (frmWait)
    Caption = moWaitBox.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    'Set the caption of WaitBox (frmWait) to a new caption
     moWaitBox.Caption = New_Caption
End Property

Public Property Get Prompt() As String
    'Return the current prompt message on WaitBox (frmWait)
    Prompt = moWaitBox.lblPrompt
End Property

Public Property Let Prompt(ByVal New_Prompt As String)
    'Set the current prompt message of WaitBox (frmWait) to a new prompt
     moWaitBox.lblPrompt = New_Prompt
End Property


Private Sub Class_Initialize()
    'Set the error flag to False
    m_bErrorFlag = False

    'Create a new instance of frmWait
    Set moWaitBox = New frmWait
    
    'Set the min and max values of the progress bar to defaults (0 and 100)
    Min = 0
    Max = 100
    moWaitBox.SetProgMinMax Min, Max
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Unload frmWait and release the reference (moWaitBox) to frmWait
    Terminate
End Sub

Public Sub Terminate()
    On Error Resume Next
    'Unload frmWait
    moWaitBox.Unload
    'Release the reference (moWaitBox) to frmWait
    Set moWaitBox = Nothing
End Sub

Public Sub Show(FrmOwner As Object)
    'If there is any error relating to setting Min, Max, Value
    'do not show the form
    If m_bErrorFlag Then Exit Sub
    
    'Set the min and max values of the progress bar to Min, Max
    'becasue the user has changed them.
    moWaitBox.SetProgMinMax Min, Max
    
    '---------------------------------------------------------------------------
    ' COMMENT: The followings are different from the cls version
    '---------------------------------------------------------------------------
    With moWaitBox
        .Left = FrmOwner.Left + (FrmOwner.Width - .Width) / 2
        .Top = FrmOwner.Top + (FrmOwner.Height - .Height) / 2
    End With
    'Show frmWait in modal mode.
    'This will prevent the owner from being acrtivated.
    moWaitBox.Show vbModal
End Sub


Private Sub moWaitBox_StartProgress()
    'When received the StartProgress event message from frmWait
    'just, raise the same event, so that the user can recognize that
    'it is the time to start his or her work.
    RaiseEvent StartProgress
End Sub

Public Property Get Min() As Long
    'Returns the current Min value (of the progress bar)
    Min = m_lngMinValue
End Property

Public Property Let Min(ByVal New_Value As Long)
    'Set the Min value (of the progress bar)
    If New_Value > Max Then
        MsgBox "The Min progress value must be less than Max." & vbCrLf & _
                vbCrLf & vbTab & "Current Min value:" & Min & _
                vbCrLf & vbTab & "Current Max value:" & Max & _
                vbCrLf & vbTab & "Current value:" & Value, _
                vbCritical, "WaitBox"
        'Flag error
        m_bErrorFlag = True
    ElseIf New_Value < 0 Then
        MsgBox "The Min progress value must be 0 or a positive value." & vbCrLf & _
                vbCrLf & vbTab & "Current Min value:" & Min & _
                vbCrLf & vbTab & "Current Max value:" & Max & _
                vbCrLf & vbTab & "Current value:" & Value, _
                vbCritical, "WaitBox"
        'Flag error
        m_bErrorFlag = True
    Else
        'Unflag error
        m_bErrorFlag = False
    End If
    'If there is an error, exit.
    If m_bErrorFlag Then Exit Property
    'Set the min value to a new value
    m_lngMinValue = New_Value
End Property

Public Property Get Max() As Long
    'Returns the current Min value (of the progress bar)
    Max = m_lngMaxValue
End Property

Public Property Let Max(ByVal New_Value As Long)
    'Set the Max value (of the progress bar)
    If New_Value < Min Then
        MsgBox "The Max progress value must be larger than Min." & vbCrLf & _
                vbCrLf & vbTab & "Current Min value:" & Min & _
                vbCrLf & vbTab & "Current Max value:" & Max & _
                vbCrLf & vbTab & "Current value:" & Value, _
                vbCritical, "WaitBox"
        'Flag error
        m_bErrorFlag = True
    ElseIf New_Value <= 0 Then
        MsgBox "The Max progress value must be a positive value." & vbCrLf & _
                vbCrLf & vbTab & "Current Min value:" & Min & _
                vbCrLf & vbTab & "Current Max value:" & Max & _
                vbCrLf & vbTab & "Current value:" & Value, _
                vbCritical, "WaitBox"
        'Flag error
        m_bErrorFlag = True
    Else
        'Unflag error
        m_bErrorFlag = False
    End If
    'If there is an error, exit.
    If m_bErrorFlag Then Exit Property
    'Set the max value to a new value
    m_lngMaxValue = New_Value
End Property

Public Property Get Value() As Long
    'Returns the current progrees value (of the progress bar)
    Value = m_lngValue
End Property

Public Property Let Value(ByVal New_Value As Long)
    'Set the current progrees value (of the progress bar) to a new value
    'Check the new value in the proper range.
    If New_Value < Min Then
        MsgBox "The progress value must be equal to or larger than Min value." & vbCrLf & _
                vbCrLf & vbTab & "Current Min value:" & Min & _
                vbCrLf & vbTab & "Current Max value:" & Max & _
                vbCrLf & vbTab & "Current value:" & Value, _
                "New value:" & New_Value, vbCritical, "WaitBox"
        'Flag error
        m_bErrorFlag = True
    ElseIf New_Value > Max Then
        'MsgBox "The progress value must be equal to or less than Max value." & vbCrLf & _
        '        vbCrLf & vbTab & "Current Min value:" & Min & _
        '        vbCrLf & vbTab & "Current Max value:" & Max & _
        '        vbCrLf & vbTab & "Current value:" & Value, _
        '        "New value:" & New_Value, vbCritical, "WaitBox"
        New_Value = Max
        'Do not flag error, we wil accept  a value larger than Max value.
        m_bErrorFlag = False
    Else
        'Unflag error
        m_bErrorFlag = False
    End If
    
    'If there is an error, exit.
    If m_bErrorFlag Then Exit Property
    
    'Set the start time
    'Need to set it for the first time and prevent from being setting again.
    'Use m_bAlreadyStarted to do this.
    If Not m_bAlreadyStarted Then
        m_bAlreadyStarted = True
        m_StartTime = Now
        'Show the start time on frmWait
        moWaitBox.lblStartTime = Me.StartTime
    End If
    
    'Set the progress value (of the progress bar) to the new value
    m_lngValue = New_Value
    moWaitBox.SetProgress m_lngValue
    
    'Calcuate  the time infos and display them  on frmWait
    CalculateElapsedTime
    
    'If the current value is equal to Max
    If New_Value >= Max Then
        On Error Resume Next
        'Set the end time
        m_EndTime = Now
        'Terminate
        'Raise the end progress event so that the user can utilize events
        RaiseEvent EndProgress
    End If
End Property

Private Sub CalculateElapsedTime()
    'Calcuate  the time infos and display them  on frmWait
    Dim NowTime As Date
    
    'Save the now time to local variable for the following calculations
    NowTime = Now
    
    'Calculate the elapsed time
    m_Elapsed = NowTime - m_StartTime
    
    'Calcuate the remaining time
    If m_lngValue > m_lngMinValue Then
        m_Remaining = (m_lngMaxValue - m_lngValue) * (m_Elapsed / (m_lngValue - m_lngMinValue))
    Else 'To prevent divide by 0
        m_Remaining = (m_lngMaxValue - m_lngValue) * (m_Elapsed)
    End If
    
    'Calcuate the end estimated time
    m_EndEstimated = NowTime + m_Remaining
    
    'Display time infos on frmWiait
    With moWaitBox
        .lblElapsedTime = Me.ElapsedTime & " (" & DateDiff("s", m_StartTime, NowTime) & " sec" & ")"
        .lblPercent = Me.Percent & " %"
        .lblEndEstimated = Me.EndEstimatedTime
        .lblRemainingTime = Me.RemainingTime & " (" & DateDiff("s", 0, m_Remaining) & " sec" & ")"
        DoEvents
    End With
End Sub

Private Sub moWaitBox_WindowStateChange(ByVal State As FormWindowStateConstants)
    'When the window state of frmWait changes,
    'just notifiy it.
    RaiseEvent WindowStateChange(State)
End Sub


