VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Wait Box Example"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton cmdStartWork 
      Caption         =   "Start Work"
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NOTE: This dll test project will not work properly if you run it VB-IDE. You need to compile this project, run the exe separately."
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "You need to make the reference to 'Creative ActiveX  - WaitBox By KSY (English version)' to compile this example."
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Declare the CWaitBox object with WithEvents
'to receive the event message (StartProgress) that indicates
'it is the correct time we start our work.
Private WithEvents moWaitBox As CWaitBox
Attribute moWaitBox.VB_VarHelpID = -1

Private Sub cmdStartWork_Click()
    'Creat a new instance of CWaitBox
    Set moWaitBox = New CWaitBox
    With moWaitBox
        'Set the min and max values of the progess bar
        .Min = 0
        .Max = 200
        'Show the Wait Box
        .Show Me
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Release reference and memory used.
    Set moWaitBox = Nothing
    
    'Unlaod all forms
    Dim Frm As Form
    For Each Frm In Forms
        Unload Frm
        Set Frm = Nothing
    Next
End Sub


Private Sub moWaitBox_StartProgress()
    'When we have received this event message
    'it is the correct time to start our work.
    StartWorks
End Sub

Private Sub StartWorks()
    Dim i As Long
    
    For i = 1 To moWaitBox.Max
        Sleep 200
        DoEvents
        'Change the progress value
        moWaitBox.Value = i
    Next
    
    'Now our work has completed.
    MsgBox "Now, the work has been completed.", vbInformation, "Wait Box"
    
    'Make sure that moWaitBox unload the current instance of frmWait
    'and set the refererernce count to 0
    moWaitBox.Terminate
    
    'Release the our Wait Box object
    Set moWaitBox = Nothing
End Sub

