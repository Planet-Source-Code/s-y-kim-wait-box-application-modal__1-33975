VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   Caption         =   "Working......"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6540
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblRemainingTime 
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblEndEstimated 
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblElapsedTime 
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblStartTime 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Elapsed Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Remaining Time:"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "End Time Estimated:"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Start Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Please wait. This will take a couple of minutes."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Event
Public Event StartProgress()
Public Event WindowStateChange(ByVal State As FormWindowStateConstants)

'Lcal variables
Private m_PrevWindowState As FormWindowStateConstants
Private mbAlreadyActivated As Boolean

'API  constants and declares for disabling the system Close button
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Private Declare Function DrawMenuBar Lib "user32" _
      (ByVal hwnd As Long) As Long
     
Private Declare Function GetMenuItemCount Lib "user32" _
      (ByVal hMenu As Long) As Long
     
Private Declare Function GetSystemMenu Lib "user32" _
      (ByVal hwnd As Long, _
       ByVal bRevert As Long) As Long
      
Private Declare Function RemoveMenu Lib "user32" _
      (ByVal hMenu As Long, _
       ByVal nPosition As Long, _
       ByVal wFlags As Long) As Long

Public Sub SetProgMinMax(Min As Long, Max As Long)
    'Set the progress bar Min, Max values
    With Me.ProgressBar1
        .Min = Min
        .Max = Max
    End With
End Sub
Public Sub Unload()
    'Allow the client to unload this form using this method
    VB.Unload Me
End Sub

Public Sub SetProgress(ByVal New_Value As Long)
    'Set the progress to a new value
    Dim iNewValue As Long
    With Me.ProgressBar1
        .Value = New_Value
        .Refresh
    End With
End Sub

Private Sub Form_Activate()
    'We need to raise the StartProgress event
    'when frmWait is shown on the screen and
    ' we have to also ensure that it is raised just one time
    If Not mbAlreadyActivated Then
        mbAlreadyActivated = True
        RaiseEvent StartProgress
    End If
End Sub

Private Sub Form_Resize()
    'When the window state of frmWait changes,
    'Raise the event so that the user can recognize it.
    'This code is not actually necessary!!! - just a try.
    If Me.WindowState <> m_PrevWindowState Then
        RaiseEvent WindowStateChange(Me.WindowState)
        m_PrevWindowState = Me.WindowState
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub Form_Load()
    
   'Set the previous window state to vbNormal
   'for inistializtion
   m_PrevWindowState = vbNormal
   
   'Disable the system Close button.
   'to prevent the user from hitting it.
   Dim hMenu As Long
   Dim menuItemCount As Long

  'Obtain the handle to the form's system menu
   hMenu = GetSystemMenu(Me.hwnd, 0)
 
   If hMenu Then
     
     'Obtain the number of items in the menu
      menuItemCount = GetMenuItemCount(hMenu)
   
     'Remove the system menu Close menu item.
     'The menu item is 0-based, so the last
     'item on the menu is menuItemCount - 1
      Call RemoveMenu(hMenu, menuItemCount - 1, _
                      MF_REMOVE Or MF_BYPOSITION)
  
     'Remove the system menu separator line
      Call RemoveMenu(hMenu, menuItemCount - 2, _
                      MF_REMOVE Or MF_BYPOSITION)
      
     'Remove the system menu Close menu item.
     'The menu item is 0-based, so the last
     'item on the menu is menuItemCount - 1
      'Call RemoveMenu(hMenu, menuItemCount - 4, _
                      MF_REMOVE Or MF_BYPOSITION)
                      
     'Force a redraw of the menu. This
     'refreshes the titlebar, dimming the X
      Call DrawMenuBar(Me.hwnd)

   End If
  
End Sub



