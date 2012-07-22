VERSION 5.00
Begin VB.Form frmSchoolTime 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "School Time"
   ClientHeight    =   480
   ClientLeft      =   12750
   ClientTop       =   7845
   ClientWidth     =   1725
   BeginProperty Font 
      Name            =   "Helvetica"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchoolTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr_Time 
      Interval        =   1000
      Left            =   240
      Top             =   0
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Helvetica"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1545
      TabIndex        =   0
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "frmSchoolTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
    'Closes the program when double clicked
    End
End Sub

Private Sub lblTime_DblClick()
    'Closes the program when double clicked
    End
End Sub

Private Sub Form_Load()
    
    'Sets form in the bottom right corner of the screen
    Me.Top = Screen.Height - Me.Height - 510
    Me.Left = Screen.Width - Me.Width
    
    'Prints a value before the form is shown, so the user doesn't see a blank window for the first second
    Call tmr_Time_Timer
    
End Sub

Private Sub tmr_Time_Timer()
    Cls
    
    'Declare variables
    'CurrTime is the current time
    Dim CurrTime As Date
    
    'Gets and adjusts time
    CurrTime = DateAdd("s", -30, DateAdd("n", -5, Now))
    
    'Outputs the time
    lblTime.Caption = Format$(CurrTime, "h:mm:ss AMPM")
    
    'Changes the color to green if it is currently green screen, otherwise it stays blue
    If (Hour(CurrTime) = 9 And Minute(CurrTime) >= 40 And Minute(CurrTime) < 45) Or (Hour(CurrTime) = 11 And Minute(CurrTime) < 45) Or (Hour(CurrTime) = 13 And Minute(CurrTime) < 5) Then
        Me.BackColor = &H8000&
    Else
        Me.BackColor = &H80000003
    End If
    
End Sub
