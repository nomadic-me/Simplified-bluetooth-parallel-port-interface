VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bluetooth <-> Parallel Port Interface"
   ClientHeight    =   2970
   ClientLeft      =   600
   ClientTop       =   840
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Tag             =   "`"
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   4680
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   975
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":0F2B
      ScaleHeight     =   495
      ScaleWidth      =   1935
      TabIndex        =   7
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1920
      Width           =   1035
   End
   Begin VB.CommandButton cmdpath 
      Caption         =   "Refresh File Path"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   4800
      Top             =   120
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":17E1
      Height          =   2415
      Left            =   5760
      TabIndex        =   13
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Instructions :"
      BeginProperty Font 
         Name            =   "CityBlueprint"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "2005 Satish Surath"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label10 
      Caption         =   "©"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "This Programme Was Designed For Demonstration Purpose Only With No Warranty/Liability Either Implied Or Expressed."
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label lblrefresh 
      Caption         =   "refresh"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblheading 
      Caption         =   "Interfacing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pin As Variant
Dim t1 As Variant
Dim t2 As Variant
Dim t3 As Variant
Dim t4 As Variant
Dim t5 As Variant
Dim t6 As Variant
Dim t7 As Variant
Dim t8 As Variant
Dim t9 As Variant
Dim t10 As Variant
Dim binpin As Integer
Dim strfilename As String
Dim m As Variant
Dim flag As Integer
Dim filenum As Variant
Dim tim As Integer
Dim counter As Integer


Private Sub cmdexit_Click()
Call PortOut(888, 0)
End
End Sub

Private Sub cmdpath_Click()
strfilename = Text1.Text
lblrefresh.Caption = "Refreshing system info.."
counter = 6
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
Command2.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Form_Load()
strfilename = "c:\out.por"
counter = 1
Call PortOut(888, 0)
lblrefresh.Caption = "Refreshing Folder info.."
Command2.Enabled = True
Command1.Enabled = False
tim = 1

End Sub


Private Sub Timer1_Timer()

On Error GoTo Trap

filenum = FreeFile

Open strfilename For Input As filenum

m = Input(LOF(filenum), filenum)

        If m = "start" Then
            Call PortOut(888, 1)
            Exit Sub
        Else

            If m = "stop" Then
                    Call PortOut(888, 0)
                    Exit Sub
            Else
                    MsgBox ("Valid values are 'start' And 'stop'. please check if you have entered either of the two case-sensitive strings")
                    Exit Sub
            End If
        
        End If


Close filenum

Kill (strfilename)



Trap:
Exit Sub

End Sub


Private Sub Timer3_Timer()
If counter <= 10 Then
lblrefresh.Caption = ""
counter = counter + 1
End If
If counter = 4 Then
lblrefresh.Caption = "Refreshing system info.."
counter = 0
End If
End Sub

