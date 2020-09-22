VERSION 5.00
Begin VB.Form frmLoad 
   Caption         =   "Select location of files"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   Icon            =   "Load.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "The following are valid commands while watching the slideshow:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   6255
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "F7:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   24
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "F8:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   23
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Right Arrow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   22
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Left Arrow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   21
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Enter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Space bar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Up Arrow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Down Arrow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "F6:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   16
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esc:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Skip Forward"
         Height          =   195
         Left            =   4560
         TabIndex        =   14
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Skip Backward"
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Decrease Slide delay"
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Increase Slide delay"
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pause Slideshow"
         Height          =   195
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Create a new slideshow"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Show/Hide Filename"
         Height          =   195
         Left            =   3840
         TabIndex        =   8
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Show/Hide Estimated Duration"
         Height          =   195
         Left            =   3840
         TabIndex        =   7
         Top             =   600
         Width           =   2190
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Show/Hide File Progress"
         Height          =   195
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "End Program"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   6255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    'Dim i As Integer
    'Dim ext As String
    '
    'For i = 0 To File1.ListCount - 1
    '    ext = LCase(Right(File1.List(i), 4))
    '    If ext = ".jpg" Or ext = "jepg" Or ext = ".gif" Or ext = ".bmp" Or ext = ".tif" Or ext = ".jpe" Or ext = ".pcx" Or ext = ".pic" Or ext = ".tga" Or ext = "tiff" Then
    '        frmMain.files.AddItem File1.Path & "\" & File1.List(i)
    '    End If
    'Next i
    
    dirListing Dir1.List(Dir1.ListIndex)
    
    Load frmMain
    frmMain.Show
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
    On Error GoTo Ignoreit
    Dir1.Path = Drive1.List(Drive1.ListIndex)

Ignoreit:
End Sub

