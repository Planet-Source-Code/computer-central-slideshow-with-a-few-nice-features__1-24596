VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Slideshow Program"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Main.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6465
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer skipTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   5640
   End
   Begin VB.ListBox files 
      Height          =   840
      Left            =   960
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Timer labelTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   5640
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   5640
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3480
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblDuration 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "Slide Timer:  3000 milliseconds."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Image imgPictures 
      Height          =   4455
      Left            =   720
      Top             =   840
      Width           =   7815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim i As Integer
    Dim counter As Integer
    Dim pause As Boolean
    Dim showNum As Boolean
    Dim skipcount As Integer
    Dim showFileName As Boolean
    Dim pass As Integer

Private Sub Form_Click()
    'unhide mouse cursor
    Do
    Loop Until ShowCursor(True) > 5
    
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Timer1.Interval >= 101 Then
            Timer1.Interval = Timer1.Interval - 100
        End If
        lblTimer.Caption = "Slide Timer:  " & Timer1.Interval & " milliseconds."
        lblTimer.Move (Screen.Width - lblTimer.Width) / 2
        lblTimer.Visible = True
        labelTimer.Enabled = True
        counter = 0
        skipTimer.Enabled = True
        skipcount = 0
        calcDuration
    ElseIf KeyCode = vbKeyUp Then
        If Timer1.Interval <= 64000 Then
            Timer1.Interval = Timer1.Interval + 100
        End If
        lblTimer.Caption = "Slide Timer:  " & Timer1.Interval & " milliseconds."
        lblTimer.Move (Screen.Width - lblTimer.Width) / 2
        lblTimer.Visible = True
        labelTimer.Enabled = True
        counter = 0
        skipTimer.Enabled = True
        skipcount = 0
        calcDuration
    ElseIf KeyCode = vbKeyRight Then
        i = i + 1
        lblTimer.Caption = "Displaying picture: " & i + 1 & " of " & files.ListCount & "."
        skipTimer.Enabled = True
        skipcount = 0
        If showFileName = True Then
            lblFilename.Caption = files.List(i)
            lblFilename.Move (Screen.Width - lblFilename.Width) / 2
            lblFilename.Visible = True
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If i > 0 Then
            i = i - 1
            lblTimer.Caption = "Displaying picture: " & i + 1 & " of " & files.ListCount & "."
            skipTimer.Enabled = True
            skipcount = 0
        End If
        If showFileName = True Then
            lblFilename.Caption = files.List(i)
            lblFilename.Move (Screen.Width - lblFilename.Width) / 2
            lblFilename.Visible = True
        End If
    ElseIf KeyCode = vbKeyReturn Then
        'unhide mouse cursor
        Do
        Loop Until ShowCursor(True) > 5
        
        frmLoad.Show
        Unload Me
    ElseIf KeyCode = vbKeySpace Then
        If pause = False Then
            pause = True
            Timer1.Enabled = False
            lblTimer.Caption = "SlideShow Paused."
            lblTimer.Move (Screen.Width - lblTimer.Width) / 2
            lblTimer.Visible = True
        Else
            pause = False
            Timer1.Enabled = True
            lblTimer.Visible = True
            labelTimer.Enabled = True
            counter = 0
        End If
    ElseIf KeyCode = vbKeyF6 Then
        If showNum = False Then
            showNum = True
            lblTimer.Caption = "Displaying picture: " & i + 1 & " of " & files.ListCount & "."
            lblTimer.Move (Screen.Width - lblTimer.Width) / 2
            lblTimer.Visible = True
        Else
            showNum = False
            lblTimer.Visible = False
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If lblDuration.Visible = False Then
            lblDuration.Visible = True
        Else
            lblDuration.Visible = False
        End If
    ElseIf KeyCode = vbKeyF8 Then
        If showFileName = False Then
            showFileName = True
            lblFilename.Caption = files.List(i)
            lblFilename.Move (Screen.Width - lblFilename.Width) / 2
            lblFilename.Visible = True
        Else
            showFileName = False
            lblFilename.Visible = False
        End If
    ElseIf KeyCode = vbKeyEscape Then
        'unhide mouse cursor
        Do
        Loop Until ShowCursor(True) > 5
        
        End
    ElseIf KeyCode = vbKeyF12 Then
        'hide everything
        If frmMain.Width = 100 Then
            frmMain.Width = Screen.Width
            frmMain.Height = Screen.Height
            frmMain.Move 0, 0
            Timer1.Enabled = True
        End If
        frmMain.Width = 100
        frmMain.Height = 100
        frmMain.Move -3000, -3000
        Timer1.Enabled = False
        frmMain.SetFocus
    End If
End Sub

Private Sub Form_Load()
    
    'Make the cursor disappear
    Do
    Loop Until ShowCursor(False) < -5
    
    imgPictures.Visible = False
    frmMain.Move 0, 0
    frmMain.Width = Screen.Width
    frmMain.Height = Screen.Height
        
    Dim x As Picture
    
    imgPictures.Visible = False
    Set x = LoadPicture(files.List(i))
    imgPictures.Picture = x
    'shrink to fit algorithym

    If imgPictures.Width > Screen.Width Then
        imgPictures.Stretch = True
        imgPictures.Height = (Screen.Width / imgPictures.Width) * imgPictures.Height
        imgPictures.Width = Screen.Width
    End If
    If imgPictures.Height > Screen.Height Then
        imgPictures.Stretch = True
        imgPictures.Width = (Screen.Height / imgPictures.Height) * imgPictures.Width
        imgPictures.Height = Screen.Height
    End If
    
    lblFilename.Move 3000, Screen.Height - 400
    imgPictures.Move (Screen.Width - imgPictures.Width) / 2, (Screen.Height - imgPictures.Height) / 2
    imgPictures.Visible = True
    showNum = True
    
    i = 0
End Sub

Private Sub imgPictures_Click()
    'unhide mouse cursor
    Do
    Loop Until ShowCursor(True) > 5
    
    End
End Sub

Private Sub labelTimer_Timer()
    If showNum = False Then
        If counter = 0 Then
            lblTimer.Visible = True
            counter = counter + 1
        ElseIf counter = 1 Then
            lblTimer.Visible = False
            labelTimer.Enabled = False
            counter = 0
        End If
    Else
        labelTimer.Enabled = False
        counter = 0
    End If
End Sub

Private Sub skipTimer_Timer()
    If skipcount = 0 Then
        Timer1.Enabled = False
        skipcount = skipcount + 1
    ElseIf skipcount = 1 Then
        Timer1.Enabled = True
        skipcount = 0
        skipTimer.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    Dim x As Picture
    pass = pass + 1
    
    If pass = 1 Then
        Timer1.Interval = 5000
    End If
    
    If pass > 32000 Then pass = 2
    
    If i >= files.ListCount Then
        dirListing frmLoad.Dir1.List(frmLoad.Dir1.ListIndex)
        i = 0
    End If
    
    imgPictures.Stretch = False
    
    If showNum = True Then
        lblTimer.Caption = "Displaying picture: " & i + 1 & " of " & files.ListCount & "."
        lblTimer.Move (Screen.Width - lblTimer.Width) / 2
        lblTimer.Visible = True
    Else
        lblTimer.Visible = False
    End If
    
    If showFileName = True Then
        lblFilename.Caption = files.List(i)
        lblFilename.Move (Screen.Width - lblFilename.Width) / 2
        lblFilename.Visible = True
    Else
        lblFilename.Visible = False
    End If
    
    calcDuration
       
    imgPictures.Visible = False
    On Error GoTo BadPic
    Set x = LoadPicture(files.List(i))
    imgPictures.Picture = x
    'shrink to fit algorithym
    On Error GoTo tooBig
    If imgPictures.Width > Screen.Width Then
        imgPictures.Stretch = True
        imgPictures.Height = (Screen.Width / imgPictures.Width) * imgPictures.Height
        imgPictures.Width = Screen.Width
    End If
    If imgPictures.Height > Screen.Height Then
        imgPictures.Stretch = True
        imgPictures.Width = (Screen.Height / imgPictures.Height) * imgPictures.Width
        imgPictures.Height = Screen.Height
    End If
    
    imgPictures.Move (Screen.Width - imgPictures.Width) / 2, (Screen.Height - imgPictures.Height) / 2
    imgPictures.Visible = True
    
    i = i + 1
    Exit Sub

tooBig:
    imgPictures.Height = Screen.Height
    imgPictures.Width = Screen.Width
    imgPictures.Move 0, 0
    imgPictures.Visible = True
    i = i + 1
    If i = files.ListCount - 1 Then
        i = 0
    End If
    Exit Sub

BadPic:
    lblTimer.Caption = files.List(i) & " was unreadable!"
    i = i + 1
    lblTimer.Move (Screen.Width - lblTimer.Width) / 2
    lblTimer.Visible = True
    labelTimer.Enabled = True
    counter = 0
    Resume
End Sub

Public Sub calcDuration()
   If lblDuration.Visible = True Then
        Dim duration As Single
        Dim h As Variant
        Dim m As Variant
        Dim s As Variant
        
        'time slideshow will last in minutes
        'assume that it will never be able to do more than 3 slides per second.
        'probably closer to 1 per second for normal pictures
        If Timer1.Interval > 900 Then
            duration = ((files.ListCount - i) * (Timer1.Interval / 1000) / 60)
        Else
            duration = ((files.ListCount - i) * 0.95) / 60
        End If
        
        h = Format(duration / 60, "#00.##0")
        m = Right(h, 4)
        m = Format(m * 60, "#00.##0")
        s = Right(m, 4)
        s = Format(s * 60, "00")
        m = Left(m, 2)
        h = Left(h, 2)
        
        lblDuration.Caption = "Slideshow will last approximately another " & h & ":" & m & ":" & s & " before it repeats."
        lblDuration.Move (Screen.Width - lblDuration.Width) / 2
    End If
End Sub
