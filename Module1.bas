Attribute VB_Name = "Module1"
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow&)

Type dirfile
    owner As String
    subDir(2000) As String
    i As Integer
    numSubDir As Integer
End Type

Public Sub dirListing(rootFldr As String)
    Dim dirs(2000) As dirfile
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim total As Integer
    Dim restartFlag As Integer
    Dim myPath As Variant
    Dim myName As Variant
    Dim final As Integer
    Dim numtries As Integer
    
    If Right(rootFldr, 1) <> "\" Then
        rootFldr = rootFldr & "\"
    End If
    
    dirs(0).i = 0
    dirs(0).owner = rootFldr
    'Display the names in cdrom that represent directories.
    myPath = dirs(0).owner
    
    Open "C:\TEST.TXT" For Output As #3
    Print #3, dirs(0).owner
    
    i = 0
    
    restartFlag = 0
    'On Error GoTo PutInCD
    'File1.Path = cdrom
    'initPath = cdrom
    'On Error GoTo 0
Restart:
    'modify this if branch
    '    final = (Dir1.ListCount - 1)
    '    restartFlag = restartFlag + 1
    '    If restartFlag = 1 Then
    '        initPath = Dir1.Path & "\"
    '    End If
    'End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    myName = Dir(myPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> ""   ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        If myName <> "." And myName <> ".." Then
            ' Use bitwise comparison to make sure MyName is a directory.
            If (GetAttr(myPath & myName) And vbDirectory) = vbDirectory Then
                dirs(i).subDir(dirs(i).i) = myName
                dirs(i).i = dirs(i).i + 1
                Print #3, myPath & myName & "\"
                  ' Display entry only if it
            End If  ' it represents a directory.
        End If
        myName = Dir    ' Get next entry.
    Loop
    dirs(i).numSubDir = dirs(i).i
    
    For i = 1 To dirs(0).numSubDir
        myPath = dirs(0).owner & dirs(0).subDir(i - 1) & "\"
        dirs(i).owner = myPath
        myName = Dir(myPath, vbDirectory)   ' Retrieve the first entry.
        Do While myName <> ""   ' Start the loop.
            ' Ignore the current directory and the encompassing directory.
            If myName <> "." And myName <> ".." Then
                ' Use bitwise comparison to make sure MyName is a directory.
                If (GetAttr(myPath & myName) And vbDirectory) = vbDirectory Then
                    dirs(i).i = dirs(i).i + 1
                    dirs(i).subDir(dirs(i).i) = myName
                    ' Display entry only if it
                End If  ' it represents a directory.
            End If
            myName = Dir      ' Get next entry.
        Loop
        dirs(i).numSubDir = dirs(i).i
    Next i
    total = i - 1

    i = total + 1

        For ii = 1 To total
            For j = 0 To dirs(ii).numSubDir
                If dirs(ii).subDir(j) <> "" Then
                    myPath = dirs(ii).owner & dirs(ii).subDir(j) & "\"
                Else
                    myPath = dirs(ii).owner
                End If
                dirs(i).owner = myPath
                myName = Dir(myPath, vbDirectory)   ' Retrieve the first entry.
                Do While myName <> ""   ' Start the loop.
                    ' Ignore the current directory and the encompassing directory.
                    If myName <> "." And myName <> ".." Then
                        ' Use bitwise comparison to make sure MyName is a directory.
                        If (GetAttr(myPath & myName) And vbDirectory) = vbDirectory Then
                            dirs(i).i = dirs(i).i + 1
                            dirs(i).subDir(dirs(i).i) = myName
                            Print #3, myPath & myName & "\"
                            ' Display entry only if it
                        End If  ' it represents a directory.
                    End If
                    myName = Dir      ' Get next entry.
                Loop
                dirs(i).numSubDir = dirs(i).i
                i = i + 1
            Next j
        Next ii
        total = i - 1
    Close
    
    Dim directory As String
    Open "c:\test.txt" For Input As #1
    Dim ext As String
    Do Until EOF(1)
        Line Input #1, directory
        frmLoad.File1.Path = directory
        For j = 0 To (frmLoad.File1.ListCount - 1)
            ext = LCase(Right(frmLoad.File1.List(j), 4))
            If ext = ".jpg" Or ext = "jpeg" Or ext = ".gif" Or ext = ".bmp" Or ext = ".tif" Or ext = ".jpe" Or ext = ".pic" Or ext = ".tga" Or ext = "tiff" Then
                If Right(frmLoad.File1.Path, 1) <> "\" Then
                    frmMain.files.AddItem frmLoad.File1.Path & "\" & frmLoad.File1.List(j)
                Else
                    frmMain.files.AddItem frmLoad.File1.Path & frmLoad.File1.List(j)
                End If
            End If
        Next j
    Loop
    Close
End Sub


