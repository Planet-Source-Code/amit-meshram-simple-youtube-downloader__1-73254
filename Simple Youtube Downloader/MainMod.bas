Attribute VB_Name = "MainMod"
Option Explicit
Public VideoName As String
Public GFileName As String
    
Function GetVideoInfo(URL As String, InetPre As Inet)
    
    Dim var_data As Variant
    Dim RespText As String
    Dim VideoID As String
    Dim URLVideoID  As String
    
    frmMain.STBar.Panels(1).Text = "Getting Video Information.."
    
    RespText = InetPre.OpenURL(URL)
    
    Do While InetPre.StillExecuting
        DoEvents
    Loop
    
    var_data = InetPre.GetChunk(1024, icString)
    RespText = RespText & var_data
    
    Do
        DoEvents
        var_data = InetPre.GetChunk(1024, icString)
        If Len(var_data) = 0 Then Exit Do
        RespText = RespText & var_data
    Loop
      
    VideoName = GetVideoTitle(RespText)
    VideoID = GetVideoID(RespText)
    URLVideoID = GetURLVideoID(URL)
    
    frmMain.lblVidName = VideoName
        
    GetVideoInfo = "http://youtube.com/get_video?video_id=" & URLVideoID & "&t=" & VideoID
End Function

Function GetURLVideoID(URL As String) As String
    Dim str1, str2
    Dim pos1, pos2 As Integer
    
    If InStr(1, URL, "v=") Then
        pos1 = InStr(1, URL, "v=")
        pos2 = InStr(pos1, URL)
        str1 = Mid(URL, pos1, 14)
        str2 = Mid(str1, 3, Len(str1) - 2)
    End If
    GetURLVideoID = str2
End Function

Function GetVideoID(RespText As String) As String
    Dim pos1, pos2 As Integer
    Dim tmp1, tmp2, tmp3 As String
    
    If InStr(1, RespText, Chr(&H22) & "t" & Chr(&H22) & ":") Then
        pos1 = InStr(1, RespText, Chr(&H22) & "t" & Chr(&H22) & ":")
        pos2 = InStr(pos1, RespText)
        tmp1 = Mid(RespText, pos1, 50)
        tmp2 = Mid(tmp1, 7, Len(tmp1))
    End If
    GetVideoID = tmp2
End Function

Function GetVideoTitle(RespText As String) As String
    Dim pos1, pos2 As Integer
    Dim tmp1, tmp2, tmp3 As String
    
    If InStr(1, RespText, "content") Then
        pos1 = InStr(1, RespText, "content=")
        pos2 = InStr(pos1, RespText, ">")
        tmp1 = Mid(RespText, pos1, pos2 - pos1 - 1)
        tmp2 = Replace(tmp1, "content=", "")
        tmp3 = Replace(tmp2, Chr(&H22), "")
    End If
    GetVideoTitle = Trim(tmp3)
End Function

Sub DownloadVideo(Link As String, FileName As String)
    Dim FileSize As Long
    Dim SrcSize As Double
    Dim FileData() As Byte
    Dim FileRemaining As Long
    Dim FileSizeCurrent As Long
    Dim PBValue As Integer
    
    Dim FileNumber As Long
        
    frmMain.STBar.Panels(1).Text = "Downloading Video..."
    frmMain.Inet2.Execute Trim(Link), "GET"
    
    Do While frmMain.Inet2.StillExecuting
        DoEvents
    Loop
    
    FileName = Replace(FileName, "/", "")
    FileName = Replace(FileName, "\", "")
    FileName = Replace(FileName, "*", "")
    FileName = Replace(FileName, ":", "")
    FileName = Replace(FileName, "?", "")
    FileName = Replace(FileName, "<", "")
    FileName = Replace(FileName, ">", "")
    FileName = Replace(FileName, "|", "")
    FileName = Replace(FileName, Chr(&H22), "")
    
    
    GFileName = FileName
    
    FileSize = frmMain.Inet2.GetHeader("Content-Length")
    SrcSize = FileSize / 1000
    
    frmMain.lblVidSize.Caption = SrcSize & " kb"
    
    FileRemaining = FileSize
    FileSizeCurrent = 0
    
    FileNumber = FreeFile
    
    Open App.Path & "/" & FileName For Binary Access Write As #FileNumber
        
        Do Until FileRemaining = 0
            If frmMain.Tag = "Cancel" Then
                frmMain.Inet2.Cancel
                frmMain.STBar.Panels(1).Text = "Stoped by user"
                Exit Sub
            End If
            
            If FileRemaining > 1024 Then
                FileData = frmMain.Inet2.GetChunk(1024, icByteArray)
                FileRemaining = FileRemaining - 1024
            Else
                FileData = frmMain.Inet2.GetChunk(FileRemaining, icByteArray)
                FileRemaining = 0
            End If
            
            FileSizeCurrent = FileSize - FileRemaining
            PBValue = CInt((100 / FileSize) * FileSizeCurrent)
            
            frmMain.lblSaved.Caption = FileSizeCurrent & " bits"
            frmMain.lblRemaining.Caption = FileSize - FileSizeCurrent & " bits"
            frmMain.lblPercent.Caption = PBValue & " %"
            frmMain.STBar.Panels(2).Text = PBValue & " %" & "Downloaded"
            
            Put #FileNumber, , FileData
        Loop
    Close #FileNumber
    MsgBox "Video Downloaded."
    Call frmMain.ResetControls
    Exit Sub
End Sub

Function GetStatus(st As Integer, Inet2 As Inet)
    Select Case st
        Case icError
            GetStatus = Left$(Inet2.ResponseInfo, Len(Inet2.ResponseInfo) - 2)
        Case icResolvingHost, icRequesting, icRequestSent
            GetStatus = "Searching... "
        Case icHostResolved
            GetStatus = "Found" & GFileName
        Case icReceivingResponse, icResponseReceived
            GetStatus = "Receiving data "
        Case icResponseCompleted
            GetStatus = "Connected"
        Case icConnecting, icConnected
            GetStatus = "Connecting..."
        Case icDisconnecting
            GetStatus = "Disconnecting..."
        Case icDisconnected
            GetStatus = "Disconnected"
        Case Else
    End Select
End Function


