Attribute VB_Name = "modGeneral"
Option Explicit

' /*
' ** Allocates the specified number of bytes from the heap.
' */
Public Declare Function GlobalAlloc _
    Lib "kernel32" ( _
        ByVal wFlags As Long, _
        ByVal dwBytes As Long) As Long

' /*
' ** Locks a global memory object and returns a pointer to
' ** the first byte of the bject's memory block.
' ** The memory block associated with a locked object cannot
' ** be moved or discarded.
'*/
Public Declare Function GlobalLock _
    Lib "kernel32" ( _
        ByVal hmem As Long) As Long

' /*
' ** Frees the specificed global memory object and
' ** invalidates its handle
' */
Public Declare Function GlobalFree _
    Lib "kernel32" ( _
        ByVal hmem As Long) As Long

Public Declare Sub CopyPtrFromStruct _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        ByVal ptr As Long, _
        struct As Any, _
        ByVal cb As Long)
        
Public Declare Sub memcpy _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        ptr1 As Any, _
        Ptr2 As Any, _
        ByVal cb As Long)
        
Public Declare Sub CopyMemory _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        ByVal ptr1 As Long, _
        ByVal Ptr2 As Long, _
        ByVal cb As Long)

Public Declare Sub CopyStructFromPtr _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        struct As Any, _
        ByVal ptr As Long, _
        ByVal cb As Long)

Public Type SHITEMID    'Browse Dialog
   cb             As Long
   abID           As Byte
End Type

Public Type ITEMIDLIST  'Browse Dialog
   mkid           As SHITEMID
End Type

Public Type BROWSEINFO  'Browse Dialog
   hOwner         As Long
   pidlRoot       As Long
   pszDisplayName As String
   lpszTitle      As String
   ulFlags        As Long
   lpfn           As Long
   lParam         As Long
   iImage         As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1 'Browse Dialog
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
 
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long



Public Cancelled As Boolean

' This removes Nulls from arrays and returns string
Public Function StripNullsArray(STR) As String
  Dim i As Long
  For i = LBound(STR) To UBound(STR)
    If STR(i) <> 0 Then StripNullsArray = StripNullsArray & Chr(STR(i))
  Next i
End Function

' This removes Nulls from string
Public Function StripNulls(STR As String) As String
  StripNulls = Left(STR, InStr(STR, Chr(0)) - 1)
End Function

Public Function ChangeExt(Filename As String, NewExt As String)
  ChangeExt = Left(Filename, InStrRev(Filename, ".")) & NewExt
End Function

' This adds slash to path if it is not there...
Public Function AddSlash(FullPath As String) As String
  AddSlash = IIf(Right(FullPath, 1) = "\", FullPath, FullPath & "\")
End Function

'Opens Browse dialog
Public Function BrowseForFolder(Optional Title As String) As String
   Dim bi As BROWSEINFO
   Dim pidl As Long
   Dim nRet As Long
   Dim szPath As String
   
   szPath = Space$(512)
   
   bi.hOwner = 0&
   bi.pidlRoot = 0&
   
   bi.lpszTitle = IIf(Title = "", "Directory", Title)
   bi.ulFlags = BIF_RETURNONLYFSDIRS
   
   'Display the dialog and get the selected path
   pidl& = SHBrowseForFolder(bi)
   SHGetPathFromIDList ByVal pidl&, ByVal szPath
   
   'Return value
   BrowseForFolder = Trim$(szPath)
End Function

' Extract track to MP3 file
Public Sub RipMP3(outMP3Name As String, StartAddr As Long, EndAddr As Long, beConfig As PBE_CONFIG)
  On Error GoTo ErrHandler
  ChDrive App.Path
  ChDir App.Path
  
  ' Check that file exists...
  Cancelled = False
  
  'Open progress window
  frmProgress.Show , frmMain
  
  Dim error As Long
  Dim dwSamples As Long, dwMP3Buffer As Long, hbeStream As Long
  
  ' Init MP3 Stream
  error = beInitStream(VarPtr(beConfig), VarPtr(dwSamples), VarPtr(dwMP3Buffer), VarPtr(hbeStream))
    
  '// Check result
  If error <> BE_ERR_SUCCESSFUL Then
    Err.Raise error, "Lame", GetErrorString(error)
  End If
  
  
  ' Open MP3 file...
  Dim WriteFile As clsFileIo
  Set WriteFile = New clsFileIo
  WriteFile.OpenFile outMP3Name
  
  
  Dim NumFrames   As Long
  Dim Dummy       As PTRACKBUF
  Dim BufferPtr1  As Long
  Dim BufferPtr2  As Long
  Dim LLen        As Long
  Dim Retries     As Long
  Dim Status      As Long
  Dim NumWritten  As Long
  Dim toRead      As Long, toWrite As Long
  Dim Done        As Long
  Dim length      As Long
  
  NumFrames = SECTORSPERREAD
    
  'Initialize Audio Extraction buffer
  BufferPtr1 = GlobalAlloc(&H40, NumFrames * 2352 + Len(Dummy))
  BufferPtr2 = GlobalLock(BufferPtr1)
  
  ' Dummy is used to inform AKRip what to extract from CD
  Dummy.startFrame = 0
  Dummy.NumFrames = 0
  Dummy.maxLen = NumFrames * 2352
  Dummy.len = 0
  Dummy.Status = 0
  Dummy.startOffset = 0
  
  'We copy Dummy into buffer...
  CopyMemory ByVal BufferPtr2, ByVal VarPtr(Dummy), Len(Dummy)
  
  Dim temp As Long
  temp = EndAddr - StartAddr
  LLen = EndAddr - StartAddr

  ' Allocate memory for MP3 buffer...
  Dim MP3Ptr1 As Long
  Dim MP3Ptr2 As Long
  
  MP3Ptr1 = GlobalAlloc(&H40, dwMP3Buffer)
  MP3Ptr2 = GlobalLock(MP3Ptr1)
  
  Dim NoOfBytes2Encode As Long
  
  Dim i As Long
  
  ' Lets start MP3 Extraction...
  
  Do While LLen
    ' Calculate how much we wanna rip from CD
    If LLen < NumFrames Then NumFrames = LLen
      
    Retries = RetriesCount
    Status = 0
    
    ' Try to read cd...
    Do While Retries > 0 And Status <> 1
      Dummy.NumFrames = NumFrames
      Dummy.startOffset = 0
      Dummy.len = 0
      Dummy.startFrame = StartAddr
      
      'Write info to buffer so that akrip knows what to read... :)
      CopyMemory ByVal BufferPtr2, ByVal VarPtr(Dummy), Len(Dummy)
      
      Status = ReadCDAudioLBA(CDHandle, BufferPtr2)
    Loop
    
    If Status <> 1 Then
      ' This is bad.... and there is nothing we can do...
      MsgBox GetAKRipError
      Exit Do
    End If
    
    ' Encode every frame we just extracted from CD
    For i = 0 To NumFrames - 1
      NoOfBytes2Encode = 2352 / 2 'One frame is 2352 bytes
                                  'Note: it is splitted because
                                  'LAME uses "Short" samples for encoding
      ' Encode buffer
      ' Note: Don't encode info... Memory position is pointer + lenght of Dummy
      error = beEncodeChunk(hbeStream, NoOfBytes2Encode, BufferPtr2 + Len(Dummy) + i * 2352, MP3Ptr2, VarPtr(toWrite))
      
      ' Write buffer to HardDrive buffer
      If toWrite > 0 Then WriteFile.WriteBytes MP3Ptr2, toWrite
    Next i
    
    ' Write HardDrive buffers to disk
    Call WriteFile.FlushBuffers
    
    
    ' We have written this much bytes and blahblahblahblah.... :)
    NumWritten = NumWritten + NumFrames * 2352
    StartAddr = StartAddr + NumFrames
    LLen = LLen - NumFrames
    
    If Cancelled Then Exit Do
    'Inform user where we go...
    frmProgress.ChangeProgress "Extracting track " & outMP3Name, CSng((temp - LLen)), CSng(temp)
    DoEvents
  Loop
    
    
  
  ' Deinitialize stream and write last bytes to MP3
  error = beDeinitStream(hbeStream, MP3Ptr2, VarPtr(toWrite))

  '//if close out was unsuccessful manually close stream
  If toWrite > 0 Then
    WriteFile.WriteBytes MP3Ptr2, toWrite
    WriteFile.FlushBuffers
  End If
  
  
  ' Clear buffers....
  GlobalFree MP3Ptr2
  GlobalFree BufferPtr2
  
  ' Close files
  Call WriteFile.CloseFile
  Set WriteFile = Nothing
  
  ' Close stream
  Call beCloseStream(hbeStream)
  
  
  ' WriteVBRHeader (if we use variable bitrate...)
  'Call beWriteVBRHeader(ChangeExt(Text1, "mp3"))
  Unload frmProgress
  
  Exit Sub
  
ErrHandler:
  ' Damn.. Something went wrong and this one should tell what...
  
  MsgBox Err.Description, vbCritical, "Critical error..."
  If BufferPtr2 Then GlobalFree BufferPtr2
  If MP3Ptr2 Then GlobalFree MP3Ptr2
  WriteFile.FlushBuffers
  WriteFile.CloseFile
  Unload frmProgress
  Err.Clear
End Sub

' Extract track to WAV file
Public Function RipWAV(Filename As String, addrStart As Long, addrEnd As Long)
  Dim StartAddr   As Long
  Dim EndAddr     As Long
  Dim NumFrames   As Long
  Dim Dummy       As PTRACKBUF
  Dim BufferPtr1  As Long
  Dim BufferPtr2  As Long
  Dim LLen        As Long
  Dim Retries     As Long
  Dim Status      As Long
  Dim NumWritten  As Long
  Dim OpenFile As clsFileIo
  
  Cancelled = False
  frmProgress.Show , frmMain

  NumFrames = SECTORSPERREAD
   
  ' Convert Addresses
  StartAddr = addrStart
  EndAddr = addrEnd
  
  'Initialize buffer
  BufferPtr1 = GlobalAlloc(&H40, NumFrames * 2352 + Len(Dummy))
  BufferPtr2 = GlobalLock(BufferPtr1)
  
  Dummy.startFrame = 0
  Dummy.NumFrames = 0
  Dummy.maxLen = NumFrames * 2352
  Dummy.len = 0
  Dummy.Status = 0
  Dummy.startOffset = 0
  
  CopyMemory ByVal BufferPtr2, ByVal VarPtr(Dummy), Len(Dummy)
  
  Dim temp As Long
  temp = EndAddr - StartAddr
  LLen = EndAddr - StartAddr
  
  ' Open files
  Set OpenFile = New clsFileIo
  
  OpenFile.OpenFile Filename
  OpenFile.writeWavHeader LLen * 2352
  
  Dim TempCount As Byte
  
  ' Lets start rippin...
  Do While LLen
    ' Calculate how much we wanna rip from CD
    If LLen < NumFrames Then NumFrames = LLen
      
    Retries = RetriesCount
    Status = 0
    
    ' Try to read cd...
    Do While Retries > 0 And Status <> 1
      Dummy.NumFrames = NumFrames
      Dummy.startOffset = 0
      Dummy.len = 0
      Dummy.startFrame = StartAddr
      
      'Write info to buffer so that akrip knows what to read... :)
      CopyMemory ByVal BufferPtr2, ByVal VarPtr(Dummy), Len(Dummy)
      
      Status = ReadCDAudioLBA(CDHandle, BufferPtr2)
    Loop
    
    If Status = 1 Then
      ' Write buffer to disk
      ' Note: Don't write info to disk... Memory position is pointer + lenght of Dummy
      OpenFile.WriteBytes BufferPtr2 + Len(Dummy), NumFrames * 2352
    Else
      ' Doh.... This is bad.... and there is nothing we can do...
      MsgBox GetAKRipError
      Exit Do
    End If
    
    ' We have written this much bytes and blahblahblahblah.... :)
    NumWritten = NumWritten + NumFrames * 2352
    StartAddr = StartAddr + NumFrames
    LLen = LLen - NumFrames
    
    'Inform user where we go...
    frmProgress.ChangeProgress "Extracting track " & Filename, CSng(temp - LLen), CSng(temp)
    DoEvents
  Loop
  
  ' Delete buffer and close files
  GlobalFree BufferPtr2
  OpenFile.CloseFile
  Set OpenFile = Nothing
  Unload frmProgress
End Function

