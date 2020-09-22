Attribute VB_Name = "modAKRip32"
Option Explicit

' How many Frames per read
Public Const SECTORSPERREAD = 26

' How many times we try to read CD before we report error...
Public Const RetriesCount = 3

Public Const TRACK_AUDIO = &H0
Public Const TRACK_DATA = &H1
Public Const MAXIDLEN = 64
Public Const MAXCDLIST = 7

'/*
' * defines for GETCDHAND  readType
' *
' */
Public Enum ReadTypeEnum
  CDR_ANY = &H0                 ' unknown
  CDR_ATAPI1 = &H1              ' ATAPI per spec
  CDR_ATAPI2 = &H2              ' alternate ATAPI
  CDR_READ6 = &H3               ' using SCSI READ(6)
  CDR_READ10 = &H4              ' using SCSI READ(10)
  CDR_READ_D8 = &H5             ' using command = &HD8 (Plextor?)
  CDR_READ_D4 = &H6             ' using command = &HD4 (NEC?)
  CDR_READ_D4_1 = &H7           ' = &HD4 with a mode select
  CDR_READ10_2 = &H8            ' different mode select w/ READ(10)
End Enum

'/*
'* defines for the read mode (CDP_READMODE)
'*/
Public Enum ReadModeEnum
  CDRM_NOJITTER = &H0          '// never jitter correct
  CDRM_JITTER = &H1            '// always jitter correct
  CDRM_JITTERONERR = &H2       '// jitter correct only after a read error
End Enum

'/*
' * constants used for queryCDParms()
' */
Public Enum QueryCDParamsEnum
  CDP_READCDR = &H1             ' can read CD-R
  CDP_READCDE = &H2             ' can read CD-E
  CDP_METHOD2 = &H3             ' can read CD-R wriiten via method 2
  CDP_WRITECDR = &H4            ' can write CD-R
  CDP_WRITECDE = &H5            ' can write CD-E
  CDP_AUDIOPLAY = &H6           ' can play audio
  CDP_COMPOSITE = &H7           ' composite audio/video stream
  CDP_DIGITAL1 = &H8            ' digital output (IEC958) on port 1
  CDP_DIGITAL2 = &H9            ' digital output (IEC958) on port 2
  CDP_M2FORM1 = &HA             ' reads Mode 2 Form 1 (XA) format
  CDP_M2FORM2 = &HB             ' reads Mode 2 Form 2 format
  CDP_MULTISES = &HC            ' reads multi-session or Photo-CD
  CDP_CDDA = &HD                ' supports cd-da
  CDP_STREAMACC = &HE           ' supports "stream is accurate"
  CDP_RW = &HF                  ' can return R-W info
  CDP_RWCORR = &H10             ' returns R-W de-interleaved and err.
                                ' corrected
  CDP_C2SUPP = &H11             ' C2 error pointers
  CDP_ISRC = &H12               ' can return the ISRC info
  CDP_UPC = &H13                ' can return the Media Catalog Number
  CDP_CANLOCK = &H14            ' prevent/allow cmd. can lock the media
  CDP_LOCKED = &H15             ' current lock state (TRUE = LOCKED)
  CDP_PREVJUMP = &H16           ' prevent/allow jumper state
  CDP_CANEJECT = &H17           ' drive can eject disk
  CDP_MECHTYPE = &H18           ' type of disk loading supported
  CDP_SEPVOL = &H19             ' independent audio level for channels
  CDP_SEPMUTE = &H1A            ' independent mute for channels
  CDP_SDP = &H1B                ' supports disk present (SDP)
  CDP_SSS = &H1C                ' Software Slot Selection
  CDP_MAXSPEED = &H1D           ' maximum supported speed of drive
  CDP_NUMVOL = &H1E             ' number of volume levels
  CDP_BUFSIZE = &H1F            ' size of output buffer
  CDP_CURRSPEED = &H20          ' current speed of drive
  CDP_SPM = &H21                ' "S" units per "M" (MSF format)
  CDP_FPS = &H22                ' "F" units per "S" (MSF format)
  CDP_INACTMULT = &H23          ' inactivity multiplier ( x 125 ms)
  CDP_MSF = &H24                ' use MSF format for READ TOC cmd
  CDP_OVERLAP = &H25            ' number of overlap frames for jitter
  CDP_JITTER = &H26             ' number of frames to check for jitter
  CDP_READMODE = &H27           ' mode to attempt jitter corr.
End Enum

'/*
' * Error codes set by functions in ASPILIB.C
' */
Public Enum AKErrorEnum
  ALERR_NOERROR = 0
  ALERR_NOWNASPI = -1
  ALERR_NOGETASPI32SUPP = -2
  ALERR_NOSENDASPICMD = -3
  ALERR_ASPI = -4
  ALERR_NOCDSELECTED = -5
  ALERR_BUFTOOSMALL = -6
  ALERR_INVHANDLE = -7
  ALERR_NOMOREHAND = -8
  ALERR_BUFPTR = -9
  ALERR_NOTACD = -10
  ALERR_LOCK = -11
  ALERR_DUPHAND = -12
  ALERR_INVPTR = -13
  ALERR_INVPARM = -14
  ALERR_JITTER = -15
End Enum


' CD Information
Public Type CDINFO
  vendor(8) As Byte
  prodId(16)  As Byte
  rev(4) As Byte
  vendSpec(20) As Byte
End Type

Public Type CDREC
  ha As Byte
  tgt As Byte
  lun As Byte
  pad As Byte
  id(MAXIDLEN) As Byte
  info As CDINFO
End Type

Public Type CDLIST
  max As Byte
  num As Byte
  cd(MAXCDLIST) As CDREC
End Type

' CD Drive Info
Public Type GETCDHAND
  Size As Byte
  ver As Byte
  ha As Byte
  tgt As Byte
  lun As Byte
  ReadType As Byte
  jitterCorr As Long
  numJutter As Byte
  numOverlap As Byte
End Type

' Table Of Contests
Public Type TOCTRACK
  trackNumber As Byte
  rsvd2 As Byte
  addr(1 To 4) As Byte
  rsvd As Byte
  ADR As Byte
End Type

Public Type TOC
  tocLen As Long
  firstTrack As Byte
  lastTrack As Byte
  tracks(1 To 100) As TOCTRACK
End Type

' Buffer
Public Type PTRACKBUF
  startFrame As Long
  NumFrames As Long
  maxLen As Long
  len As Long
  Status As Long
  startOffset As Long
  'buf(BUFFERSIZE) As Byte
End Type

Public Type DWORD
  HIWORD As Integer
  LOWORD As Integer
End Type

Public Type TestType
  one As Long
  two As Long
End Type

' Declarations
Public Declare Function CloseCDHandle Lib "akrip32.dll" (ByVal hCD As Long) As Boolean
Public Declare Function GetCDHandle Lib "akrip32.dll" (lpcd As Any) As Byte
Public Declare Function GetCDId Lib "akrip32.dll" (ByVal hCD As Long, ByVal Buffer As String, ByVal BufferLen As Integer) As Long
Public Declare Function GetCDList Lib "akrip32.dll" (lpcd As Any) As Integer
Public Declare Function GetDriveInfo Lib "akrip32.dll" (ByVal ha As Byte, ByVal tgt As Byte, ByVal lun As Byte, VCDREC As CDREC) As Long
Public Declare Function GetNumAdapters Lib "akrip32.dll" () As Long

Public Declare Function ReadCDAudioLBA Lib "akrip32.dll" (ByVal hCD As Long, ByVal lpTrackBuf As Long) As Long
Public Declare Function ReadCDAudioLBAEx Lib "akrip32.dll" (ByVal hCD As Long, ByVal lpTrackBuf As Long, ByVal lpOverlap As Long) As Long
Public Declare Function ReadTOC Lib "akrip32.dll" (ByVal hCD As Integer, lpTOC As Any) As Long
Public Declare Function GetAKRipDllVersion Lib "akrip32.dll" () As DWORD
Public Declare Function GetAspiLibError Lib "akrip32.dll" () As AKErrorEnum

Public Declare Function ModifyCDParms Lib "akrip32.dll" (ByVal hCD As Long, ByVal which As Integer, ByVal val As Long) As Boolean
Public Declare Function QueryCDParms Lib "akrip32.dll" (ByVal hCD As Long, ByVal which As Integer, ByRef pNum As Long) As Boolean

Public CDHandle As Long
Public DiscToc As TOC

Public Function GetDriveInformation(DriveNo As Long) As CDREC
  Dim CD_List As CDLIST
  GetCDList CD_List
  GetDriveInformation = CD_List.cd(DriveNo)
End Function

Public Function InitCDDrive(DriveNo As Long) As Boolean
  Dim CD_Info As CDREC
  Dim CDH As GETCDHAND

  ' Get Info
  CD_Info = GetDriveInformation(DriveNo)
  
  CDH.Size = Len(CDH)
  CDH.ver = 1
  CDH.ha = CD_Info.ha
  CDH.tgt = CD_Info.tgt
  CDH.lun = CD_Info.lun
  CDH.ReadType = CDR_ANY '      // set for autodetect
  
  ' Get Handle
  CDHandle = GetCDHandle(ByVal VarPtr(CDH))
  If CDHandle = 0 Then
    InitCDDrive = False
    Exit Function
  End If
  
  Call ModifyCDParms(CDHandle, CDP_MSF, False)
  
  ' Read Table of Contents
  If ReadTOC(CDHandle, ByVal VarPtr(DiscToc)) <> 1 Then
    InitCDDrive = False
    Exit Function
  End If
  
  InitCDDrive = True
End Function


' Lil helper... Hope I converted it right...
Public Function MSB2LONG(b() As Byte) As Long
  MSB2LONG = CLng(b(1)) * 256 * 256 * 256
  MSB2LONG = MSB2LONG + CLng(b(2)) * 256 * 256
  MSB2LONG = MSB2LONG + CLng(b(3)) * 256
  MSB2LONG = MSB2LONG + CLng(b(4))
End Function

' Close CD Drive
Public Function DeInitCDDrive() As Boolean
  DeInitCDDrive = CloseCDHandle(CDHandle)
End Function

' This one tells what went wrong....
Public Function GetAKRipError() As String
  Dim ErrNo As AKErrorEnum
  Err.Clear
  ErrNo = GetAspiLibError
  If Err Then Debug.Assert 0
  Select Case ErrNo
  Case ALERR_NOERROR: GetAKRipError = "No error..."
  Case ALERR_NOWNASPI: GetAKRipError = "Unable to load WNASPI32.DLL (95/98/NT/2000) or use SCSI passthrough (NT/2000)"
  Case ALERR_NOGETASPI32SUPP: GetAKRipError = "Could not load ASPI function GetASPI32SupportInfo. Only occurs when an ASPI manager (WNASPI32.DLL) is found, but cannot be correctly loaded by the system. Most often indicates that ASPI is improperly installed."
  Case ALERR_NOSENDASPICMD: GetAKRipError = "Could not load ASPI function SendASPI32Command. Only occurs when an ASPI manager (WNASPI32.DLL) is found, but cannot be correctly loaded by the system. Most often indicates that ASPI is improperly installed."
  Case ALERR_ASPI: GetAKRipError = "An error was returned by the ASPI manager. Use GetAspiLibAspiError to retrieve the actual ASPI error."
  Case ALERR_NOCDSELECTED: GetAKRipError = "Unused in the current implementation"
  Case ALERR_BUFTOOSMALL: GetAKRipError = "The buffer passed to ReadCDAudioLBA or ReadCDAudioLBAEx is too small for the requested number of frames."
  Case ALERR_INVHANDLE: GetAKRipError = "The handle to the CD-ROM unit is invalid."
  Case ALERR_NOMOREHAND: GetAKRipError = "All available slots for CD-ROM handles have been allocated."
  Case ALERR_BUFPTR: GetAKRipError = "Results from passing a bad buf parameter to GetCDId (ie. a NULL pointer)"
  Case ALERR_NOTACD: GetAKRipError = "The ha:tgt:lun values passed to GetCDHandle do not refer to a CD-ROM device."
  Case ALERR_LOCK: GetAKRipError = "Unable to obtain an exclusive lock on a CD handle."
  Case ALERR_DUPHAND: GetAKRipError = "Occurs when attempting to call GetCDHandle for a ha:tgt:lun value that has already been allocated."
  Case ALERR_INVPTR: GetAKRipError = "Invalid value for the LPGETCDHAND parameter to GetCDHandle"
  Case ALERR_INVPARM: GetAKRipError = "Invalid version or size specified in LPGETCDHAND parameter to GetCDHandle"
  Case ALERR_JITTER: GetAKRipError = "An automatic jitter adjust failed during a call to ReadCDAudioLBAEx"
  End Select
End Function

