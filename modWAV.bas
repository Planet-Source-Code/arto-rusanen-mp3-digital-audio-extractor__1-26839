Attribute VB_Name = "modWAV"
Option Explicit
'/*
' * WAV file header format
' */
Public Type WAVHDR
  riff(1 To 4) As Byte        '/* must be "RIFF"                */
  len As Long                 '/* #bytes + 44 - 8               */
  cWavFmt(1 To 8) As Byte     '/* must be "WAVEfmt"             */
  dwHdrLen As Long
  wFormat As Integer
  wNumChannels As Integer
  dwSampleRate As Long
  dwBytesPerSec As Long
  wBlockAlign As Integer
  wBitsPerSample As Integer
  cData(1 To 4) As Byte       '/* must be "data"               */
  dwDataLen As Long           '/* #bytes                       */
End Type

