VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dCheckSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PowerTable(7) As Byte

Enum TCrcType
    c16 = 0
    c32 = 1
End Enum

Private Sub IniTable()
    PowerTable(0) = 1
    PowerTable(1) = 2
    PowerTable(2) = 4
    PowerTable(3) = 8
    PowerTable(4) = 16
    PowerTable(5) = 32
    PowerTable(6) = 64
    PowerTable(7) = 128
End Sub

Public Function GetCheckSum(lpStr As String, Optional CrcType As TCrcType = c16) As Long
Dim Count As Long
Dim Cnt As Long
Dim a As Boolean
Dim b As Boolean
Dim BitFlag As Boolean
Dim sBytes() As Byte
Dim mCrc As Long

    'Convert string to byte array.
    sBytes = StrConv(lpStr, vbFromUnicode)
    'Loop tho all the bytes
    Do Until (Count > UBound(sBytes))
        Cnt = 7
        'Loop tho the power table
        Do Until (Cnt < 0)
            'Build the checksum
            a = ((mCrc And &H8000&) = &H8000&)
            b = ((sBytes(Count) And PowerTable(Cnt)) = PowerTable(Cnt))
            mCrc = (mCrc And &H7FFF) * 2
            BitFlag = (a Xor b)
            'Test BitFlag is true
            If (BitFlag) Then
                Select Case CrcType
                    Case c16
                        'Make Crc 16bit
                        mCrc = (mCrc Xor &H1021&)
                    Case c32
                        'Make crc 32bit
                        mCrc = (mCrc Xor &H8005&)
                End Select
            End If
            Cnt = Cnt - 1
        Loop
        'INC
        Count = Count + 1
    Loop

    GetCheckSum = mCrc
    BitFlag = False
    Count = 0
    Erase sBytes
    
End Function

Private Sub Class_Initialize()
    Call IniTable
End Sub

Private Sub Class_Terminate()
    Erase PowerTable
End Sub
