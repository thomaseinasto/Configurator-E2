Attribute VB_Name = "read"
Option Explicit
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long) ' Sleep "ms"
Public Function modbus_read(adr As Long) As String
On Error GoTo cannot_open_port

Dim packet_bytes(7) As Byte ' 8 baiti
Dim packet_str  As String
Dim abi As Long
Dim regH  As Byte
Dim regL  As Byte
Dim crcH  As Byte
Dim crcL  As Byte
Dim data_out As Long

Dim delay_count As Long

regH = adr \ 256
abi = regH * 256#
regL = adr - abi

packet_bytes(0) = Form1.slave_id_global   ' global variable
packet_bytes(1) = 3          ' read register, function=3
packet_bytes(2) = regH
packet_bytes(3) = regL
packet_bytes(4) = 0          ' toimub ainult ühe reg. lugemine
packet_bytes(5) = 1          ' toimub ainult ühe reg. lugemine

Call CRC_calc(packet_bytes(), 6, crcH, crcL) ' "6" baiti on vaja läbi töötada

packet_bytes(6) = crcL
packet_bytes(7) = crcH

Call Bytes_to_string(packet_bytes(), packet_str)

'Sleep(t_ms) funktsiooni kirjeldus: kui  0 < t_ms < 16 siis delay ~ 15 ms
'                                   kui 15 < t_ms < 32 siis delay ~ 30 ms
'                                   kui 31 < t_ms < 47 siis delay ~ 45 ms

'3.5 CHARACTER DELAY
If Form1.frame_delay = True Then ' If(baud=1200 or baud=2400)
   Sleep (31)
End If

Sleep (15)

'SEND PACKET
Form1.MSComm1.Output = packet_str
Form1.MSComm1.InBufferCount = 0

'DELAY
For delay_count = 0 To 50 ' TIMEOUT ~ 500 ms
   Sleep (15)
   If Form1.MSComm1.InBufferCount = 7 Then   ' 7 CHAR for reading!
      Exit For
   End If
Next

If Form1.MSComm1.InBufferCount = 7 Then
   packet_str = Form1.MSComm1.Input
   Call String_to_bytes(packet_str, packet_bytes(), 7)
   Call CRC_calc(packet_bytes, 5, crcH, crcL) ' kui loetakse üks register, siis väljastatakse 5 baidine pakett

   If packet_bytes(5) <> crcL Or packet_bytes(6) <> crcH Then ' crc kontroll
      modbus_read = "CRC ERR"
   Else
      data_out = packet_bytes(3) * 256# ' Bytes to Long calculation
      data_out = data_out + packet_bytes(4)
      modbus_read = CStr(data_out)
   End If
   
ElseIf Form1.MSComm1.InBufferCount = 5 Then
   packet_str = Form1.MSComm1.Input
   Call String_to_bytes(packet_str, packet_bytes(), 5)
   Call CRC_calc(packet_bytes, 3, crcH, crcL)
   
   If packet_bytes(3) <> crcL Or packet_bytes(4) <> crcH Then ' crc kontroll
      modbus_read = "CRC ERR"
   Else
      If packet_bytes(2) = 2 Then
         modbus_read = "ILLEGAL ADDRESS"
      Else
         modbus_read = "UNDEFINED ERROR"
      End If
   End If
Else
   modbus_read = "NO DEVICE"
End If

Exit Function
cannot_open_port:
modbus_read = "READ ERROR"
End Function
Private Sub Bytes_to_string(packet_as_bytes() As Byte, packet_as_string As String)

Dim str_adr     As String
Dim str_func    As String
Dim str_regH    As String
Dim str_regL    As String
Dim str_dataH   As String
Dim str_dataL   As String
Dim str_crcH    As String
Dim str_crcL    As String
Dim Saved    As String



str_adr = Chr$(packet_as_bytes(0))
str_func = Chr$(packet_as_bytes(1))
str_regH = Chr$(packet_as_bytes(2))
str_regL = Chr$(packet_as_bytes(3))
str_dataH = Chr$(packet_as_bytes(4))
str_dataL = Chr$(packet_as_bytes(5))
str_crcH = Chr$(packet_as_bytes(6))
str_crcL = Chr$(packet_as_bytes(7))


packet_as_string = str_adr + str_func + str_regH + str_regL + str_dataH + str_dataL + str_crcH + str_crcL


End Sub
Private Sub String_to_bytes(packet_as_string_out As String, packet_as_bytes_out() As Byte, nr_of_bytes As Integer)

Dim substring As String
Dim X As Byte
nr_of_bytes = nr_of_bytes - 1

For X = 0 To nr_of_bytes

X = X + 1
substring = Mid$(packet_as_string_out, X, 1) ' paketi x element salvestatakse substring-i
X = X - 1
packet_as_bytes_out(X) = Asc(substring) ' decimal nr into byte array

Next
End Sub
Private Sub CRC_calc(input_byte_array() As Byte, nr_of_bytes As Byte, crcH As Byte, crcL As Byte)

Dim CRC As Long
Dim abi As Long
Dim counter1 As Byte
Dim counter2 As Byte
Dim reminder As Byte

CRC = 65535
nr_of_bytes = nr_of_bytes - 1

For counter1 = 0 To nr_of_bytes
CRC = CRC Xor input_byte_array(counter1)

For counter2 = 0 To 7

reminder = CRC Mod 2
CRC = CRC \ 2 ' nihe paremale ehk "\" täisarvuline jagamine
If reminder = 1 Then
CRC = CRC Xor 40961
End If
Next

Next

crcH = CRC \ 256
abi = crcH * 256#
crcL = CRC - abi

End Sub
