Attribute VB_Name = "write"
Option Explicit
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long) ' Sleep "ms"
Public Function modbus_write(reg_adr As Long, data_in As Long) As String
On Error GoTo cannot_open_port

Dim packet_bytes(7) As Byte
Dim packet_str  As String
Dim abi As Long
Dim compare As String

Dim regH  As Byte
Dim regL  As Byte
Dim dataH As Byte
Dim dataL As Byte

Dim crcH  As Byte
Dim crcL  As Byte

Dim delay_count As Long

regH = reg_adr \ 256
abi = regH * 256#
regL = reg_adr - abi

dataH = data_in \ 256
abi = dataH * 256#
dataL = data_in - abi


packet_bytes(0) = Form1.slave_id_global
packet_bytes(1) = 6              ' write register, function=6
packet_bytes(2) = regH
packet_bytes(3) = regL
packet_bytes(4) = dataH          'data to write
packet_bytes(5) = dataL          'data to write

Call CRC_calc(packet_bytes(), 6, crcH, crcL) ' "6" baiti on vaja läbi töötada

packet_bytes(6) = crcL
packet_bytes(7) = crcH
'Convert Modbus packet dec numbers to characters for Transmitting
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
compare = packet_str

'DELAY
For delay_count = 0 To 50 ' TIMEOUT ~ 500 ms
   Sleep (15)
   If Form1.MSComm1.InBufferCount = 8 Then   ' 8 CHAR for writing!
      Exit For
   End If
Next


If Form1.MSComm1.InBufferCount = 8 Then
   packet_str = Form1.MSComm1.Input
   Call String_to_bytes(packet_str, packet_bytes(), 8) ' kirjutamisel saadakse vastu täpselt sama pakett mis saadeti PC poolt
   Call CRC_calc(packet_bytes, 6, crcH, crcL) ' kui loetakse üks register, siis RTU seade väljastab 7 baidise paketi
   If packet_bytes(6) <> crcL Or packet_bytes(7) <> crcH Then ' CRC kontroll
      modbus_write = "CRC ERR"
   ElseIf packet_str <> compare Then
      modbus_write = "MISMATCH"
   Else
      modbus_write = "WRITE COMPLETE"
   End If
ElseIf Form1.MSComm1.InBufferCount = 5 Then
   Call String_to_bytes(packet_str, packet_bytes(), 5)
   Call CRC_calc(packet_bytes, 3, crcH, crcL)
   
   packet_str = Form1.MSComm1.Input
   Call String_to_bytes(packet_str, packet_bytes(), 5)
   Call CRC_calc(packet_bytes, 3, crcH, crcL)
   
   If packet_bytes(3) <> crcL Or packet_bytes(4) <> crcH Then ' crc kontroll
      modbus_write = "CRC ERR"
   Else
      If packet_bytes(2) = 2 Then
         modbus_write = "ILLEGAL ADDRESS"
      ElseIf packet_bytes(2) = 3 Then
         modbus_write = "ILLEGAL VALUE"
      ElseIf packet_bytes(2) = 4 Then
         modbus_write = "DEVICE FAILURE"
      End If
   End If
Else
   modbus_write = "NO DEVICE"
End If

Exit Function
'error handling: if input data invalid; port disconnected
cannot_open_port:
modbus_write = "WRITE ERROR"
End Function
Public Function write_no_respond(reg_adr As Long, data_in As Long) As String
On Error GoTo cannot_open_port

Dim packet_bytes(7) As Byte
Dim packet_str  As String
Dim abi As Long

Dim regH  As Byte
Dim regL  As Byte
Dim dataH As Byte
Dim dataL As Byte

Dim crcH  As Byte
Dim crcL  As Byte

regH = reg_adr \ 256
abi = regH * 256#
regL = reg_adr - abi

dataH = data_in \ 256
abi = dataH * 256#
dataL = data_in - abi


packet_bytes(0) = Form1.slave_id_global
packet_bytes(1) = 6              ' write register, function=6
packet_bytes(2) = regH
packet_bytes(3) = regL
packet_bytes(4) = dataH          'data to write
packet_bytes(5) = dataL          'data to write

Call CRC_calc(packet_bytes(), 6, crcH, crcL) ' "6" baiti on vaja läbi töötada

packet_bytes(6) = crcL
packet_bytes(7) = crcH
'Convert Modbus packet dec numbers to characters for Transmitting
Call Bytes_to_string(packet_bytes(), packet_str)

'DELAY
If Form1.frame_delay = True Then
Sleep (200)
Else
Sleep (31)
End If

'SEND PACKET
Form1.MSComm1.Output = packet_str

write_no_respond = "WRITE COMPLETE"

Exit Function
'error handling: if input data invalid; port disconnected
cannot_open_port:
write_no_respond = "WRITE ERROR"
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
Dim x As Byte
nr_of_bytes = nr_of_bytes - 1

For x = 0 To nr_of_bytes

x = x + 1
substring = Mid$(packet_as_string_out, x, 1) ' paketi x element salvestatakse substring-i
x = x - 1
packet_as_bytes_out(x) = Asc(substring) ' decimal nr into byte array

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
