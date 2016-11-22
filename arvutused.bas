Attribute VB_Name = "arvutused"
Option Explicit
Public Function arv_taiendkoodiks(arv_sisse As Long) As Long

Dim abi1 As Long
abi1 = 65536
If arv_sisse < 0 Then
arv_taiendkoodiks = abi1 + arv_sisse
ElseIf arv_sisse > 0 Then
arv_taiendkoodiks = arv_sisse
Else
arv_taiendkoodiks = 0
End If
 
End Function
Public Function taiendkood_arvuks(arv_sisse As Long) As Long

Dim abi1 As Long
abi1 = 65536
If arv_sisse > 32767 Then
taiendkood_arvuks = arv_sisse - abi1
ElseIf arv_sisse < 32768 Then
taiendkood_arvuks = arv_sisse
End If

End Function
Public Function analog_scale_calc(I1 As Long, I2 As Long, T1 As Long, T2 As Long, Current_Voltage As String, MAX_MIN_value As String) As Long
Dim delta_I As Long
Dim delta_T As Long
Dim tõus As Single
Dim konstant As Single
Dim abi As Single

delta_I = I2 - I1
delta_T = T2 - T1

tõus = delta_I / delta_T

abi = T2 * tõus
konstant = I2 - abi

' Current
If Current_Voltage = "Current" Then
   
   If MAX_MIN_value = "MAX" Then
      '100% output
      abi = 20 - konstant
      analog_scale_calc = abi / tõus
   ElseIf MAX_MIN_value = "MIN" Then
      '0% output
      abi = 4 - konstant
      analog_scale_calc = abi / tõus
   End If
' Voltage
ElseIf Current_Voltage = "Voltage" Then
    
   If MAX_MIN_value = "MAX" Then
      '100% output
      abi = 10 - konstant
      analog_scale_calc = abi / tõus
   ElseIf MAX_MIN_value = "MIN" Then
      '0% output
      abi = 0 - konstant
      analog_scale_calc = abi / tõus
   End If
   
End If

End Function
Public Function status_reg_read(reg_in As Long, bit_to_test As String) As String
'test_sensor, test_analog, test_jumper1, test_jumper2, diagnostics1, diagnostics2
Dim reminder As Integer

If bit_to_test = "test_sensor" Then
   reminder = reg_in Mod 2
   If reminder = 1 Then
   status_reg_read = "NO SENS"
   End If
ElseIf bit_to_test = "test_analog" Then
   reminder = reg_in \ 2
   reminder = reminder Mod 2
   If reminder = 1 Then
   status_reg_read = "ON"
   Else
   status_reg_read = "OFF"
   End If
ElseIf bit_to_test = "test_jumper1" Then
   reminder = reg_in \ 64
   reminder = reminder Mod 2
   If reminder = 1 Then
   status_reg_read = "Voltage"
   Else
   status_reg_read = "Current"
   End If
ElseIf bit_to_test = "test_jumper2" Then
   reminder = reg_in \ 128
   reminder = reminder Mod 2
   If reminder = 1 Then
   status_reg_read = "Voltage"
   Else
   status_reg_read = "Current"
   End If
ElseIf bit_to_test = "diagnostics1" Then
   reminder = reg_in \ 4
   reminder = reminder Mod 2
   If reminder = 1 Then
      reminder = reg_in \ 8
      reminder = reminder Mod 2
      If reminder = 1 Then
         status_reg_read = "D:21.5mA"
      Else
         status_reg_read = "D: 3.5mA"
      End If
   Else
   status_reg_read = "OFF"
   End If
ElseIf bit_to_test = "diagnostics2" Then
   reminder = reg_in \ 16
   reminder = reminder Mod 2
   If reminder = 1 Then
      reminder = reg_in \ 32
      reminder = reminder Mod 2
      If reminder = 1 Then
         status_reg_read = "D:21.5mA"
      Else
         status_reg_read = "D: 3.5mA"
      End If
   Else
   status_reg_read = "OFF"
   End If
ElseIf bit_to_test = "test_LED" Then
   reminder = reg_in \ 256
   reminder = reminder Mod 2
   If reminder = 1 Then
   status_reg_read = "ON"
   Else
   status_reg_read = "OFF"
   End If
ElseIf bit_to_test = "test_buzzer" Then
   reminder = reg_in \ 512
   reminder = reminder Mod 2
   If reminder = 1 Then
   status_reg_read = "ON"
   Else
   status_reg_read = "OFF"
   End If

End If ' end of bit testing

End Function

