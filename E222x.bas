Attribute VB_Name = "E222x"
Option Explicit
Public Function E222x_read_all() As String
Dim output_string As String
Dim output_long As Long
Dim output_single As Single



If Form1.Text_hardware.Visible = False Then ' Factory key
 If Form1.Frame_E222X.Caption = "E26XX" Then
    output_string = modbus_read(1) 'KONTROLL
    If Not IsNumeric(output_string) Then
      E222x_read_all = output_string ' funktsioon väljastab error sõnumi
      Exit Function
    ElseIf Mid$(output_string, 1, 4) <> "2608" Then ' register peab sisaldama 26xx seadme numbrit
      E222x_read_all = "WRONG DEVICE"
      Exit Function
    End If
 End If
 If Form1.Frame_E222X.Caption = "E24XX" Then
    output_string = modbus_read(1) 'KONTROLL
    If Not IsNumeric(output_string) Then
      E222x_read_all = output_string ' funktsioon väljastab error sõnumi
      Exit Function
    ElseIf Mid$(output_string, 1, 2) <> "24" Then ' register peab sisaldama 26xx seadme numbrit
      E222x_read_all = "WRONG DEVICE"
      Exit Function
    End If
 End If
 If Form1.Frame_E222X.Caption = "E27XX" Then
    output_string = modbus_read(1) 'KONTROLL
    If Not IsNumeric(output_string) Then
      E222x_read_all = output_string ' funktsioon väljastab error sõnumi
      Exit Function
    ElseIf Mid$(output_string, 1, 2) <> "27" Then ' register peab sisaldama 26xx seadme numbrit
      E222x_read_all = "WRONG DEVICE"
      Exit Function
    End If
 End If
 If Form1.Frame_E222X.Caption = "E22XX" Then
    output_string = modbus_read(1) 'KONTROLL
    If Not IsNumeric(output_string) Then
      E222x_read_all = output_string ' funktsioon väljastab error sõnumi
      Exit Function
    ElseIf Mid$(output_string, 1, 2) <> "22" Then ' register peab sisaldama 26xx seadme numbrit
      E222x_read_all = "WRONG DEVICE"
      Exit Function
    End If
 End If
 If Form1.Frame_E222X.Caption = "PVT100" Then
    output_string = modbus_read(1) 'KONTROLL
    If Not IsNumeric(output_string) Then
      E222x_read_all = output_string ' funktsioon väljastab error sõnumi
      Exit Function
    ElseIf Mid$(output_string, 1, 2) <> "20" Then ' register peab sisaldama 26xx seadme numbrit
      E222x_read_all = "WRONG DEVICE"
      Exit Function
    End If
 End If
 If Form1.Frame_E222X.Caption = "PVT10" Then
    output_string = modbus_read(1) 'KONTROLL
    If Not IsNumeric(output_string) Then
      E222x_read_all = output_string ' funktsioon väljastab error sõnumi
      Exit Function
    ElseIf Mid$(output_string, 1, 2) <> "20" Then ' register peab sisaldama 26xx seadme numbrit
      E222x_read_all = "WRONG DEVICE"
      Exit Function
    End If
 End If
ElseIf Form1.Frame_E222X.Caption = "E26XX" Then
   output_string = modbus_read(145)    'HEATER PULSE
   If IsNumeric(output_string) Then
   Form1.Label_heater_pulse.Caption = output_string
   Else
   GoTo reading_error
   End If
   'SENSOR PULSE
   output_string = modbus_read(146)
   If IsNumeric(output_string) Then
   Form1.Label_sensor_pulse.Caption = output_string
   Else
   GoTo reading_error
   End If
   'PARAMETER B
   output_string = modbus_read(147)
   If IsNumeric(output_string) Then
   Form1.Label_const_B.Caption = output_string
   Else
   GoTo reading_error
   End If
   'PARAMETER C
   output_string = modbus_read(148)
   If IsNumeric(output_string) Then
   Form1.Label_const_C.Caption = output_string
   Else
   GoTo reading_error
   End If
   'PARAMETER D
   output_string = modbus_read(149)
   If IsNumeric(output_string) Then
   Form1.Label_const_D.Caption = output_string
   Else
   GoTo reading_error
   End If
   'PARAMETER E
   output_string = modbus_read(150)
   If IsNumeric(output_string) Then
   Form1.Label_const_E.Caption = output_string
   Else
   GoTo reading_error
   End If
   'SENSOR TYPE + GAS UNITS
   output_string = modbus_read(151)
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If output_long > 32767 Then
         output_long = output_long - 32768
      End If
      If output_long > 16383 Then
         output_long = output_long - 16384
      End If
      Select Case output_long
         Case 0
            Form1.Label_gas_type.Caption = "CH4"
         Case 1
            Form1.Label_gas_type.Caption = "CO"
         Case 2
            Form1.Label_gas_type.Caption = "O2"
         Case 3
            Form1.Label_gas_type.Caption = "NH3"
         Case 4
            Form1.Label_gas_type.Caption = "H2"
         Case 5
            Form1.Label_gas_type.Caption = "VOC"
         Case 6
            Form1.Label_gas_type.Caption = "LPG"
         Case 7
            Form1.Label_gas_type.Caption = "HFC"
         Case 8
            Form1.Label_gas_type.Caption = "O3"
         Case 9
            Form1.Label_gas_type.Caption = "H2S"
         Case 10
            Form1.Label_gas_type.Caption = "HCL"
         Case 11
            Form1.Label_gas_type.Caption = "CL2"
         Case 12
            Form1.Label_gas_type.Caption = "SO2"
         Case 13
            Form1.Label_gas_type.Caption = "S2H4"
         Case 14
            Form1.Label_gas_type.Caption = "ETO"
         Case 15
            Form1.Label_gas_type.Caption = "NO"
         Case 16
            Form1.Label_gas_type.Caption = "NO2"
      End Select
      output_long = Val(output_string)
      output_long = output_long \ 16384 ' 14 nihet paremale
      Select Case output_long
         Case 0
            Form1.Label_gas_units.Caption = "ppm"
            Form1.Label_meas_gas_units.Caption = "ppm"
         Case 1
            Form1.Label_gas_units.Caption = "‰"
            Form1.Label_meas_gas_units.Caption = "‰"
         Case 2
            Form1.Label_gas_units.Caption = "%"
            Form1.Label_meas_gas_units.Caption = "%"
      End Select
   Else
   GoTo reading_error
   End If
   ' TEMPERATURE ZERO ADJ.
    output_string = modbus_read(162)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 100
    Form1.Label_zero_T.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' GAS ZERO ADJ.
    output_string = modbus_read(165)
    If IsNumeric(output_string) Then
    output_long = taiendkood_arvuks(Val(output_string))
    Form1.Label_zero_RH.Caption = output_long
    Else
    GoTo reading_error
    End If
    ' GAS SLOPE ADJ.
    output_string = modbus_read(166)
    If IsNumeric(output_string) Then
    Form1.Label_RH_slope.Caption = output_string
    Else
    GoTo reading_error
    End If
    ' GAS RATE ADJ.
    output_string = modbus_read(167)
    If IsNumeric(output_string) Then
    Form1.Label_RH_rate.Caption = output_string
    Else
    GoTo reading_error
    End If
    'INTEGRATING TIME CONSTANT
    output_string = modbus_read(168)
    If IsNumeric(output_string) Then
    Form1.Label_RC_filter.Caption = output_string
    Else
    GoTo reading_error
    End If
ElseIf Form1.Frame_E222X.Caption = "E22XX" Then
    output_string = modbus_read(162) ' TEMPERATURE ZERO ADJ.
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 100
    Form1.Label_zero_T.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(153) ' K
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_B.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(254) ' Tm
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_C.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(256) ' Ts
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_D.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' TEMPERATURE SLOPE ADJ.
    output_string = modbus_read(163)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 1000
    Form1.Label_sensor_pulse.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' TEMPERATURE RATE ADJ.
    output_string = modbus_read(164)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_heater_pulse.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY ZERO ADJ.
    output_string = modbus_read(165)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 100
    Form1.Label_zero_RH.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY SLOPE ADJ.
    output_string = modbus_read(166)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 1000
    Form1.Label_RH_slope.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY RATE ADJ.
    output_string = modbus_read(167)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_RH_rate.Caption = output_single
    Else
    GoTo reading_error
    End If
    'INTEGRATING TIME CONSTANT
    output_string = modbus_read(168)
    If IsNumeric(output_string) Then
    Form1.Label_RC_filter.Caption = output_string
    Else
    GoTo reading_error
    End If
    
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
    output_string = modbus_read(162) ' TEMPERATURE ZERO ADJ.
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 100
    Form1.Label_zero_T.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(153) ' K
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_B.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(254) ' Tm
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_C.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(256) ' Ts
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_D.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' TEMPERATURE SLOPE ADJ.
    output_string = modbus_read(163)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 1000
    Form1.Label_sensor_pulse.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' TEMPERATURE RATE ADJ.
    output_string = modbus_read(164)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_heater_pulse.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY ZERO ADJ.
    output_string = modbus_read(165)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 100
    Form1.Label_zero_RH.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY SLOPE ADJ.
    output_string = modbus_read(166)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 1000
    Form1.Label_RH_slope.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY RATE ADJ.
    output_string = modbus_read(167)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_RH_rate.Caption = output_single
    Else
    GoTo reading_error
    End If
    'INTEGRATING TIME CONSTANT
    output_string = modbus_read(168)
    If IsNumeric(output_string) Then
    Form1.Label_RC_filter.Caption = output_string
    Else
    GoTo reading_error
    End If
    
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
    output_string = modbus_read(162) ' TEMPERATURE ZERO ADJ.
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 100
    Form1.Label_zero_T.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(153) ' K
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_B.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(254) ' Tm
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_C.Caption = output_single
    Else
    GoTo reading_error
    End If
    output_string = modbus_read(256) ' Ts
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long
    Form1.Label_const_D.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' TEMPERATURE SLOPE ADJ.
    output_string = modbus_read(163)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 1000
    Form1.Label_sensor_pulse.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' TEMPERATURE RATE ADJ.
    output_string = modbus_read(164)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_heater_pulse.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY ZERO ADJ.
    output_string = modbus_read(165)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 100
    Form1.Label_zero_RH.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY SLOPE ADJ.
    output_string = modbus_read(166)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_long = taiendkood_arvuks(output_long)
    output_single = output_long / 1000
    Form1.Label_RH_slope.Caption = output_single
    Else
    GoTo reading_error
    End If
    ' HUMIDITY RATE ADJ.
    output_string = modbus_read(167)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_RH_rate.Caption = output_single
    Else
    GoTo reading_error
    End If
    'INTEGRATING TIME CONSTANT
    output_string = modbus_read(168)
    If IsNumeric(output_string) Then
    Form1.Label_RC_filter.Caption = output_string
    Else
    GoTo reading_error
    End If
    
End If

'HARDWARE
output_string = modbus_read(1)
If IsNumeric(output_string) Then
Form1.Label_hardware.Caption = output_string
Else
GoTo reading_error
End If
 If Form1.Label_hardware.Caption = 20566 Then
Form1.Label_hardware.Caption = "PV"
 End If
'SOFTWARE
output_string = modbus_read(2)
If IsNumeric(output_string) Then
Form1.Label_software.Caption = output_string
Else
GoTo reading_error
End If
 If Form1.Label_software.Caption = 21553 Then
Form1.Label_software.Caption = "T1"
 End If
' SN
output_string = modbus_read(3)
If IsNumeric(output_string) Then
Form1.Label_SN.Caption = output_string
Else
GoTo reading_error
End If
 If Form1.Label_SN.Caption = 12336 Then
 Form1.Label_SN.Caption = "00"
 ElseIf Form1.Label_SN.Caption = 12288 Then
 Form1.Label_SN.Caption = "0"
 End If
'RESPONSE DELAY
output_string = modbus_read(6)
If IsNumeric(output_string) Then
Form1.Label_response.Caption = output_string
Else
GoTo reading_error
End If
' SLAVE ID
output_string = modbus_read(4)
If IsNumeric(output_string) Then
Form1.Label_slave_id.Caption = output_string
Else
GoTo reading_error
End If
' BAUD RATE
output_string = modbus_read(5)
If IsNumeric(output_string) Then
Form1.Label_baud.Caption = output_string
Else
GoTo reading_error
End If
' STOP BITS
output_string = modbus_read(7)
If IsNumeric(output_string) Then
Form1.Label_stop_bit = output_string
Else
GoTo reading_error
End If
If Form1.Frame_E222X.Caption = "E26XX" Then
'AN1 OUTPUT PARAMETER
 output_string = modbus_read(201)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN1_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN1_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN1_onoff.Caption = "GAS"
   ElseIf output_string = "9" Then
   Form1.Label_AN1_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
'AN2 OUTPUT PARAMETER
 output_string = modbus_read(202)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN2_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN2_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN2_onoff.Caption = "GAS"
   ElseIf output_string = "9" Then
   Form1.Label_AN2_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
'AN1 OUTPUT PARAMETER
 output_string = modbus_read(201)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN1_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN1_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN1_onoff.Caption = "RH"
   ElseIf output_string = "3" Then
   Form1.Label_AN1_onoff.Caption = "DEW"
   ElseIf output_string = "9" Then
   Form1.Label_AN1_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
'AN2 OUTPUT PARAMETER
 output_string = modbus_read(202)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN2_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN2_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN2_onoff.Caption = "RH"
   ElseIf output_string = "3" Then
   Form1.Label_AN2_onoff.Caption = "DEW"
   ElseIf output_string = "9" Then
   Form1.Label_AN2_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
'AN1 OUTPUT PARAMETER
 output_string = modbus_read(201)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN1_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN1_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN1_onoff.Caption = "RH"
   ElseIf output_string = "3" Then
   Form1.Label_AN1_onoff.Caption = "DEW"
   ElseIf output_string = "9" Then
   Form1.Label_AN1_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
'AN2 OUTPUT PARAMETER
 output_string = modbus_read(202)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN2_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN2_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN2_onoff.Caption = "RH"
   ElseIf output_string = "3" Then
   Form1.Label_AN2_onoff.Caption = "DEW"
   ElseIf output_string = "9" Then
   Form1.Label_AN2_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
ElseIf Form1.Frame_E222X.Caption = "E22XX" Then
'AN1 OUTPUT PARAMETER
 output_string = modbus_read(201)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN1_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN1_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN1_onoff.Caption = "RH"
   ElseIf output_string = "3" Then
   Form1.Label_AN1_onoff.Caption = "DEW"
   ElseIf output_string = "9" Then
   Form1.Label_AN1_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
'AN2 OUTPUT PARAMETER
 output_string = modbus_read(202)
 If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_AN2_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_AN2_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_AN2_onoff.Caption = "RH"
   ElseIf output_string = "3" Then
   Form1.Label_AN2_onoff.Caption = "DEW"
   ElseIf output_string = "9" Then
   Form1.Label_AN2_onoff.Caption = "Modbus"
   End If
 Else
 GoTo reading_error
 End If
End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' STATUS REG PROCESSING
output_string = modbus_read(255)
If IsNumeric(output_string) Then
output_long = Val(output_string)
Form1.Label_global_AN.Caption = status_reg_read(output_long, "test_analog")
Form1.Label_AN1_I_U.Caption = status_reg_read(output_long, "test_jumper1")
Form1.Label_AN2_I_U.Caption = status_reg_read(output_long, "test_jumper2")
Form1.Label_AN1_diag.Caption = status_reg_read(output_long, "diagnostics1")
Form1.Label_AN2_diag.Caption = status_reg_read(output_long, "diagnostics2")
Form1.Label_buzzer.Caption = status_reg_read(output_long, "test_buzzer")
Form1.Label_LED.Caption = status_reg_read(output_long, "test_LED")
'SENSOR TEST*****************************************************************************************
If status_reg_read(output_long, "test_sensor") <> "NO SENS" Then
   
   'ATTENTION!!! output_long manipulation occurs!!!
   ' MEASURED TEMP
   output_string = modbus_read(258)
   If IsNumeric(output_string) Then
   output_long = Val(output_string)
   output_long = taiendkood_arvuks(output_long)
   output_single = output_long / 100
   Form1.Label_temp.Caption = output_single
   Else
   GoTo reading_error
   End If
   
   
   
   'E26XX
    If Form1.Frame_E222X.Caption = "E26XX" Then
   ' MEASURED CONCENTRATION
   output_string = modbus_read(259)
   If IsNumeric(output_string) Then
   Form1.Label_hum.Caption = output_string
   Else
   GoTo reading_error
   End If
   End If
   
   
   'E22XX
   ' MEASURED HUMIDITY
    If Form1.Frame_E222X.Caption = "E22XX" Then
      output_string = modbus_read(259)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_hum.Caption = output_single
    Else
    GoTo reading_error
    End If
    End If
   
   'PVT100
   ' MEASURED HUMIDITY
    If Form1.Frame_E222X.Caption = "PVT100" Then
      output_string = modbus_read(259)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_hum.Caption = output_single
    Else
    GoTo reading_error
    End If
    End If
   
   'PVT10
   ' MEASURED HUMIDITY
    If Form1.Frame_E222X.Caption = "PVT10" Then
      output_string = modbus_read(259)
    If IsNumeric(output_string) Then
    output_long = Val(output_string)
    output_single = output_long / 100
    Form1.Label_hum.Caption = output_single
    Else
    GoTo reading_error
    End If
    End If
   
   'E26XX
   If Form1.Frame_E222X.Caption = "E26XX" Then
   ' MEASURED ADC
   output_string = modbus_read(257)
   If IsNumeric(output_string) Then
   Form1.Label30.Caption = output_string
   Else
   GoTo reading_error
   End If
   End If
   
   
   'E22XX
   If Form1.Frame_E222X.Caption = "E22XX" Then
   ' MEASURED DEWPOINT
        output_string = modbus_read(260)
   If IsNumeric(output_string) Then
   output_long = Val(output_string)
   output_long = taiendkood_arvuks(output_long)
   output_single = output_long / 100
   Form1.Label30.Caption = output_single
   Else
   GoTo reading_error
   End If
   End If
   
    'PVT100
   If Form1.Frame_E222X.Caption = "PVT100" Then
   ' MEASURED DEWPOINT
        output_string = modbus_read(260)
   If IsNumeric(output_string) Then
   output_long = Val(output_string)
   output_long = taiendkood_arvuks(output_long)
   output_single = output_long / 100
   Form1.Label30.Caption = output_single
   Else
   GoTo reading_error
   End If
   End If
   
   'PVT10
   If Form1.Frame_E222X.Caption = "PVT10" Then
   ' MEASURED DEWPOINT
        output_string = modbus_read(260)
   If IsNumeric(output_string) Then
   output_long = Val(output_string)
   output_long = taiendkood_arvuks(output_long)
   output_single = output_long / 100
   Form1.Label30.Caption = output_single
   Else
   GoTo reading_error
   End If
   End If
   
Else
   Form1.Label_temp.Caption = "absent"
   Form1.Label_hum.Caption = "absent"
End If
'SENSOR TEST*****************************************************************************************
Else
GoTo reading_error
End If ' end of STATUS REG PROCESSING
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ANALOG1 0% VALUE.
output_string = modbus_read(261)
If IsNumeric(output_string) Then
output_long = Val(output_string)
Form1.Label_AN1_0_value = taiendkood_arvuks(output_long)
Else
GoTo reading_error
End If
' ANALOG1 100% VALUE.
output_string = modbus_read(262)
If IsNumeric(output_string) Then
output_long = Val(output_string)
Form1.Label_AN1_100_value = taiendkood_arvuks(output_long)
Else
GoTo reading_error
End If
' ANALOG2 0% VALUE.
output_string = modbus_read(263)
If IsNumeric(output_string) Then
output_long = Val(output_string)
Form1.Label_AN2_0_value = taiendkood_arvuks(output_long)
Else
GoTo reading_error
End If
' ANALOG2 100% VALUE.
output_string = modbus_read(264)
If IsNumeric(output_string) Then
output_long = Val(output_string)
Form1.Label_AN2_100_value = taiendkood_arvuks(output_long)
Else
GoTo reading_error
End If
' RELAY1 PARAMETER
output_string = modbus_read(211)
If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_RE1_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_RE1_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_RE1_onoff.Caption = "GAS"
   ElseIf output_string = "9" Then
   Form1.Label_RE1_onoff.Caption = "Modbus"
   End If
Else
GoTo reading_error
End If
' RELAY2 PARAMETER
output_string = modbus_read(212)
If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_RE2_onoff.Caption = "OFF"
   ElseIf output_string = "1" Then
   Form1.Label_RE2_onoff.Caption = "TEMP"
   ElseIf output_string = "2" Then
   Form1.Label_RE2_onoff.Caption = "GAS"
   ElseIf output_string = "9" Then
   Form1.Label_RE2_onoff.Caption = "Modbus"
   End If
Else
GoTo reading_error
End If
' RELAY1 MODE
output_string = modbus_read(219)
If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_RE1_mode.Caption = "NONE"
   ElseIf output_string = "1" Then
   Form1.Label_RE1_mode.Caption = "high values"
   ElseIf output_string = "2" Then
   Form1.Label_RE1_mode.Caption = "low values"
   ElseIf output_string = "3" Then
   Form1.Label_RE1_mode.Caption = "within range"
   ElseIf output_string = "4" Then
   Form1.Label_RE1_mode.Caption = "outside range"
   End If
Else
GoTo reading_error
End If
' RELAY2 MODE
output_string = modbus_read(220)
If IsNumeric(output_string) Then
   If output_string = "0" Then
   Form1.Label_RE2_mode.Caption = "NONE"
   ElseIf output_string = "1" Then
   Form1.Label_RE2_mode.Caption = "high values"
   ElseIf output_string = "2" Then
   Form1.Label_RE2_mode.Caption = "low values"
   ElseIf output_string = "3" Then
   Form1.Label_RE2_mode.Caption = "within range"
   ElseIf output_string = "4" Then
   Form1.Label_RE2_mode.Caption = "outside range"
   End If
Else
GoTo reading_error
End If
' RELAY1 LOW
output_string = modbus_read(221)
If IsNumeric(output_string) Then
output_long = Val(output_string)
output_long = taiendkood_arvuks(output_long)
output_single = output_long / 100
Form1.Label_RE1_L = output_single
Else
GoTo reading_error
End If
' RELAY1 HIGH
output_string = modbus_read(222)
If IsNumeric(output_string) Then
output_long = Val(output_string)
output_long = taiendkood_arvuks(output_long)
output_single = output_long / 100
Form1.Label_RE1_H = output_single
Else
GoTo reading_error
End If
' RELAY2 LOW
output_string = modbus_read(223)
If IsNumeric(output_string) Then
output_long = Val(output_string)
output_long = taiendkood_arvuks(output_long)
output_single = output_long / 100
Form1.Label_RE2_L = output_single
Else
GoTo reading_error
End If
' RELAY2 HIGH
output_string = modbus_read(224)
If IsNumeric(output_string) Then
output_long = Val(output_string)
output_long = taiendkood_arvuks(output_long)
output_single = output_long / 100
Form1.Label_RE2_H = output_single
Else
GoTo reading_error
End If
' RELAY1 delay
output_string = modbus_read(215)
If IsNumeric(output_string) Then
Form1.Label_RE1_delay.Caption = output_string
Else
GoTo reading_error
End If
' RELAY2 delay
output_string = modbus_read(216)
If IsNumeric(output_string) Then
Form1.Label_RE2_delay.Caption = output_string
Else
GoTo reading_error
End If
' RELAY1 minimal on/off
output_string = modbus_read(217)
If IsNumeric(output_string) Then
Form1.Label_RE1_time.Caption = output_string
Else
GoTo reading_error
End If
' RELAY2 minimal on/off
output_string = modbus_read(218)
If IsNumeric(output_string) Then
Form1.Label_RE2_time.Caption = output_string
Else
GoTo reading_error
End If

E222x_read_all = "READ SUCCESSFUL"
Exit Function

reading_error:
E222x_read_all = output_string

End Function
'***************************************************************************************************
'*******************************************WRITING ROUTINE*****************************************
'***************************************************************************************************
Public Function E222x_write_all() As String
Dim output_string As String
Dim input_long As Long
Dim input_single As Single

If Form1.Text_factory.Text = "0xA55A" Then ' FACTORY KEY
  If Form1.Frame_E222X.Caption = "E26XX" Then
   'PASSWORD "A55Ah"
   output_string = modbus_write(255, 42330)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
   'HARDWARE
   output_string = modbus_write(1, Val(Form1.Text_hardware.Text))
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
   'SN
   output_string = modbus_write(3, Val(Form1.Text_SN.Text))
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
   'HEATER VOLTAGE
 If Form1.Text_heater_pulse.Visible = True Then
   output_string = modbus_write(145, Val(Form1.Text_heater_pulse.Text))
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
 End If
   'SENSOR VOLTAGE
 If Form1.Text_sensor_pulse.Visible = True Then
   output_string = modbus_write(146, Val(Form1.Text_sensor_pulse.Text))
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
 End If
   ' PARAMETER B
 If Form1.Text_const_B.Visible = True Then
   output_string = modbus_write(147, Val(Form1.Text_const_B.Text))
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
 End If
   ' PARAMETER C
 If Form1.Text_const_C.Visible = True Then
   output_string = modbus_write(148, Val(Form1.Text_const_C.Text))
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
 End If
   ' PARAMETER D
 If Form1.Text_const_D.Visible = True Then
   input_long = arv_taiendkoodiks(Val(Form1.Text_const_D.Text))
   output_string = modbus_write(149, input_long)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
 End If
   ' PARAMETER E
 If Form1.Text_const_E.Visible = True Then
   input_long = arv_taiendkoodiks(Val(Form1.Text_const_E.Text))
   output_string = modbus_write(150, input_long)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
 End If
   'SENSOR TYPE + GAS UNITS
 If Form1.Combo_gas_units.Visible = True And Form1.Combo_gas_type.Visible = True Then
 
   input_long = Form1.Combo_gas_units.ListIndex * 16384#
   Select Case Form1.Combo_gas_type.ListIndex
      Case 0
         input_long = input_long + 0
      Case 1
         input_long = input_long + 1
      Case 2
         input_long = input_long + 2
      Case 3
         input_long = input_long + 3
      Case 4
         input_long = input_long + 4
      Case 5
         input_long = input_long + 5
      Case 6
         input_long = input_long + 6
      Case 7
         input_long = input_long + 7
      Case 8
         input_long = input_long + 8
      Case 9
         input_long = input_long + 9
      Case 10
         input_long = input_long + 10
      Case 11
         input_long = input_long + 11
      Case 12
         input_long = input_long + 12
      Case 13
         input_long = input_long + 13
      Case 14
         input_long = input_long + 14
      Case 15
         input_long = input_long + 15
      Case 16
         input_long = input_long + 16
   End Select
   output_string = modbus_write(151, input_long)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
 End If
   'ZERO ADJ TEMP
   input_single = CDbl(Form1.Text_zero_T.Text)
   input_long = input_single * 100
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(162, input_long)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
   'ZERO ADJ GAS
   input_long = CDbl(Form1.Text_zero_RH.Text)
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(165, input_long)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
   'SLOPE GAS
   input_long = CDbl(Form1.Text_RH_slope.Text)
   output_string = modbus_write(166, input_long)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
   'RATE GAS
   input_long = CDbl(Form1.Text_RH_rate.Text)
   output_string = modbus_write(167, input_long)
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If
   'INTEGRATING FILTER
   output_string = modbus_write(168, Val(Form1.Text_RC_filter.Text))
   If output_string <> "WRITE COMPLETE" Then
   GoTo writing_error
   End If

  ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
   output_string = modbus_write(255, 42330)
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'HARDWARE
    output_string = modbus_write(1, Val(Form1.Text_hardware.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SN
    output_string = modbus_write(3, Val(Form1.Text_SN.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SW
    output_string = modbus_write(2, Val(Form1.Text_SW.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'K
    output_string = modbus_write(153, Val(Form1.Text_const_B.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'ZERO ADJ TEMP
    input_single = CDbl(Form1.Text_zero_T.Text)
    input_long = input_single * 100
    input_long = arv_taiendkoodiks(input_long)
    output_string = modbus_write(162, input_long)
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SLOPE TEMP
     input_single = CDbl(Form1.Text_sensor_pulse.Text)
     input_long = input_single * 1000
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(163, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'RATE TEMP
     input_single = CDbl(Form1.Text_heater_pulse.Text)
     input_long = input_single * 100
     output_string = modbus_write(164, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'ZERO ADJ HUM
     input_single = CDbl(Form1.Text_zero_RH.Text)
     input_long = input_single * 100
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(165, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'SLOPE HUM
     input_single = CDbl(Form1.Text_RH_slope.Text)
     input_long = input_single * 1000
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(166, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'RATE HUM
     input_single = CDbl(Form1.Text_RH_rate.Text)
     input_long = input_single * 100
     output_string = modbus_write(167, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
    'INTEGRATING FILTER
    output_string = modbus_write(168, Val(Form1.Text_RC_filter.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If

  ElseIf Form1.Frame_E222X.Caption = "E22XX" Then
   output_string = modbus_write(255, 42330)
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'HARDWARE
    output_string = modbus_write(1, Val(Form1.Text_hardware.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SN
    output_string = modbus_write(3, Val(Form1.Text_SN.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'K
    output_string = modbus_write(153, Val(Form1.Text_const_B.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'ZERO ADJ TEMP
    input_single = CDbl(Form1.Text_zero_T.Text)
    input_long = input_single * 100
    input_long = arv_taiendkoodiks(input_long)
    output_string = modbus_write(162, input_long)
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SLOPE TEMP
     input_single = CDbl(Form1.Text_sensor_pulse.Text)
     input_long = input_single * 1000
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(163, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'RATE TEMP
     input_single = CDbl(Form1.Text_heater_pulse.Text)
     input_long = input_single * 100
     output_string = modbus_write(164, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'ZERO ADJ HUM
     input_single = CDbl(Form1.Text_zero_RH.Text)
     input_long = input_single * 100
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(165, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'SLOPE HUM
     input_single = CDbl(Form1.Text_RH_slope.Text)
     input_long = input_single * 1000
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(166, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'RATE HUM
     input_single = CDbl(Form1.Text_RH_rate.Text)
     input_long = input_single * 100
     output_string = modbus_write(167, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
    'INTEGRATING FILTER
    output_string = modbus_write(168, Val(Form1.Text_RC_filter.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
   output_string = modbus_write(255, 42330)
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'HARDWARE
    output_string = modbus_write(1, Val(Form1.Text_hardware.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SN
    output_string = modbus_write(3, Val(Form1.Text_SN.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SW
    output_string = modbus_write(2, Val(Form1.Text_SW.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'K
    output_string = modbus_write(153, Val(Form1.Text_const_B.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'ZERO ADJ TEMP
    input_single = CDbl(Form1.Text_zero_T.Text)
    input_long = input_single * 100
    input_long = arv_taiendkoodiks(input_long)
    output_string = modbus_write(162, input_long)
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
    'SLOPE TEMP
     input_single = CDbl(Form1.Text_sensor_pulse.Text)
     input_long = input_single * 1000
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(163, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'RATE TEMP
     input_single = CDbl(Form1.Text_heater_pulse.Text)
     input_long = input_single * 100
     output_string = modbus_write(164, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'ZERO ADJ HUM
     input_single = CDbl(Form1.Text_zero_RH.Text)
     input_long = input_single * 100
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(165, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'SLOPE HUM
     input_single = CDbl(Form1.Text_RH_slope.Text)
     input_long = input_single * 1000
     input_long = arv_taiendkoodiks(input_long)
     output_string = modbus_write(166, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
     'RATE HUM
     input_single = CDbl(Form1.Text_RH_rate.Text)
     input_long = input_single * 100
     output_string = modbus_write(167, input_long)
     If output_string <> "WRITE COMPLETE" Then
     GoTo writing_error
     End If
    'INTEGRATING FILTER
    output_string = modbus_write(168, Val(Form1.Text_RC_filter.Text))
    If output_string <> "WRITE COMPLETE" Then
    GoTo writing_error
    End If
  End If
  Else
  'KONTROLL
   output_string = modbus_read(1)
   If Not IsNumeric(output_string) Then
      E222x_write_all = output_string ' funktsioon väljastab error sõnumi
      Exit Function
   ElseIf Mid$(output_string, 1, 2) <> "26" And Mid$(output_string, 1, 2) <> "22" And Mid$(output_string, 1, 2) <> "20" Then
      E222x_write_all = "WRONG DEVICE"
      Exit Function
   End If
   
End If

'SLAVE ID
output_string = modbus_write(4, Val(Form1.Text_slave_id.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If

'STATUS REG WRITING:
'***************************************************************************************************
input_long = 0
'global analog check
If Form1.Combo_global_AN.ListIndex = 0 Then
   input_long = input_long + 2
   
   ' SET analog diagnostics bits
   If Form1.Combo_AN1_diag.ListIndex = 0 Then
      input_long = input_long + 4 + 8
   End If
   If Form1.Combo_AN1_diag.ListIndex = 1 Then
      input_long = input_long + 4 '+ 0
   End If
   
   If Form1.Combo_AN2_diag.ListIndex = 0 Then
      input_long = input_long + 16 + 32
   End If
   
   If Form1.Combo_AN2_diag.ListIndex = 1 Then
      input_long = input_long + 16 '+ 0
   End If

End If

If Form1.Combo_LED.ListIndex = 0 Then
   input_long = input_long + 256
End If

If Form1.Combo_buzzer.ListIndex = 0 Then
   input_long = input_long + 512
End If
'write status reg, modbus register address=255
output_string = modbus_write(255, input_long)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'************************************************************************************************************
'AN1 PARAMETER
If Form1.Combo_AN1_onoff.ListIndex = 3 Then
   input_long = 9
Else
   input_long = Val(Form1.Combo_AN1_onoff.ListIndex)
End If
output_string = modbus_write(201, input_long)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'AN2 PARAMETER
If Form1.Combo_AN2_onoff.ListIndex = 3 Then
   input_long = 9
Else
   input_long = Val(Form1.Combo_AN2_onoff.ListIndex)
End If
output_string = modbus_write(202, input_long)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'AN1 0% OUTPUT
input_long = analog_scale_calc(Val(Form1.Text_AN1_4ma), Val(Form1.Text_AN1_20ma), Val(Form1.Text_AN1_0deg), Val(Form1.Text_AN1_100deg), Form1.Combo_AN1_I_U.Text, "MIN")
If input_long > 32000 Or input_long < -32000 Then
   output_string = "INVALID AN1 PARAMETER"
Else
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(261, input_long)
End If
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'AN1 100% OUTPUT
input_long = analog_scale_calc(Val(Form1.Text_AN1_4ma), Val(Form1.Text_AN1_20ma), Val(Form1.Text_AN1_0deg), Val(Form1.Text_AN1_100deg), Form1.Combo_AN1_I_U.Text, "MAX")
If input_long > 32000 Or input_long < -32000 Then
   output_string = "INVALID AN1 PARAMETER"
Else
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(262, input_long)
End If
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'AN2 0% OUTPUT
input_long = analog_scale_calc(Val(Form1.Text_AN2_4ma), Val(Form1.Text_AN2_20ma), Val(Form1.Text_AN2_0deg), Val(Form1.Text_AN2_100deg), Form1.Combo_AN2_I_U.Text, "MIN")
If input_long > 32000 Or input_long < -32000 Then
   output_string = "INVALID AN2 PARAMETER"
Else
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(263, input_long)
End If
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'AN2 100% OUTPUT
input_long = analog_scale_calc(Val(Form1.Text_AN2_4ma), Val(Form1.Text_AN2_20ma), Val(Form1.Text_AN2_0deg), Val(Form1.Text_AN2_100deg), Form1.Combo_AN2_I_U.Text, "MAX")
If input_long > 32000 Or input_long < -32000 Then
   output_string = "INVALID AN2 PARAMETER"
Else
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(264, input_long)
End If
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE1 PARAMETER
If Form1.Combo_RE1_onoff.ListIndex = 3 Then
   input_long = 9
Else
   input_long = Val(Form1.Combo_RE1_onoff.ListIndex)
End If
output_string = modbus_write(211, input_long)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE2 PARAMETER
If Form1.Combo_RE2_onoff.ListIndex = 3 Then
   input_long = 9
Else
   input_long = Val(Form1.Combo_RE2_onoff.ListIndex)
End If
output_string = modbus_write(212, input_long)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE1 MODE
input_long = Val(Form1.Combo_RE1_mode.ListIndex)
output_string = modbus_write(219, input_long)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE2 MODE
input_long = Val(Form1.Combo_RE2_mode.ListIndex)
output_string = modbus_write(220, input_long)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE1 LOW
If Form1.Combo_RE1_onoff.ListIndex = 1 Then
   input_single = CDbl(Form1.Text_RE1_L.Text)
   input_long = input_single * 100
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(221, input_long)
ElseIf Form1.Combo_RE1_onoff.ListIndex = 2 Then
   input_long = CDbl(Form1.Text_RE1_L.Text)
   output_string = modbus_write(221, input_long)
End If
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE1 HIGH
If Form1.Combo_RE1_onoff.ListIndex = 1 Then
   input_single = CDbl(Form1.Text_RE1_H.Text)
   input_long = input_single * 100
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(222, input_long)
ElseIf Form1.Combo_RE1_onoff.ListIndex = 2 Then
   input_long = CDbl(Form1.Text_RE1_H.Text)
   output_string = modbus_write(222, input_long)
End If
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE2 LOW
If Form1.Combo_RE2_onoff.ListIndex = 1 Then
   input_single = CDbl(Form1.Text_RE2_L.Text)
   input_long = input_single * 100
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(223, input_long)
ElseIf Form1.Combo_RE2_onoff.ListIndex = 2 Then
   input_long = CDbl(Form1.Text_RE2_L.Text)
   output_string = modbus_write(223, input_long)
End If
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE2 HIGH
If Form1.Combo_RE2_onoff.ListIndex = 1 Then
   input_single = CDbl(Form1.Text_RE2_H.Text)
   input_long = input_single * 100
   input_long = arv_taiendkoodiks(input_long)
   output_string = modbus_write(224, input_long)
ElseIf Form1.Combo_RE2_onoff.ListIndex = 2 Then
   input_long = CDbl(Form1.Text_RE2_H.Text)
   output_string = modbus_write(224, input_long)
End If

If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE1 DELAY
output_string = modbus_write(215, Val(Form1.Text_RE1_delay.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE2 DELAY
output_string = modbus_write(216, Val(Form1.Text_RE1_delay.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE1 ON/OFF TIME
output_string = modbus_write(217, Val(Form1.Text_RE1_time.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RE2 ON/OFF TIME
output_string = modbus_write(218, Val(Form1.Text_RE2_time.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''IGAKS JUHUKS LÕPPU, MCU BUG
'STOP BITS
output_string = modbus_write(7, Val(Form1.Combo_stop_bit.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'BAUD
output_string = modbus_write(5, Val(Form1.Combo_baud.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RESPONSE DELAY
output_string = modbus_write(6, Val(Form1.Text_response.Text))
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If

'RESET'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
output_string = write_no_respond(17, 42330)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If
'RESET uuesti
output_string = write_no_respond(17, 42330)
If output_string <> "WRITE COMPLETE" Then
GoTo writing_error
End If

E222x_write_all = "WRITE COMPLETE"
Exit Function

writing_error:
E222x_write_all = output_string
output_string = write_no_respond(17, 42330) ' igaks juhuks reset
End Function


