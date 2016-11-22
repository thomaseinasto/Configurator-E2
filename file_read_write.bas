Attribute VB_Name = "file_read_write"
Private port0 As String
Private baud0 As String
Private stop_bit0 As String
Private slave0 As String
'*******************************************************************************************************
Private hardware As String
Private serial As String
Private modbus_delay As String
Private slave As String
Private baud As String
Private stop_bit As String
Private zero_T As String
Private zero_RH As String
Private slope_RH As String
Private rate_RH As String
Private global_analog As String
Private global_filter As String
Private factory_key As String
'ANALOG1***********************************************************************************************
Private AN1_onoff As String
Private AN1_I_U As String
Private AN1_4ma As String
Private AN1_20ma As String
Private AN1_0deg As String
Private AN1_100deg As String
Private AN1_diag As String
'ANALOG2*********************************************************************************************
Private AN2_onoff As String
Private AN2_I_U As String
Private AN2_4ma As String
Private AN2_20ma As String
Private AN2_0deg As String
Private AN2_100deg As String
Private AN2_diag As String
'RELAY1**********************************************************************************************
Private RE1_onoff As String
Private RE1_mode As String
Private RE1_H As String
Private RE1_L As String
Private RE1_delay As String
Private RE1_time As String
'RELAY2***********************************************************************************************
Private RE2_onoff As String
Private RE2_mode As String
Private RE2_H As String
Private RE2_L As String
Private RE2_delay As String
Private RE2_time As String
'*****************************************************************************************************
'CUSTOM
'*****************************************************************************************************
'custom1
Private Check1 As Byte
Private Saved As Byte
Private Check9 As Byte
Private Check10 As Byte
Private Text_name1 As String
Private Text_c_adr1 As String
Private Text_c_write1 As String
Private Check_c_neg1 As Byte
'custom2
Private Check2 As Byte
Private Text_name2 As String
Private Text_c_adr2 As String
Private Text_c_write2 As String
Private Check_c_neg2 As Byte
'custom3
Private Check3 As Byte
Private Text_name3 As String
Private Text_c_adr3 As String
Private Text_c_write3 As String
Private Check_c_neg3 As Byte
'custom4
Private Check4 As Byte
Private Text_name4 As String
Private Text_c_adr4 As String
Private Text_c_write4 As String
Private Check_c_neg4 As Byte
'custom5
Private Check5 As Byte
Private Text_name5 As String
Private Text_c_adr5 As String
Private Text_c_write5 As String
Private Check_c_neg5 As Byte
'custom6
Private Check6 As Byte
Private Text_name6 As String
Private Text_c_adr6 As String
Private Text_c_write6 As String
Private Check_c_neg6 As Byte
'custom7
Private Check7 As Byte
Private Text_name7 As String
Private Text_c_adr7 As String
Private Text_c_write7 As String
Private Check_c_neg7 As Byte
'custom8
Private Check8 As Byte
Private Text_name8 As String
Private Text_c_adr8 As String
Private Text_c_write8 As String
Private Check_c_neg8 As Byte
'GAASIDETEKTORI LISAD******************************************************************************
Private heater_pulse As String
Private sensor_pulse As String
Private const_B As String
Private const_C As String
Private const_D As String
Private const_E As String
Private buzzer As String
Private valgusdiood As String
Private gas_units As String
Private gas_type As String

Public Sub read_from_txt()

On Error GoTo file_not_found
Dim strEmpFileName  As String
    Dim strBackSlash  As String
    Dim intEmpFileNbr As Integer
    
  
    strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
    strEmpFileName = App.Path & strBackSlash & "E26XX.DAT"
    intEmpFileNbr = FreeFile
    
    Open strEmpFileName For Input As #intEmpFileNbr
    
    Input #intEmpFileNbr, port0, baud0, stop_bit0, slave0, hardware, serial, modbus_delay, slave, baud, stop_bit, zero_T, zero_RH, slope_RH, rate_RH, global_analog, global_filter, factory_key, AN1_onoff, AN1_I_U, AN1_4ma, AN1_20ma, AN1_0deg, AN1_100deg, AN1_diag, AN2_onoff, AN2_I_U, AN2_4ma, AN2_20ma, AN2_0deg, AN2_100deg, AN2_diag, RE1_onoff, RE1_mode, RE1_H, RE1_L, RE1_delay, RE1_time, RE2_onoff, RE2_mode, RE2_H, RE2_L, RE2_delay, RE2_time, Check1, Text_name1, Text_c_adr1, Text_c_write1, Check_c_neg1, Check2, Text_name2, Text_c_adr2, Text_c_write2, Check_c_neg2, Check3, Text_name3, Text_c_adr3, Text_c_write3, Check_c_neg3, Check4, Text_name4, Text_c_adr4, Text_c_write4, Check_c_neg4, Check5, Text_name5, Text_c_adr5, Text_c_write5, Check_c_neg5, Check6, Text_name6, Text_c_adr6, Text_c_write6, Check_c_neg6, Check7, Text_name7, Text_c_adr7, Text_c_write7, Check_c_neg7, Check8, Text_name8, Text_c_adr8, Text_c_write8, Check_c_neg8, _
                          heater_pulse, sensor_pulse, const_B, const_C, const_D, const_E, buzzer, valgusdiood, gas_type, gas_units
       
    Close #intEmpFileNbr
    
    Form1.Combo_port_nr.ListIndex = port0
    Form1.Combo_baud0.ListIndex = baud0
    Form1.Combo_stop_bit0.ListIndex = stop_bit0
    Form1.Text_slave_id0.Text = slave0
    
    Form1.Text_hardware.Text = hardware
    Form1.Text_SN.Text = serial
    Form1.Text_response.Text = modbus_delay
    Form1.Text_slave_id.Text = slave
    
    Form1.Combo_baud.ListIndex = baud
    Form1.Combo_stop_bit.ListIndex = stop_bit
    Form1.Combo_global_AN.ListIndex = global_analog
    
    Form1.Text_zero_T.Text = zero_T
    Form1.Text_zero_RH.Text = zero_RH
    Form1.Text_RH_slope.Text = slope_RH
    Form1.Text_RH_rate.Text = rate_RH
    
    Form1.Text_RC_filter.Text = global_filter
    Form1.Text_factory.Text = factory_key
    
    Form1.Combo_AN1_onoff.ListIndex = AN1_onoff
    Form1.Combo_AN1_I_U.ListIndex = AN1_I_U
    Form1.Text_AN1_4ma.Text = AN1_4ma
    Form1.Text_AN1_20ma.Text = AN1_20ma
    Form1.Text_AN1_0deg.Text = AN1_0deg
    Form1.Text_AN1_100deg.Text = AN1_100deg
    Form1.Combo_AN1_diag.ListIndex = AN1_diag
    
    Form1.Combo_AN2_onoff.ListIndex = AN2_onoff
    Form1.Combo_AN2_I_U.ListIndex = AN2_I_U
    Form1.Text_AN2_4ma.Text = AN2_4ma
    Form1.Text_AN2_20ma.Text = AN2_20ma
    Form1.Text_AN2_0deg.Text = AN2_0deg
    Form1.Text_AN2_100deg.Text = AN2_100deg
    Form1.Combo_AN2_diag.ListIndex = AN2_diag
    
    Form1.Combo_RE1_onoff.ListIndex = RE1_onoff
    Form1.Combo_RE1_mode.ListIndex = RE1_mode
    Form1.Text_RE1_L.Text = RE1_L
    Form1.Text_RE1_H.Text = RE1_H
    Form1.Text_RE1_delay.Text = RE1_delay
    Form1.Text_RE1_time.Text = RE1_time
    
    Form1.Combo_RE2_onoff.ListIndex = RE2_onoff
    Form1.Combo_RE2_mode.ListIndex = RE2_mode
    Form1.Text_RE2_L.Text = RE2_L
    Form1.Text_RE2_H.Text = RE2_H
    Form1.Text_RE2_delay.Text = RE2_delay
    Form1.Text_RE2_time.Text = RE2_time
        
'CUSTOM*********************************************************************************************
    Form1.Check1.Value = Check1
    
    
    Form1.Text_name1.Text = Text_name1
    Form1.Text_c_adr1.Text = Text_c_adr1
    Form1.Text_c_write1.Text = Text_c_write1
    Form1.Check_c_neg1.Value = Check_c_neg1
    
    Form1.Check2.Value = Check2
    Form1.Text_name2.Text = Text_name2
    Form1.Text_c_adr2.Text = Text_c_adr2
    Form1.Text_c_write2.Text = Text_c_write2
    Form1.Check_c_neg2.Value = Check_c_neg2
    
    Form1.Check3.Value = Check3
    Form1.Text_name3.Text = Text_name3
    Form1.Text_c_adr3.Text = Text_c_adr3
    Form1.Text_c_write3.Text = Text_c_write3
    Form1.Check_c_neg3.Value = Check_c_neg3
    
    Form1.Check4.Value = Check4
    Form1.Text_name4.Text = Text_name4
    Form1.Text_c_adr4.Text = Text_c_adr4
    Form1.Text_c_write4.Text = Text_c_write4
    Form1.Check_c_neg4.Value = Check_c_neg4
    
    Form1.Check5.Value = Check5
    Form1.Text_name5.Text = Text_name5
    Form1.Text_c_adr5.Text = Text_c_adr5
    Form1.Text_c_write5.Text = Text_c_write5
    Form1.Check_c_neg5.Value = Check_c_neg5
    
    Form1.Check6.Value = Check6
    Form1.Text_name6.Text = Text_name6
    Form1.Text_c_adr6.Text = Text_c_adr6
    Form1.Text_c_write6.Text = Text_c_write6
    Form1.Check_c_neg6.Value = Check_c_neg6
    
    Form1.Check7.Value = Check7
    Form1.Text_name7.Text = Text_name7
    Form1.Text_c_adr7.Text = Text_c_adr7
    Form1.Text_c_write7.Text = Text_c_write7
    Form1.Check_c_neg7.Value = Check_c_neg7
    
    Form1.Check8.Value = Check8
    Form1.Text_name8.Text = Text_name8
    Form1.Text_c_adr8.Text = Text_c_adr8
    Form1.Text_c_write8.Text = Text_c_write8
    Form1.Check_c_neg8.Value = Check_c_neg8
    'GAASIDETEKTOR:
    Form1.Text_heater_pulse.Text = heater_pulse
    Form1.Text_sensor_pulse.Text = sensor_pulse
    Form1.Text_const_B.Text = const_B
    Form1.Text_const_C.Text = const_C
    Form1.Text_const_D.Text = const_D
    Form1.Text_const_E.Text = const_E
    Form1.Combo_buzzer.ListIndex = buzzer
    Form1.Combo_LED.ListIndex = valgusdiood
    Form1.Combo_gas_type.ListIndex = gas_type
    Form1.Combo_gas_units.ListIndex = gas_units
    
Exit Sub

file_not_found:
MsgBox Err.Description
    
End Sub
Public Sub write_to_txt()

Dim strEmpFileName As String
Dim strBackSlash   As String
Dim intEmpFileNbr  As Integer
    
 port0 = Form1.Combo_port_nr.ListIndex
 baud0 = Form1.Combo_baud0.ListIndex
 stop_bit0 = Form1.Combo_stop_bit0.ListIndex
 slave0 = Form1.Text_slave_id0.Text
    
 hardware = Form1.Text_hardware.Text
 serial = Form1.Text_SN.Text
 modbus_delay = Form1.Text_response.Text
 slave = Form1.Text_slave_id.Text
    
 baud = Form1.Combo_baud.ListIndex
 stop_bit = Form1.Combo_stop_bit.ListIndex
 global_analog = Form1.Combo_global_AN.ListIndex
    
 zero_T = Form1.Text_zero_T.Text
 zero_RH = Form1.Text_zero_RH.Text
 slope_RH = Form1.Text_RH_slope.Text
 rate_RH = Form1.Text_RH_rate.Text
    
 global_filter = Form1.Text_RC_filter.Text
 factory_key = Form1.Text_factory.Text

            
 AN1_onoff = Form1.Combo_AN1_onoff.ListIndex
 AN1_I_U = Form1.Combo_AN1_I_U.ListIndex
 AN1_4ma = Form1.Text_AN1_4ma.Text
 AN1_20ma = Form1.Text_AN1_20ma.Text
 AN1_0deg = Form1.Text_AN1_0deg.Text
 AN1_100deg = Form1.Text_AN1_100deg.Text
 AN1_diag = Form1.Combo_AN1_diag.ListIndex
            
 AN2_onoff = Form1.Combo_AN2_onoff.ListIndex
 AN2_I_U = Form1.Combo_AN2_I_U.ListIndex
 AN2_4ma = Form1.Text_AN2_4ma.Text
 AN2_20ma = Form1.Text_AN2_20ma.Text
 AN2_0deg = Form1.Text_AN2_0deg.Text
 AN2_100deg = Form1.Text_AN2_100deg.Text
 AN2_diag = Form1.Combo_AN2_diag.ListIndex
 
 RE1_onoff = Form1.Combo_RE1_onoff.ListIndex
 RE1_mode = Form1.Combo_RE1_mode.ListIndex
 RE1_L = Form1.Text_RE1_L.Text
 RE1_H = Form1.Text_RE1_H.Text
 RE1_delay = Form1.Text_RE1_delay.Text
 RE1_time = Form1.Text_RE1_time.Text
    
 RE2_onoff = Form1.Combo_RE2_onoff.ListIndex
 RE2_mode = Form1.Combo_RE2_mode.ListIndex
 RE2_L = Form1.Text_RE2_L.Text
 RE2_H = Form1.Text_RE2_H.Text
 RE2_delay = Form1.Text_RE2_delay.Text
 RE2_time = Form1.Text_RE2_time.Text
 
 'CUSTOM********************************************************************************************
  Check1 = Form1.Check1.Value
  'Saved = Form1.Saved.Value
  'Check9 = Form1.Check9.Value
  'Check10 = Form1.Check10.Value
 Text_name1 = Form1.Text_name1.Text
 Text_c_adr1 = Form1.Text_c_adr1.Text
 Text_c_write1 = Form1.Text_c_write1.Text
 Check_c_neg1 = Form1.Check_c_neg1.Value
    
 Check2 = Form1.Check2.Value
 Text_name2 = Form1.Text_name2.Text
 Text_c_adr2 = Form1.Text_c_adr2.Text
 Text_c_write2 = Form1.Text_c_write2.Text
 Check_c_neg2 = Form1.Check_c_neg2.Value
    
 Check3 = Form1.Check3.Value
 Text_name3 = Form1.Text_name3.Text
 Text_c_adr3 = Form1.Text_c_adr3.Text
 Text_c_write3 = Form1.Text_c_write3.Text
 Check_c_neg3 = Form1.Check_c_neg3.Value
    
 Check4 = Form1.Check4.Value
 Text_name4 = Form1.Text_name4.Text
 Text_c_adr4 = Form1.Text_c_adr4.Text
 Text_c_write4 = Form1.Text_c_write4.Text
 Check_c_neg4 = Form1.Check_c_neg4.Value
    
 Check5 = Form1.Check5.Value
 Text_name5 = Form1.Text_name5.Text
 Text_c_adr5 = Form1.Text_c_adr5.Text
 Text_c_write5 = Form1.Text_c_write5.Text
 Check_c_neg5 = Form1.Check_c_neg5.Value
    
 Check6 = Form1.Check6.Value
 Text_name6 = Form1.Text_name6.Text
 Text_c_adr6 = Form1.Text_c_adr6.Text
 Text_c_write6 = Form1.Text_c_write6.Text
 Check_c_neg6 = Form1.Check_c_neg6.Value
    
 Check7 = Form1.Check7.Value
 Text_name7 = Form1.Text_name7.Text
 Text_c_adr7 = Form1.Text_c_adr7.Text
 Text_c_write7 = Form1.Text_c_write7.Text
 Check_c_neg7 = Form1.Check_c_neg7.Value
    
 Check8 = Form1.Check8.Value
 Text_name8 = Form1.Text_name8.Text
 Text_c_adr8 = Form1.Text_c_adr8.Text
 Text_c_write8 = Form1.Text_c_write8.Text
 Check_c_neg8 = Form1.Check_c_neg8.Value
 'GAASIDETEKTORI ASJAD:
 heater_pulse = Form1.Text_heater_pulse.Text
 sensor_pulse = Form1.Text_sensor_pulse.Text
 const_B = Form1.Text_const_B.Text
 const_C = Form1.Text_const_C.Text
 const_D = Form1.Text_const_D.Text
 const_E = Form1.Text_const_E.Text
 buzzer = Form1.Combo_buzzer.ListIndex
 valgusdiood = Form1.Combo_LED.ListIndex
 gas_type = Form1.Combo_gas_type.ListIndex
 gas_units = Form1.Combo_gas_units.ListIndex



 strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
    strEmpFileName = App.Path & strBackSlash & "E26XX.DAT"
    intEmpFileNbr = FreeFile
    
    Open strEmpFileName For Output As #intEmpFileNbr
 
    Write #intEmpFileNbr, port0, baud0, stop_bit0, slave0, hardware, serial, modbus_delay, slave, baud, stop_bit, zero_T, zero_RH, slope_RH, rate_RH, global_analog, global_filter, factory_key, AN1_onoff, AN1_I_U, AN1_4ma, AN1_20ma, AN1_0deg, AN1_100deg, AN1_diag, AN2_onoff, AN2_I_U, AN2_4ma, AN2_20ma, AN2_0deg, AN2_100deg, AN2_diag, RE1_onoff, RE1_mode, RE1_H, RE1_L, RE1_delay, RE1_time, RE2_onoff, RE2_mode, RE2_H, RE2_L, RE2_delay, RE2_time, Check1, Text_name1, Text_c_adr1, Text_c_write1, Check_c_neg1, Check2, Text_name2, Text_c_adr2, Text_c_write2, Check_c_neg2, Check3, Text_name3, Text_c_adr3, Text_c_write3, Check_c_neg3, Check4, Text_name4, Text_c_adr4, Text_c_write4, Check_c_neg4, Check5, Text_name5, Text_c_adr5, Text_c_write5, Check_c_neg5, Check6, Text_name6, Text_c_adr6, Text_c_write6, Check_c_neg6, Check7, Text_name7, Text_c_adr7, Text_c_write7, Check_c_neg7, Check8, Text_name8, Text_c_adr8, Text_c_write8, Check_c_neg8, _
                          heater_pulse, sensor_pulse, const_B, const_C, const_D, const_E, buzzer, valgusdiood, gas_type, gas_units
 Close #intEmpFileNbr
 
 
End Sub

