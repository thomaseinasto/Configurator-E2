Attribute VB_Name = "Custom"
Option Explicit
Public Function Custom_read() As String
Dim output_string As String
Dim output_long As Long
Dim error_flag As String
error_flag = "READ SUCCESSFUL"

'custom1
If Form1.Check1.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr1.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg1.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read1.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

'custom2
If Form1.Check2.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr2.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg2.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read2.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

'custom3
If Form1.Check3.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr3.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg3.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read3.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

'custom4
If Form1.Check4.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr4.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg4.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read4.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

'custom5
If Form1.Check5.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr5.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg5.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read5.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

'custom6
If Form1.Check6.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr6.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg6.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read6.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

'custom7
If Form1.Check7.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr7.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg7.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read7.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

'custom8
If Form1.Check8.Value = 1 Then
   output_string = modbus_read(Val(Form1.Text_c_adr8.Text))
   If IsNumeric(output_string) Then
      output_long = Val(output_string)
      If Form1.Check_c_neg8.Value = 1 Then
         output_long = taiendkood_arvuks(output_long)
      End If
      Form1.Label_c_read8.Caption = output_long
   Else
      error_flag = output_string
   End If
End If

Custom_read = error_flag

End Function
Public Function Custom_write() As String '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''RHT WRITING
Dim output_string As String
Dim input_long As Long
Dim error_flag As String
error_flag = "WRITE COMPLETE"

'custom1
If Form1.Check1.Value = 1 Then
   input_long = Val(Form1.Text_c_write1.Text)
   If Form1.Check_c_neg1.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr1.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

'custom2
If Form1.Check2.Value = 1 Then
   input_long = Val(Form1.Text_c_write2.Text)
   If Form1.Check_c_neg2.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr2.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

'custom3
If Form1.Check3.Value = 1 Then
   input_long = Val(Form1.Text_c_write3.Text)
   If Form1.Check_c_neg3.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr3.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

'custom4
If Form1.Check4.Value = 1 Then
   input_long = Val(Form1.Text_c_write4.Text)
   If Form1.Check_c_neg4.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr4.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

'custom5
If Form1.Check5.Value = 1 Then
   input_long = Val(Form1.Text_c_write5.Text)
   If Form1.Check_c_neg5.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr5.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

'custom6
If Form1.Check6.Value = 1 Then
   input_long = Val(Form1.Text_c_write6.Text)
   If Form1.Check_c_neg6.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr6.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

'custom7
If Form1.Check7.Value = 1 Then
   input_long = Val(Form1.Text_c_write7.Text)
   If Form1.Check_c_neg7.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr7.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

'custom8
If Form1.Check8.Value = 1 Then
   input_long = Val(Form1.Text_c_write8.Text)
   If Form1.Check_c_neg8.Value = 1 Then
      input_long = arv_taiendkoodiks(input_long)
   End If
   output_string = modbus_write(Val(Form1.Text_c_adr8.Text), input_long)
   If output_string <> "WRITE COMPLETE" Then
      error_flag = output_string
   End If
End If

Custom_write = error_flag


End Function
