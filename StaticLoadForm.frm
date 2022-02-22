
'*******************************************************************
'*   Ten phuong thuc: ApplyBtn_Click                               *
'*   Noi dung: Xu li khi click button Apply                        *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub ApplyBtn_Click()
    Call writeCommand
End Sub

'*******************************************************************
'*   Ten phuong thuc: CornForceLwrLeft_Change                      *
'*   Noi dung: Validate textbox CornForceLwrLeft                   *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub CornForceLwrLeft_Change()
    If (Not IsNumeric(Me.CornForceLwrLeft)) And (Me.CornForceLwrLeft.Value <> "") Then
        Me.CornForceLwrLeft.BackColor = 65535
    Else
        Me.CornForceLwrLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: CornForceUprLeft_Change                      *
'*   Noi dung: Validate textbox CornForceUprLeft                   *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub CornForceUprLeft_Change()
    If (Not IsNumeric(Me.CornForceUprLeft)) And (Me.CornForceUprLeft.Value <> "") Then
        Me.CornForceUprLeft.BackColor = 65535
    Else
        Me.CornForceUprLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: CornForceLwrRight_Change                     *
'*   Noi dung: Validate textbox CornForceLwrRight                  *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub CornForceLwrRight_Change()
    If (Not IsNumeric(Me.CornForceLwrRight)) And (Me.CornForceLwrRight <> "") Then
        Me.CornForceLwrRight.BackColor = 65535
    Else
        Me.CornForceLwrRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: CornForceUprRight_Change                     *
'*   Noi dung: Validate textbox CornForceUprRight                  *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub CornForceUprRight_Change()
    If (Not IsNumeric(Me.CornForceUprRight)) And (Me.CornForceUprRight <> "") Then
        Me.CornForceUprRight.BackColor = 65535
    Else
        Me.CornForceUprRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: BraForceLwrLeft_Change                       *
'*   Noi dung: Validate textbox BraForceLwrLeft                    *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub BraForceLwrLeft_Change()
    If (Not IsNumeric(Me.BraForceLwrLeft)) And (Me.BraForceLwrLeft.Value <> "") Then
        Me.BraForceLwrLeft.BackColor = 65535
    Else
        Me.BraForceLwrLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: BraForceUprLeft_Change                       *
'*   Noi dung: Validate textbox BraForceUprLeft                    *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub BraForceUprLeft_Change()
    If (Not IsNumeric(Me.BraForceUprLeft)) And (Me.BraForceUprLeft.Value <> "") Then
        Me.BraForceUprLeft.BackColor = 65535
    Else
        Me.BraForceUprLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: BraForceLwrRight_Change                      *
'*   Noi dung: Validate textbox BraForceLwrRight                   *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub BraForceLwrRight_Change()
    If (Not IsNumeric(Me.BraForceLwrRight)) And (Me.BraForceLwrRight <> "") Then
        Me.BraForceLwrRight.BackColor = 65535
    Else
        Me.BraForceLwrRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: BraForceUprRight_Change                      *
'*   Noi dung: Validate textbox BraForceUprRight                   *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub BraForceUprRight_Change()
    If (Not IsNumeric(Me.BraForceUprRight)) And (Me.BraForceUprRight <> "") Then
        Me.BraForceUprRight.BackColor = 65535
    Else
        Me.BraForceUprRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: TracForceLwrLeft_Change                      *
'*   Noi dung: Validate textbox TracForceLwrLeft                   *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub TracForceLwrLeft_Change()
    If (Not IsNumeric(Me.TracForceLwrLeft)) And (Me.TracForceLwrLeft.Value <> "") Then
        Me.TracForceLwrLeft.BackColor = 65535
    Else
        Me.TracForceLwrLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: TracForceUprLeft_Change                      *
'*   Noi dung: Validate textbox TracForceUprLeft                   *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub TracForceUprLeft_Change()
    If (Not IsNumeric(Me.TracForceUprLeft)) And (Me.TracForceUprLeft.Value <> "") Then
        Me.TracForceUprLeft.BackColor = 65535
    Else
        Me.TracForceUprLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: TracForceLwrRight_Change                     *
'*   Noi dung: Validate textbox TracForceLwrRight                  *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub TracForceLwrRight_Change()
    If (Not IsNumeric(Me.TracForceLwrRight)) And (Me.TracForceLwrRight <> "") Then
        Me.TracForceLwrRight.BackColor = 65535
    Else
        Me.TracForceLwrRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: TracForceUprRight_Change                     *
'*   Noi dung: Validate textbox TracForceUprRight                  *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub TracForceUprRight_Change()
    If (Not IsNumeric(Me.TracForceUprRight)) And (Me.TracForceUprRight <> "") Then
        Me.TracForceUprRight.BackColor = 65535
    Else
        Me.TracForceUprRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: VerLenLwrLeft_Change                         *
'*   Noi dung: Validate textbox VerLenLwrLeft                      *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub VerLenLwrLeft_Change()
    If (Not IsNumeric(Me.VerLenLwrLeft)) And (Me.VerLenLwrLeft.Value <> "") Then
        Me.VerLenLwrLeft.BackColor = 65535
    Else
        Me.VerLenLwrLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: VerLenUprLeft_Change                         *
'*   Noi dung: Validate textbox VerLenUprLeft                      *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub VerLenUprLeft_Change()
    If (Not IsNumeric(Me.VerLenUprLeft)) And (Me.VerLenUprLeft.Value <> "") Then
        Me.VerLenUprLeft.BackColor = 65535
    Else
        Me.VerLenUprLeft.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: VerLenLwrRight_Change                        *
'*   Noi dung: Validate textbox VerLenLwrRight                     *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub VerLenLwrRight_Change()
    If (Not IsNumeric(Me.VerLenLwrRight)) And (Me.VerLenLwrRight <> "") Then
        Me.VerLenLwrRight.BackColor = 65535
    Else
        Me.VerLenLwrRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: VerLenUprRight_Change                        *
'*   Noi dung: Validate textbox VerLenUprRight                     *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub VerLenUprRight_Change()
    If (Not IsNumeric(Me.VerLenUprRight)) And (Me.VerLenUprRight <> "") Then
        Me.VerLenUprRight.BackColor = 65535
    Else
        Me.VerLenUprRight.BackColor = -2147483648#
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: OKBtn_Click                                  *
'*   Noi dung: Xu li khi click button OK                           *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub OKBtn_Click()
    Call writeCommand
End Sub

'*******************************************************************
'*   Ten phuong thuc: OutPrefix_Change                             *
'*   Noi dung: Validate textbox OutPrefix                          *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub OutPrefix_Change()
    If IsNumeric(Left(Me.OutPrefix, 1)) Then
        Me.OutPrefix.BackColor = 65535
    Else
        Me.OutPrefix.BackColor = -2147483643
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: UserForm_Initialize (phuong thuc khoi tao)   *
'*   Noi dung: Khoi tao gia tri cho cac combobox, radiobox         *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub UserForm_Initialize()

    'Tao list cho SimulationMode
    Me.SimulationMode.AddItem Constant.SIMULATE_MODE_INTERACTIVE
    Me.SimulationMode.AddItem Constant.SIMULATE_MODE_GRAPHICAL
    Me.SimulationMode.AddItem Constant.SIMULATE_MODE_BACKGROUND
    Me.SimulationMode.AddItem Constant.SIMULATE_MODE_FILE_ONLY
    Me.SimulationMode.AddItem Constant.SIMULATE_MODE_EVENT_ONLY
    
    'Tao list cho VerticalMode
    Me.VerticalMode.AddItem Constant.VERTICAL_MODE_WHEEL_CENTER
    Me.VerticalMode.AddItem Constant.VERTICAL_MODE_CONTACT_PATCH
    
    'Tao list cho VerticalInput
    Me.VerticalInput.AddItem Constant.VERTICAL_INPUT_CON_PAT_HEI
    Me.VerticalInput.AddItem Constant.VERTICAL_INPUT_WHE_CEN_HEI
    Me.VerticalInput.AddItem Constant.VERTICAL_INPUT_WHE_VER_FOR
    
    'Check absolute
    Me.Absolute = True
End Sub

'*******************************************************************
'*   Ten phuong thuc: CancelBtn_Click                              *
'*   Noi dung: Tat form khi click button cancel                    *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub CancelBtn_Click()
    Unload Me
End Sub

'*******************************************************************
'*   Ten phuong thuc: StepNum_Change                               *
'*   Noi dung: Validate textbox StepNum                            *
'*               Neu co loi thi hien thi nen vang                  *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub StepNum_Change()
    If Not IsNumeric(Me.StepNum) Then
        Me.StepNum.BackColor = 65535
    Else
        Me.StepNum.BackColor = -2147483643
    End If
End Sub

'*******************************************************************
'*   Ten phuong thuc: VerticalInput_Change                         *
'*   Noi dung: Enable/Disable nut chon absolute/relative           *
'*               theo combobox VerticalInput                       *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub VerticalInput_Change()
    If Me.VerticalInput.Value = Constant.VERTICAL_INPUT_WHE_VER_FOR Then
        Me.Absolute.Enabled = False
        Me.Relative.Enabled = False
    Else
        Me.Absolute.Enabled = True
        Me.Relative.Enabled = True
    End If
End Sub


'*******************************************************************
'*   Ten phuong thuc: generateCommand                              *
'*   Noi dung: Tao command ADAMS theo noi dung nhap tren form      *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Function generateCommand() As String
    Dim result As String                                'Gia tri output
    Dim verInputSelection As String                     'Lay gia tri VerticalInput ghi vao command
    Dim absoluteRelativeSelection As String             'Lay gia tri Absolute/Relative ghi vao command
    
    result = "if condition=(!db_exists("".EVENT_SETS.default.evs_" & Me.OutPrefix & "_static_load""))" & vbCrLf
    result = result + "    acar analysis event create &" & vbCrLf
    result = result + "     name=""" & Me.OutPrefix & "_static_load"" &" & vbCrLf
    result = result + "     class=.ACAR.macros.mac_ana_sus_sta_sub &" & vbCrLf
    result = result + "     model=" & Me.AssyName & " &" & vbCrLf
    result = result + "     variant='default'" & vbCrLf
    result = result + "end" & vbCrLf
    
    '----------------------------------------
    result = result + "acar analysis event modify &" & vbCrLf
    result = result + " event=.EVENT_SETS.default.evs_" & Me.OutPrefix & "_static_load &" & vbCrLf
    result = result + " class=.ACAR.macros.mac_ana_sus_sta_sub &" & vbCrLf
    result = result + " model=" & Me.AssyName & " &" & vbCrLf
    result = result + " variant='default' &" & vbCrLf
    result = result + " attributes=""variant"", &" & vbCrLf
    result = result + "     ""analysis_mode"", &" & vbCrLf
    result = result + "     ""simulationSteps"", &" & vbCrLf
    result = result + "     ""stat_steer_pos"", &" & vbCrLf
    result = result + "     ""steering_input"", &" & vbCrLf
    result = result + "     ""loadingMethod"", &" & vbCrLf
    result = result + "     ""vertical_input"", &" & vbCrLf
    result = result + "     ""vertical_type"", &" & vbCrLf
    result = result + "     ""coordinate_system"", &" & vbCrLf
    result = result + "     ""log_file"", &" & vbCrLf
    result = result + "     ""output_prefix"" &" & vbCrLf
    result = result + "  values=""default"", &" & vbCrLf
    result = result + "     """ & Me.SimulationMode & """, &" & vbCrLf
    result = result + "     """ & Me.StepNum & """, &" & vbCrLf
    result = result + "     """", &" & vbCrLf
    result = result + "     ""angle"", &" & vbCrLf
    result = result + "     ""wheel_center_height"", &" & vbCrLf
    result = result + "     """", &" & vbCrLf
    result = result + "     """", &" & vbCrLf
    result = result + "     ""vehicle"", &" & vbCrLf
    result = result + "     ""True"", &" & vbCrLf
    result = result + "     """ & Me.OutPrefix & """" & vbCrLf
    
    '----------------------------------------
    result = result + "acar analysis event modify event=.EVENT_SETS.default.evs_" & Me.OutPrefix & "_static_load &" & vbCrLf
    result = result + " attributes=""vertical_sta_input"", ""vertical_sta_type"", &" & vbCrLf
    result = result + "            ""align_tor_lwr_l"", ""align_tor_upr_l"", ""align_tor_lwr_r"",""align_tor_upr_r"", &" & vbCrLf
    result = result + "            ""later_for_lwr_l"", ""later_for_upr_l"", ""later_for_lwr_r"",""later_for_upr_r"", &" & vbCrLf
    result = result + "            ""brake_for_lwr_l"", ""brake_for_upr_l"", ""brake_for_lwr_r"",""brake_for_upr_r"", &" & vbCrLf
    result = result + "            ""drivn_for_lwr_l"", ""drivn_for_upr_l"", ""drivn_for_lwr_r"",""drivn_for_upr_r"", &" & vbCrLf
    result = result + "            ""verti_for_lwr_l"", ""verti_for_upr_l"", ""verti_for_lwr_r"",""verti_for_upr_r"", &" & vbCrLf
    result = result + "            ""otm_lwr_l"", ""otm_upr_l"", ""otm_lwr_r"",""otm_upr_r"", &" & vbCrLf
    result = result + "            ""roll_res_tor_lwr_l"", ""roll_res_tor_upr_l"", ""roll_res_tor_lwr_r"",""roll_res_tor_upr_r"", &" & vbCrLf
    result = result + "            ""damage_for_lwr_l"", ""damage_for_upr_l"", ""damage_for_lwr_r"",""damage_for_upr_r"", &" & vbCrLf
    result = result + "            ""damage_rad_l"", ""damage_rad_r"", ""steer_lower"",""steer_upper"" &" & vbCrLf
    
    If Me.VerticalInput.Value = Constant.VERTICAL_INPUT_CON_PAT_HEI Then
        verInputSelection = "contact_patch_height"
    ElseIf (Me.VerticalInput.Value = Constant.VERTICAL_INPUT_WHE_CEN_HEI) Then
        verInputSelection = "wheel_center_height"
    ElseIf (Me.VerticalInput.Value = Constant.VERTICAL_INPUT_WHE_CEN_HEI) Then
        verInputSelection = "wheel_vertical_force"
    End If
    
    If Me.Absolute.Enabled = False Then
        absoluteRelativeSelection = ""
    ElseIf Me.Absolute.Value = True Then
        absoluteRelativeSelection = "absolute"
    Else
        absoluteRelativeSelection = "relative"
    End If
    result = result + " value=""" & verInputSelection & """, """ & absoluteRelativeSelection & """, &" & vbCrLf
    result = result + "            """ & Me.CornForceLwrLeft.Value & """, """ & Me.CornForceUprLeft.Value & """, """ & Me.CornForceLwrRight.Value & """, """ & Me.CornForceUprRight.Value & """, &" & vbCrLf
    result = result + "            """ & Me.BraForceLwrLeft.Value & """, """ & Me.BraForceUprLeft.Value & """, """ & Me.BraForceLwrRight.Value & """, """ & Me.BraForceUprRight.Value & """, &" & vbCrLf
    result = result + "            """ & Me.TracForceLwrLeft.Value & """, """ & Me.TracForceUprLeft.Value & """, """ & Me.TracForceLwrRight.Value & """, """ & Me.TracForceUprRight.Value & """, &" & vbCrLf
    result = result + "            """ & Me.VerLenLwrLeft.Value & """, """ & Me.VerLenUprLeft.Value & """, """ & Me.VerLenLwrRight.Value & """, """ & Me.VerLenUprRight.Value & """, &" & vbCrLf
    result = result + "            """", """", """", """", &" & vbCrLf
    result = result + "            """", """", """", """", &" & vbCrLf
    result = result + "            """", """", """", """", &" & vbCrLf
    result = result + "            """", """", """", """", &" & vbCrLf
    result = result + "            """", """", """", """"" & vbCrLf
    result = result + "acar analysis instance event execute instance_name=.EVENT_SETS.default.evs_" & Me.OutPrefix & "_static_load analysis_mode=" & Me.SimulationMode
    
    '----------------------------------------
    generateCommand = result
    
End Function

'*******************************************************************
'*   Ten phuong thuc: writeCommand                                 *
'*   Noi dung: Ghi command vao sheet Command                       *
'*             Ghi link file output command vao sheet Command      *
'*   Author: LVN-HoaPV                                             *
'*   Date: 2022/01/10                                              *
'*   Version: 1.0  | Tao moi                                       *
'*******************************************************************
Private Sub writeCommand()
    Dim cmdSheet As Worksheet                           'Doi tuong chua sheet Command
    Dim obj As Object                                   'Object lay cac object con cua form nhap de validate
    
    On Error GoTo writeCommandError
    
    'Validate cac object trong form, neu co object nen vang thi bao loi
    For Each obj In Me.Controls
        If obj.BackColor = 65535 Then
            MsgBox "[ERROR] Hay nhap gia tri hop le", vbCritical
            Exit Sub
        End If
    Next
    
    'Tim o cuoi cung con trong trong sheet Command
    Set cmdSheet = ThisWorkbook.Worksheets("Command")
    cmdSheet.Cells(2, "C").Value = ActiveWorkbook.Path & "\static_load_command.cmd"
    For i = 6 To 200
        If cmdSheet.Cells(i, "C") = "" Then
            Exit For
        End If
    Next
    'Ghi command, tat form
    cmdSheet.Cells(i, "C").Value = generateCommand
    Unload Me
    Exit Sub
    
writeCommandError:
    Call Common.catchError("Khong tim thay sheet command")
End Sub
