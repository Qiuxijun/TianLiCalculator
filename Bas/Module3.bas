Attribute VB_Name = "Module3"
Public Sub SolveCharBonus(ByRef c As Chars)
Dim i%, t As Single, temp As Variant, j As Integer
    i = CBoxFlag + 1 '命座等级
        Select Case CharList(c.cNumber, 1)
            Case "阿贝多"
                If CurrSkill = "c1e2" Then
                j = GetBuffCount("阿贝多天赋2")
                       If j > 0 Then AddBonus c, 25, 3, 1, CharList(c.cNumber, 1) + "天赋2"
                End If
                If InStr(2, CurrSkill, "d") > 0 Then
                j = GetBuffCount("阿贝多命座4")
                       If j > 0 Then AddBonus c, 30, 3, 1, CharList(c.cNumber, 1) + "命座4"
                End If
                
            Case "宵宫"
                j = GetBuffCount("宵宫天赋2")
                If j > 0 Then c.PyroDMG = c.PyroDMG + 2 * j: c.Bonustip = c.Bonustip + "来自宵宫天赋2的火元素伤害加成：" + CStr(2 * j) + "%" + vbCrLf
                j = GetBuffCount("宵宫命座1")
                If j > 0 Then AddBonus c, 20, 1, 1, CharList(c.cNumber, 1) + "命座1"
                j = GetBuffCount("宵宫命座2")
                If j > 0 Then c.PyroDMG = c.PyroDMG + 25: c.Bonustip = c.Bonustip + "来自宵宫命座2的火元素伤害加成：25%" + vbCrLf
                
            Case "芭芭拉"
                j = GetBuffCount("芭芭拉命座2")
                If j > 0 Then c.HydroDMG = c.HydroDMG + 15: c.Bonustip = c.Bonustip + "来自芭芭拉命座2的水元素伤害加成：15%" + vbCrLf
                
                
            Case "魈"
                j = GetBuffCount("魈：处于")
                If j > 0 Then j = (Int((j - 1) / 3) + 1) * 5
                If (CurrSkill = "c2d1" Or CurrSkill = "c2d2" Or CurrSkill = "c2d3" Or CurrSkill = "c2a1" Or CurrSkill = "c2c1") Then
                If j > 0 Then
                    FrmMain.Label2(2).Caption = "风"
                    temp = Array(58.45, 61.95, 65.45, 70, 73.5, 77, 81.55, 86.1, 90.65, 95.2, 99.75, 104.3, 108.85, 113.4, 117.95)
                    c.AnemoDMG = c.AnemoDMG + temp(FrmMain.LevelBox(2).ListIndex - 1) + j
                    c.Bonustip = c.Bonustip + "来自Q技能的增伤：" + CStr(temp(FrmMain.LevelBox(2).ListIndex - 1)) + "%" + vbCrLf
                    c.Bonustip = c.Bonustip + "来自魈天赋2的增伤：" + CStr(j) + "%" + vbCrLf
                Else
                    FrmMain.Label2(2).Caption = "物理"
                End If
                End If
                
            Case "安柏"
                If CurrSkill = "c5q1" And c.Level >= 42 Then
                    c.CritRate = c.CritRate + 10
                    c.CritRatetip = c.CritRatetip + "来自安柏天赋2的暴击率：10%" + vbCrLf
                    If CBoxFlag = 5 Then AddBonus c, 15, 1, 1, "安柏命座6"
                End If
                
            Case "北斗"
                If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then
                    j = GetBuffCount("北斗天赋2")
                    If j > 0 Then AddBonus c, 15, 3, 1, "北斗天赋2"
                End If
                
            Case "胡桃"
                j = GetBuffCount("胡桃：是否")
                temp = Array(3.84, 4.07, 4.3, 4.6, 4.83, 5.06, 5.36, 5.66, 5.96, 6.26, 6.56, 6.85, 7.15, 7.45, 7.75)
                t = c.MaxHP * (1 + c.HPBonus / 100) + c.HPFlag
                t = t * temp(FrmMain.LevelBox(1).ListIndex - 1) / 100
                If j > 0 Then AddBonus c, t, 2, 1, "胡桃元素战技"

                If c.lowHP Then
                    c.PyroDMG = c.PyroDMG + 33
                    c.Bonustip = c.Bonustip + "来自胡桃天赋3的火元素伤害加成：33%" + vbCrLf
                End If
            
            
            Case "迪卢克"
                j = GetBuffCount("迪卢克：是否开启")
                If j > 0 Then
                    c.PyroDMG = c.PyroDMG + 20
                    c.Bonustip = c.Bonustip + "来自迪卢克天赋3的火元素伤害加成：20%" + vbCrLf
                End If
                j = GetBuffCount("迪卢克命座1")
                    If j > 0 Then AddBonus c, 15, 3, 1, "迪卢克命座1"
                j = GetBuffCount("迪卢克命座2")
                    If j > 0 Then AddBonus c, 10 * j, 1, 1, "迪卢克命座2"
                j = GetBuffCount("迪卢克命座4")
                    If j > 0 Then AddBonus c, 40, 3, 1, "迪卢克命座4"
                j = GetBuffCount("迪卢克命座6")
                    If j > 0 Then AddBonus c, 30, 3, 1, "迪卢克命座6"
                
                
            
            Case "神里绫华"
                j = GetBuffCount("神里绫华命座6")
                    If j > 0 And InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 298, 3, 1, "神里绫华命座6"
                j = GetBuffCount("神里绫华：是否使用")
                    If j > 0 Then
                        c.CryoDMG = c.CryoDMG + 18
                        c.Bonustip = c.Bonustip + "来自神里绫华天赋3的冰元素伤害加成：18%" + vbCrLf
                End If
                j = GetBuffCount("神里绫华：天赋2")
                If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 30, 3, 1, "神里绫华天赋2"
            Case "雷电将军"
                        t = (c.Energy - 100) * 0.4
                        c.ElectroDMG = c.ElectroDMG + t
                        c.Bonustip = c.Bonustip + "来自雷电将军天赋2的雷元素伤害加成：" + CStr(t) + "%" + vbCrLf
        End Select
        
        If c.绝缘4 = True Then
                t = Round(c.Energy * 0.25, 2)
                If t > 75 Then t = 75
                Call Jug2(c, t)
                c.Bonustip = c.Bonustip + "来自绝缘4件套的增伤：" + CStr(t) + "%" + vbCrLf
        End If
End Sub
