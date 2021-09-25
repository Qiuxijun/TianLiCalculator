Attribute VB_Name = "Module2"
Public Sub SolveBonus(ByRef c As Chars)
Dim i%, t As Single, j%, t2 As Single, flag As Boolean
    i = RBoxFlag + 1 '精炼等级
    j = 0
    Select Case WeaponList(c.cWeapon, 1)
        Case "磐岩结绿"
            c.HPBonus = 15 + i * 5
            c.HPtip = c.HPtip + "来自磐岩结绿的生命值：" + CStr(15 + i * 5) + "%" + vbCrLf
            t = c.MaxHP * (1 + c.HPBonus / 100) + c.HPFlag
            t = Round(t * (0.9 + i * 0.3) / 100, 2)
            c.ATKFlag = c.ATKFlag + t
            c.ATKtip = c.ATKtip + "来自磐岩结绿的攻击力：" + CStr(t) + vbCrLf
        Case "和璞鸢"
            j = GetBuffCount("和璞鸢")
            If j > 0 Then
                c.ATKBonus = c.ATKBonus + (2.5 + i * 0.7) * j
                c.ATKtip = c.ATKtip + "来自和璞鸢的攻击力：" + CStr((2.5 + i * 0.7) * j) + "%" + vbCrLf
                    If j = 7 Then
                        Call Jug2(c, 9 + i * 3)
                        c.Bonustip = c.Bonustip + "来自和璞鸢的增伤：" + CStr(9 + i * 3) + "%" + vbCrLf
                    End If
            End If
        Case "弓藏"
            If InStr(2, CurrSkill, "a") > 0 Then
                Call Jug2(c, 30 + 10 * i)
                c.Bonustip = c.Bonustip + "来自弓藏的增伤：" + CStr(30 + 10 * i) + "%" + vbCrLf
            End If
            If InStr(2, CurrSkill, "c") > 0 Then
                Call Jug2(c, -10)
                c.Bonustip = c.Bonustip + "来自弓藏的增伤：" + CStr(-10) + "%" + vbCrLf
            End If
        Case "决斗之枪"
            j = GetBuffCount("决斗之枪")
            If j < 2 Then
                c.ATKBonus = c.ATKBonus + (18 + i * 6)
                c.ATKtip = c.ATKtip + "来自决斗之枪的攻击力：" + CStr(18 + i * 6) + "%" + vbCrLf
            Else
                c.ATKBonus = c.ATKBonus + (12 + i * 4)
                c.ATKtip = c.ATKtip + "来自决斗之枪的攻击力：" + CStr(12 + i * 4) + "%" + vbCrLf
                c.DEFBonus = c.DEFBonus + (12 + i * 4)
                c.DEFtip = c.DEFtip + "来自决斗之枪的防御力" + CStr(12 + i * 4) + "%" + vbCrLf
            End If
            
        Case "冷刃"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(3).Value = Checked Or FrmMain.CheckState(7).Value = Checked Then
                Call Jug2(c, 9 + i * 3)
                c.Bonustip = c.Bonustip + "来自冷刃的增伤：" + CStr(9 + i * 3) + "%" + vbCrLf
            End If
            
        Case "匣里日月"
            j = GetBuffCount("匣里日月")
            If j = 1 Then
                Call Jug2(c, 15 + i * 5)
                c.Bonustip = c.Bonustip + "来自匣里日月的增伤：" + CStr(9 + i * 3) + "%" + vbCrLf
            End If
        Case "匣里灭辰"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                Call Jug2(c, 16 + i * 4)
                c.Bonustip = c.Bonustip + "来自匣里灭辰的增伤：" + CStr(16 + i * 4) + "%" + vbCrLf
            End If
        Case "匣里龙吟"
            If FrmMain.CheckState(4).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                Call Jug2(c, 16 + i * 4)
                c.Bonustip = c.Bonustip + "来自匣里龙吟的增伤：" + CStr(16 + i * 4) + "%" + vbCrLf
            End If
        Case "千岩古剑"
            j = GetBuffCount("千岩古剑")
            If j > 0 Then
                c.ATKBonus = c.ATKBonus + (6 + i) * j
                c.ATKtip = c.ATKtip + "来自千岩古剑的攻击力：" + CStr((6 + i) * j) + "%" + vbCrLf
                c.CritRate = c.CritRate + (2 + i) * j
                c.CritRatetip = c.CritRatetip + "来自千岩古剑的暴击率：" + CStr((2 + i) * j) + "%" + vbCrLf
            End If
        Case "千岩长枪"
            j = GetBuffCount("千岩长枪")
            If j > 0 Then
                c.ATKBonus = c.ATKBonus + (6 + i) * j
                c.ATKtip = c.ATKtip + "来自千岩长枪的攻击力：" + CStr((6 + i) * j) + "%" + vbCrLf
                c.CritRate = c.CritRate + (2 + i) * j
                c.CritRatetip = c.CritRatetip + "来自千岩长枪的暴击率：" + CStr((2 + i) * j) + "%" + vbCrLf
            End If
                
        Case "喜多院十文字"
            If InStr(2, CurrSkill, "e") > 0 Then AddBonus c, 6, 3, i, WeaponList(c.cWeapon, 1)
            
                
        Case "嘟嘟可故事集"
             j = GetBuffCount("嘟嘟可故事集")
                If j Mod 100 = 1 Then AddBonus c, 8, 1, i, "嘟嘟可故事集"
                If j >= 100 And InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 16, 3, i, "嘟嘟可故事集"
                
                
        Case "四风原典"
            j = GetBuffCount("四风原典")
            AddBonus c, 8 * j, 3, i, "四风原典"
                
        Case "天目影打刀"
                
        Case "天空之傲"
                AddBonus c, 8, 3, i, "天空之傲"
        Case "天空之刃"
                AddBonus c, 4, 4, i, WeaponList(c.cWeapon, 1)
                
        Case "天空之卷"
                AddBonus c, 12, 3, i, "天空之卷"
                
        Case "天空之翼"
                AddBonus c, 20, 5, i, "天空之翼"
        Case "天空之脊"
                AddBonus c, 8, 4, i, WeaponList(c.cWeapon, 1)
                
        Case "宗室大剑"
            j = GetBuffCount("宗室大剑")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "来自宗室大剑的暴击率加成：" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "宗室猎枪"
            j = GetBuffCount("宗室猎枪")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "来自宗室猎枪的暴击率加成：" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "宗室秘法录"
            j = GetBuffCount("宗室秘法录")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "来自宗室秘法录的暴击率加成：" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "宗室长剑"
            j = GetBuffCount("宗室长剑")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "来自宗室长剑的暴击率加成：" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "宗室长弓"
            j = GetBuffCount("宗室长弓")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "来自宗室长弓的暴击率加成：" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "尘世之锁"
            j = GetBuffCount("尘世之锁")
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, "尘世之锁"
                End If
                c.SPower = c.SPower + 15 + i * 5
                
        Case "幽夜华尔兹"
                 j = GetBuffCount(WeaponList(c.cWeapon, 1))
                      If InStr(2, CurrSkill, "e") > 0 And j > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)
                      If InStr(2, CurrSkill, "a") > 0 And j > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)
                      
                
        Case "万国诸海图谱"
            j = GetBuffCount("万国诸海图谱")
                If j > 0 Then
                    Call Jug2(c, (6 + i * 2) * j)
                    c.Bonustip = c.Bonustip + "来自万国诸海图谱的增伤：" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
                
        Case "弹弓"
            j = GetBuffCount("弹弓")
                If j > 0 Then
                    AddBonus c, 36, 3, i, "弹弓"
                Else
                    AddBonus c, -10, 3, 1, "弹弓"
                End If
                
                
        Case "忍冬之果"
                
        Case "护摩之杖"
            c.HPBonus = c.HPBonus + 15 + i * 5
            c.HPtip = c.HPtip + "来自护摩之杖的生命值：" + CStr(15 + i * 5) + "%" + vbCrLf
            t2 = c.MaxHP * (1 + c.HPBonus / 100) + c.HPFlag
            t = Round(t2 * (0.6 + i * 0.2) / 100, 2)
            If c.lowHP Then
                t = t + Round(t2 * (0.8 + i * 0.2) / 100, 2)
            End If
            c.ATKFlag = c.ATKFlag + t
            c.ATKtip = c.ATKtip + "来自护摩之杖的攻击力：" + CStr(t) + vbCrLf
            
                
        Case "斫峰之刃"
             c.SPower = c.SPower + 15 + i * 5
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, WeaponList(c.cWeapon, 1)
                End If
                
        Case "旅行剑"
                
        Case "无工之剑"
            c.SPower = c.SPower + 15 + i * 5
            j = GetBuffCount("无工之剑")
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, "无工之剑"
                End If
                c.SPower = c.SPower + 15 + i * 5
                
                
        Case "昭心"
                
        Case "暗巷猎手"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 2 * j, 3, i, WeaponList(c.cWeapon, 1)
                
                
        Case "暗巷的酒与诗"
            j = GetBuffCount("暗巷的酒与诗")
                If j = 1 Then AddBonus c, 20, 1, i, "暗巷的酒与诗"
                
        Case "暗巷闪光"
            AddBonus c, 12, 3, i, WeaponList(c.cWeapon, 1)
                
        Case "暗铁剑"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 20, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "松籁响起之时"
            AddBonus c, 16, 1, i, "松籁响起之时"
            j = GetBuffCount("松籁响起之时")
                If j = 1 And InStr(1, c.ATKtip, "之歌") <= 0 Then AddBonus c, 20, 1, i, "揭旗之歌"
            
            
                
        Case "桂木斩长正"
             If InStr(2, CurrSkill, "e") > 0 Then AddBonus c, 6, 3, i, "桂木斩长正"
             
        Case "沐浴龙血的剑"
            If FrmMain.CheckState(4).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                AddBonus c, 12, 3, i, "沐浴龙血的剑"
            End If
        Case "流月针"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)

                
        Case "流浪乐章"
            j = GetBuffCount("流浪乐章")
                If j = 1 Then AddBonus c, 60, 1, i, "流浪乐章特效"
                If j = 2 Then AddBonus c, 48, 3, i, "流浪乐章特效"
                If j = 3 Then AddBonus c, 240, 10, i, "流浪乐章特效"
            
                
        Case "狼的末路"
            AddBonus c, 20, 1, i, "狼的末路"
            j = GetBuffCount("狼的末路")
            If j = 1 Then AddBonus c, 40, 1, i, "狼的末路特效"
                
        Case "甲级宝珏"
            j = GetBuffCount("甲级宝珏")
                If j = 1 Then
                    c.ATKBonus = c.ATKBonus + (10 + i * 2)
                    c.ATKtip = c.ATKtip + "来自甲级宝珏的攻击力：" + CStr((10 + i * 2)) + "%" + vbCrLf
                End If
        Case "白影剑"
             j = GetBuffCount("白影剑")
                If j > 0 Then
                    AddBonus c, 6 * j, 1, i, "白影剑"
                    AddBonus c, 6 * j, 8, i, "白影剑"
                End If
             
        Case "白缨枪"
            If InStr(2, CurrSkill, "a") > 0 Then AddBonus c, 24, 3, i, WeaponList(c.cWeapon, 1)
                
        Case "白辰之环"
             j = GetBuffCount("白辰之环")
                If j > 0 And InStr(1, c.Bonustip, "白辰之环") <= 0 Then
                    AddBonus c, 10, 3, i, "白辰之环"
                End If
                
        Case "白铁大剑"
                
        Case "破魔之弓"
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If InStr(2, CurrSkill, "a") > 0 Then AddBonus c, 16 * IIf(j > 0, 2, 1), 3, i, WeaponList(c.cWeapon, 1)
                If InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 12 * IIf(j > 0, 2, 1), 3, i, WeaponList(c.cWeapon, 1)
                
                
        Case "神射手之誓"
             j = GetBuffCount("神射手之誓")
                If j > 0 Then AddBonus c, 24, 3, i, "神射手之誓"
                
        Case "祭礼剑"
                
        Case "祭礼大剑"
                
        Case "祭礼弓"
                
        Case "祭礼残章"
                
        Case "笛剑"
                
        Case "终末嗟叹之诗"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            AddBonus c, 60, 10, i, WeaponList(c.cWeapon, 1)
            If j > 0 Then
                If InStr(1, c.EMtip, "之歌") <= 0 Then AddBonus c, 100, 10, i, "离别之歌"
                If InStr(1, c.ATKtip, "之歌") <= 0 Then AddBonus c, 20, 1, i, "离别之歌"
            End If
                
                
        Case "绝弦"
             If InStr(2, CurrSkill, "e") > 0 Or InStr(2, CurrSkill, "q") > 0 Then AddBonus c, 24, 3, i, "绝弦"
             
             
        Case "翡玉法球"
            j = GetBuffCount("翡玉法球")
                If j = 1 Then
                    c.ATKBonus = c.ATKBonus + (15 + i * 5)
                    c.ATKtip = c.ATKtip + "来自翡玉法球的攻击力：" + CStr((15 + i * 5)) + "%" + vbCrLf
                End If
                
        Case "腐殖之剑"
            If InStr(2, CurrSkill, "e") > 0 Then
                AddBonus c, 16, 3, i, WeaponList(c.cWeapon, 1)
                AddBonus c, 6, 4, i, WeaponList(c.cWeapon, 1)
            End If
                
        Case "苍古自由之誓"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            AddBonus c, 10, 3, i, WeaponList(c.cWeapon, 1)
            If j > 0 Then
                  If (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Or InStr(2, CurrSkill, "d") > 0) And (InStr(1, c.Bonustip, "之歌") <= 0) Then AddBonus c, 16, 3, i, "抗争之歌"
                  If InStr(1, c.ATKtip, "之歌") <= 0 Then AddBonus c, 20, 1, i, "抗争之歌"
            End If
                
        Case "苍翠猎弓"
                
        Case "薙草之稻光"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then c.Energy = c.Energy + 25 + i * 5
            t = c.Energy - 100
            t = t * (21 + i * 7) / 100
            If t > (70 + i * 10) Then t = (70 + i * 10)
            AddBonus c, t, 1, 1, "薙草之稻光"
             
                
        Case "螭骨剑"
            j = GetBuffCount("螭骨剑")
                If j > 0 Then AddBonus c, 6 * j, 3, i, "螭骨剑"
                


                
                
        Case "试作古华"
                
        Case "试作斩岩"
             j = GetBuffCount("试作斩岩")
                If j > 0 Then
                    AddBonus c, 4 * j, 1, i, "试作斩岩"
                    AddBonus c, 4 * j, 8, i, "试作斩岩"
                End If
                
        Case "试作星镰"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 And (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0) Then AddBonus c, 8 * j, 3, i, WeaponList(c.cWeapon, 1)

            
                
        Case "试作澹月"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 36 * j, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "贯虹之槊"
            c.SPower = c.SPower + 15 + i * 5
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, WeaponList(c.cWeapon, 1)
                End If
                
                
        Case "钟剑"
                If c.InSheild Then AddBonus c, 12, 3, i, "钟剑（有盾）"
                
        Case "钢轮弓"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 4 * j, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "钺矛"
                
        Case "铁影阔剑"
            j = GetBuffCount("铁影阔剑")
                If j > 0 Then AddBonus c, 30, 3, i, "铁影阔剑"
                
                
                
        Case "铁蜂刺"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 6 * j, 3, i, WeaponList(c.cWeapon, 1)
        
                
        Case "阿莫斯之弓"
            If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                AddBonus c, 12 + 8 * j, 3, i, WeaponList(c.cWeapon, 1)
            End If
        
                
        Case "雨裁"
            If FrmMain.CheckState(4).Value = Checked Or FrmMain.CheckState(2).Value = Checked Then
                AddBonus c, 20, 3, i, "雨裁"
            End If
            
        Case "雪葬的星银"
                
        Case "雾切之回光"
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                AddBonus c, 12, 3, i, WeaponList(c.cWeapon, 1), True
                If j > 0 Then
                    Select Case CharList(Val(FrmMain.AlphaImageChar.tag), 4)
                        Case "岩"
                            c.GeoDMG = c.GeoDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "来自雾切之回光的岩元素伤害加成：" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "风"
                            c.AnemoDMG = c.AnemoDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "来自雾切之回光的风元素伤害加成：" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "雷"
                            c.ElectroDMG = c.ElectroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "来自雾切之回光的雷元素伤害加成：" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "火"
                            c.PyroDMG = c.PyroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "来自雾切之回光的火元素伤害加成：" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "冰"
                            c.CryoDMG = c.CryoDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "来自雾切之回光的冰元素伤害加成：" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "水"
                            c.HydroDMG = c.HydroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "来自雾切之回光的水元素伤害加成：" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "草"
                            c.DendroDMG = c.DendroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "来自雾切之回光的草元素伤害加成：" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                    End Select
                End If
                
                
                
        Case "风花之颂"
             If InStr(2, CurrSkill, "e") > 0 Then
                 AddBonus c, 16, 1, i, WeaponList(c.cWeapon, 1)
             Else
                    j = GetBuffCount(WeaponList(c.cWeapon, 1))
                        If j > 0 Then AddBonus c, 16, 1, i, WeaponList(c.cWeapon, 1)
             End If
             
        Case "风鹰剑"
            AddBonus c, 20, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "飞天大御剑"
            j = GetBuffCount("铁影阔剑")
                If j > 0 Then AddBonus c, 6 * j, 1, i, "飞天大御剑"
                
        Case "飞天御剑"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "飞雷之弦振"
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                AddBonus c, 20, 1, i, WeaponList(c.cWeapon, 1), True
                If InStr(2, CurrSkill, "a") > 0 And j > 0 Then AddBonus c, IIf(j = 3, 40, j * 12), 3, i, WeaponList(c.cWeapon, 1)
        
                
        Case "魔导绪论"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(4).Value = Checked Then
                Call Jug2(c, 9 + i * 3)
                c.Bonustip = c.Bonustip + "来自魔导绪论的增伤：" + CStr(9 + i * 3) + "%" + vbCrLf
            End If
            
        Case "鸦羽弓"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                AddBonus c, 12, 3, i, "鸦羽弓"
            End If
            
            
        Case "黎明神剑"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
               AddBonus c, 14, 4, i, WeaponList(c.cWeapon, 1)
                
                
        Case "黑剑"
            If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)
                
        Case "黑岩刺枪"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12 * j, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "黑岩战弓"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12 * j, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "黑岩斩刀"
            j = GetBuffCount("黑岩斩刀")
                If j > 0 Then
                    c.ATKBonus = c.ATKBonus + (9 + i * 3) * j
                    c.ATKtip = c.ATKtip + "来自黑岩斩刀的攻击力：" + CStr((9 + i * 3) * j) + "%" + vbCrLf
                End If
                
        Case "黑岩绯玉"
            j = GetBuffCount("黑岩绯玉")
                If j > 0 Then
                    c.ATKBonus = c.ATKBonus + (9 + i * 3) * j
                    c.ATKtip = c.ATKtip + "来自黑岩绯玉的攻击力：" + CStr((9 + i * 3) * j) + "%" + vbCrLf
                End If
        Case "黑岩长剑"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12 * j, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "黑缨枪"
            If FrmMain.BuffComboBox1.Text = "史莱姆" Then AddBonus c, 40, 3, i, WeaponList(c.cWeapon, 1)
            
                
        Case "龙脊长枪"

        Case "「渔获」"
            If InStr(2, CurrSkill, "q") > 0 Then
                AddBonus c, 16, 3, i, WeaponList(c.cWeapon, 1)
                AddBonus c, 6, 4, i, WeaponList(c.cWeapon, 1)
            End If
            
            
    End Select
End Sub

Public Function GetBuffCount(s As String) As Integer
Dim i%
    For i = 1 To SelectCount
        If Mid(FrmMain.SelectBuffLabel(i), 1, Len(s)) = s And FrmMain.SelectBuffBox(i).Visible = True Then
            If FrmMain.BuffCheck(i).Visible = False Then '托条模式
                GetBuffCount = Val(FrmMain.BuffLabel(i).tag)
                    If FrmMain.BuffCheck2(i).Value = Checked And FrmMain.BuffCheck2(i).Visible Then
                        GetBuffCount = GetBuffCount + 100
                    End If
                Exit Function
            Else
                If FrmMain.BuffCheck(i).Value = Checked Then
                    GetBuffCount = 1
                Else
                    GetBuffCount = 0
                End If
                If FrmMain.BuffCheck2(i).Value = Checked And FrmMain.BuffCheck2(i).Visible Then
                    GetBuffCount = GetBuffCount + 100
                Else
                    GetBuffCount = GetBuffCount + 0
                End If
            
            End If
        End If
        
    Next
End Function


Public Sub AddBonus(ByRef c As Chars, v As Single, Atype As Integer, rrank As Integer, tip As String, Optional ele As Boolean)
Dim M As Single
M = Round(v * 3 / 4 + v / 4 * rrank, 1)
    If Atype = 1 Then
        c.ATKBonus = c.ATKBonus + M
        c.ATKtip = c.ATKtip + "来自" + tip + "的攻击力：" + CStr(M) + "%" + vbCrLf
    End If
    
    If Atype = 2 Then
        c.ATKFlag = c.ATKFlag + M
        c.ATKtip = c.ATKtip + "来自" + tip + "的攻击力：" + CStr(M) + vbCrLf
    End If
    
    
    If Atype = 3 Then
        If Not IsMissing(ele) Then
            Call Jug2(c, M, ele)
         Else
            Call Jug2(c, M)
         End If
        c.Bonustip = c.Bonustip + "来自" + tip + "的伤害加成：" + CStr(M) + "%" + vbCrLf
    End If
    
    If Atype = 4 Then
        c.CritRate = c.CritRate + M
        c.CritRatetip = c.CritRatetip + "来自" + tip + "的暴击率：" + CStr(M) + "%" + vbCrLf
    End If
        
    If Atype = 5 Then
        c.CritDmg = c.CritDmg + M
        c.CritDMGtip = c.CritDMGtip + "来自" + tip + "的暴击伤害：" + CStr(M) + "%" + vbCrLf
    End If
    
    If Atype = 6 Then
        c.HPBonus = c.HPBonus + M
        c.HPtip = c.HPtip + "来自" + tip + "的生命值：" + CStr(M) + "%" + vbCrLf
     End If
     
     If Atype = 7 Then
        c.HPFlag = c.HPFlag + M
        c.HPtip = c.HPtip + "来自" + tip + "的生命值：" + CStr(M) + vbCrLf
     End If
     
     If Atype = 8 Then
        c.DEFBonus = c.DEFBonus + M
        c.DEFtip = c.DEFtip + "来自" + tip + "的防御力：" + CStr(M) + "%" + vbCrLf
     End If
     
     If Atype = 9 Then
        c.DEFFlag = c.DEFFlag + M
        c.DEFtip = c.DEFtip + "来自" + tip + "的防御力：" + CStr(M) + vbCrLf
     End If
     
     If Atype = 10 Then
        c.EM = c.EM + M
        c.EMtip = c.EMtip + "来自" + tip + "的元素精通：" + CStr(M) + vbCrLf
     End If
End Sub
