Attribute VB_Name = "Module2"
Public Sub SolveBonus(ByRef c As Chars)
Dim i%, t As Single, j%, t2 As Single, flag As Boolean
    i = RBoxFlag + 1 '�����ȼ�
    j = 0
    Select Case WeaponList(c.cWeapon, 1)
        Case "���ҽ���"
            c.HPBonus = 15 + i * 5
            c.HPtip = c.HPtip + "�������ҽ��̵�����ֵ��" + CStr(15 + i * 5) + "%" + vbCrLf
            t = c.MaxHP * (1 + c.HPBonus / 100) + c.HPFlag
            t = Round(t * (0.9 + i * 0.3) / 100, 2)
            c.ATKFlag = c.ATKFlag + t
            c.ATKtip = c.ATKtip + "�������ҽ��̵Ĺ�������" + CStr(t) + vbCrLf
        Case "����"
            j = GetBuffCount("����")
            If j > 0 Then
                c.ATKBonus = c.ATKBonus + (2.5 + i * 0.7) * j
                c.ATKtip = c.ATKtip + "���Ժ���Ĺ�������" + CStr((2.5 + i * 0.7) * j) + "%" + vbCrLf
                    If j = 7 Then
                        Call Jug2(c, 9 + i * 3)
                        c.Bonustip = c.Bonustip + "���Ժ�������ˣ�" + CStr(9 + i * 3) + "%" + vbCrLf
                    End If
            End If
        Case "����"
            If InStr(2, CurrSkill, "a") > 0 Then
                Call Jug2(c, 30 + 10 * i)
                c.Bonustip = c.Bonustip + "���Թ��ص����ˣ�" + CStr(30 + 10 * i) + "%" + vbCrLf
            End If
            If InStr(2, CurrSkill, "c") > 0 Then
                Call Jug2(c, -10)
                c.Bonustip = c.Bonustip + "���Թ��ص����ˣ�" + CStr(-10) + "%" + vbCrLf
            End If
        Case "����֮ǹ"
            j = GetBuffCount("����֮ǹ")
            If j < 2 Then
                c.ATKBonus = c.ATKBonus + (18 + i * 6)
                c.ATKtip = c.ATKtip + "���Ծ���֮ǹ�Ĺ�������" + CStr(18 + i * 6) + "%" + vbCrLf
            Else
                c.ATKBonus = c.ATKBonus + (12 + i * 4)
                c.ATKtip = c.ATKtip + "���Ծ���֮ǹ�Ĺ�������" + CStr(12 + i * 4) + "%" + vbCrLf
                c.DEFBonus = c.DEFBonus + (12 + i * 4)
                c.DEFtip = c.DEFtip + "���Ծ���֮ǹ�ķ�����" + CStr(12 + i * 4) + "%" + vbCrLf
            End If
            
        Case "����"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(3).Value = Checked Or FrmMain.CheckState(7).Value = Checked Then
                Call Jug2(c, 9 + i * 3)
                c.Bonustip = c.Bonustip + "�������е����ˣ�" + CStr(9 + i * 3) + "%" + vbCrLf
            End If
            
        Case "ϻ������"
            j = GetBuffCount("ϻ������")
            If j = 1 Then
                Call Jug2(c, 15 + i * 5)
                c.Bonustip = c.Bonustip + "����ϻ�����µ����ˣ�" + CStr(9 + i * 3) + "%" + vbCrLf
            End If
        Case "ϻ����"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                Call Jug2(c, 16 + i * 4)
                c.Bonustip = c.Bonustip + "����ϻ���𳽵����ˣ�" + CStr(16 + i * 4) + "%" + vbCrLf
            End If
        Case "ϻ������"
            If FrmMain.CheckState(4).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                Call Jug2(c, 16 + i * 4)
                c.Bonustip = c.Bonustip + "����ϻ�����������ˣ�" + CStr(16 + i * 4) + "%" + vbCrLf
            End If
        Case "ǧ�ҹŽ�"
            j = GetBuffCount("ǧ�ҹŽ�")
            If j > 0 Then
                c.ATKBonus = c.ATKBonus + (6 + i) * j
                c.ATKtip = c.ATKtip + "����ǧ�ҹŽ��Ĺ�������" + CStr((6 + i) * j) + "%" + vbCrLf
                c.CritRate = c.CritRate + (2 + i) * j
                c.CritRatetip = c.CritRatetip + "����ǧ�ҹŽ��ı����ʣ�" + CStr((2 + i) * j) + "%" + vbCrLf
            End If
        Case "ǧ�ҳ�ǹ"
            j = GetBuffCount("ǧ�ҳ�ǹ")
            If j > 0 Then
                c.ATKBonus = c.ATKBonus + (6 + i) * j
                c.ATKtip = c.ATKtip + "����ǧ�ҳ�ǹ�Ĺ�������" + CStr((6 + i) * j) + "%" + vbCrLf
                c.CritRate = c.CritRate + (2 + i) * j
                c.CritRatetip = c.CritRatetip + "����ǧ�ҳ�ǹ�ı����ʣ�" + CStr((2 + i) * j) + "%" + vbCrLf
            End If
                
        Case "ϲ��Ժʮ����"
            If InStr(2, CurrSkill, "e") > 0 Then AddBonus c, 6, 3, i, WeaponList(c.cWeapon, 1)
            
                
        Case "�ཿɹ��¼�"
             j = GetBuffCount("�ཿɹ��¼�")
                If j Mod 100 = 1 Then AddBonus c, 8, 1, i, "�ཿɹ��¼�"
                If j >= 100 And InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 16, 3, i, "�ཿɹ��¼�"
                
                
        Case "�ķ�ԭ��"
            j = GetBuffCount("�ķ�ԭ��")
            AddBonus c, 8 * j, 3, i, "�ķ�ԭ��"
                
        Case "��ĿӰ��"
                
        Case "���֮��"
                AddBonus c, 8, 3, i, "���֮��"
        Case "���֮��"
                AddBonus c, 4, 4, i, WeaponList(c.cWeapon, 1)
                
        Case "���֮��"
                AddBonus c, 12, 3, i, "���֮��"
                
        Case "���֮��"
                AddBonus c, 20, 5, i, "���֮��"
        Case "���֮��"
                AddBonus c, 8, 4, i, WeaponList(c.cWeapon, 1)
                
        Case "���Ҵ�"
            j = GetBuffCount("���Ҵ�")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "�������Ҵ󽣵ı����ʼӳɣ�" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "������ǹ"
            j = GetBuffCount("������ǹ")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "����������ǹ�ı����ʼӳɣ�" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "�����ط�¼"
            j = GetBuffCount("�����ط�¼")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "���������ط�¼�ı����ʼӳɣ�" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "���ҳ���"
            j = GetBuffCount("���ҳ���")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "�������ҳ����ı����ʼӳɣ�" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "���ҳ���"
            j = GetBuffCount("���ҳ���")
                If j > 0 Then
                    c.CritRate = c.CritRate + (6 + i * 2) * j
                    c.CritRatetip = c.CritRatetip + "�������ҳ����ı����ʼӳɣ�" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
        Case "����֮��"
            j = GetBuffCount("����֮��")
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, "����֮��"
                End If
                c.SPower = c.SPower + 15 + i * 5
                
        Case "��ҹ������"
                 j = GetBuffCount(WeaponList(c.cWeapon, 1))
                      If InStr(2, CurrSkill, "e") > 0 And j > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)
                      If InStr(2, CurrSkill, "a") > 0 And j > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)
                      
                
        Case "����ͼ��"
            j = GetBuffCount("����ͼ��")
                If j > 0 Then
                    Call Jug2(c, (6 + i * 2) * j)
                    c.Bonustip = c.Bonustip + "��������ͼ�׵����ˣ�" + CStr((6 + i * 2) * j) + "%" + vbCrLf
                End If
                
        Case "����"
            j = GetBuffCount("����")
                If j > 0 Then
                    AddBonus c, 36, 3, i, "����"
                Else
                    AddBonus c, -10, 3, 1, "����"
                End If
                
                
        Case "�̶�֮��"
                
        Case "��Ħ֮��"
            c.HPBonus = c.HPBonus + 15 + i * 5
            c.HPtip = c.HPtip + "���Ի�Ħ֮�ȵ�����ֵ��" + CStr(15 + i * 5) + "%" + vbCrLf
            t2 = c.MaxHP * (1 + c.HPBonus / 100) + c.HPFlag
            t = Round(t2 * (0.6 + i * 0.2) / 100, 2)
            If c.lowHP Then
                t = t + Round(t2 * (0.8 + i * 0.2) / 100, 2)
            End If
            c.ATKFlag = c.ATKFlag + t
            c.ATKtip = c.ATKtip + "���Ի�Ħ֮�ȵĹ�������" + CStr(t) + vbCrLf
            
                
        Case "��֮��"
             c.SPower = c.SPower + 15 + i * 5
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, WeaponList(c.cWeapon, 1)
                End If
                
        Case "���н�"
                
        Case "�޹�֮��"
            c.SPower = c.SPower + 15 + i * 5
            j = GetBuffCount("�޹�֮��")
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, "�޹�֮��"
                End If
                c.SPower = c.SPower + 15 + i * 5
                
                
        Case "����"
                
        Case "��������"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 2 * j, 3, i, WeaponList(c.cWeapon, 1)
                
                
        Case "����ľ���ʫ"
            j = GetBuffCount("����ľ���ʫ")
                If j = 1 Then AddBonus c, 20, 1, i, "����ľ���ʫ"
                
        Case "��������"
            AddBonus c, 12, 3, i, WeaponList(c.cWeapon, 1)
                
        Case "������"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 20, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "��������֮ʱ"
            AddBonus c, 16, 1, i, "��������֮ʱ"
            j = GetBuffCount("��������֮ʱ")
                If j = 1 And InStr(1, c.ATKtip, "֮��") <= 0 Then AddBonus c, 20, 1, i, "����֮��"
            
            
                
        Case "��ľն����"
             If InStr(2, CurrSkill, "e") > 0 Then AddBonus c, 6, 3, i, "��ľն����"
             
        Case "��ԡ��Ѫ�Ľ�"
            If FrmMain.CheckState(4).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                AddBonus c, 12, 3, i, "��ԡ��Ѫ�Ľ�"
            End If
        Case "������"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)

                
        Case "��������"
            j = GetBuffCount("��������")
                If j = 1 Then AddBonus c, 60, 1, i, "����������Ч"
                If j = 2 Then AddBonus c, 48, 3, i, "����������Ч"
                If j = 3 Then AddBonus c, 240, 10, i, "����������Ч"
            
                
        Case "�ǵ�ĩ·"
            AddBonus c, 20, 1, i, "�ǵ�ĩ·"
            j = GetBuffCount("�ǵ�ĩ·")
            If j = 1 Then AddBonus c, 40, 1, i, "�ǵ�ĩ·��Ч"
                
        Case "�׼�����"
            j = GetBuffCount("�׼�����")
                If j = 1 Then
                    c.ATKBonus = c.ATKBonus + (10 + i * 2)
                    c.ATKtip = c.ATKtip + "���Լ׼�����Ĺ�������" + CStr((10 + i * 2)) + "%" + vbCrLf
                End If
        Case "��Ӱ��"
             j = GetBuffCount("��Ӱ��")
                If j > 0 Then
                    AddBonus c, 6 * j, 1, i, "��Ӱ��"
                    AddBonus c, 6 * j, 8, i, "��Ӱ��"
                End If
             
        Case "��ӧǹ"
            If InStr(2, CurrSkill, "a") > 0 Then AddBonus c, 24, 3, i, WeaponList(c.cWeapon, 1)
                
        Case "�׳�֮��"
             j = GetBuffCount("�׳�֮��")
                If j > 0 And InStr(1, c.Bonustip, "�׳�֮��") <= 0 Then
                    AddBonus c, 10, 3, i, "�׳�֮��"
                End If
                
        Case "������"
                
        Case "��ħ֮��"
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If InStr(2, CurrSkill, "a") > 0 Then AddBonus c, 16 * IIf(j > 0, 2, 1), 3, i, WeaponList(c.cWeapon, 1)
                If InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 12 * IIf(j > 0, 2, 1), 3, i, WeaponList(c.cWeapon, 1)
                
                
        Case "������֮��"
             j = GetBuffCount("������֮��")
                If j > 0 Then AddBonus c, 24, 3, i, "������֮��"
                
        Case "����"
                
        Case "�����"
                
        Case "����"
                
        Case "�������"
                
        Case "�ѽ�"
                
        Case "��ĩ�֮̾ʫ"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            AddBonus c, 60, 10, i, WeaponList(c.cWeapon, 1)
            If j > 0 Then
                If InStr(1, c.EMtip, "֮��") <= 0 Then AddBonus c, 100, 10, i, "���֮��"
                If InStr(1, c.ATKtip, "֮��") <= 0 Then AddBonus c, 20, 1, i, "���֮��"
            End If
                
                
        Case "����"
             If InStr(2, CurrSkill, "e") > 0 Or InStr(2, CurrSkill, "q") > 0 Then AddBonus c, 24, 3, i, "����"
             
             
        Case "������"
            j = GetBuffCount("������")
                If j = 1 Then
                    c.ATKBonus = c.ATKBonus + (15 + i * 5)
                    c.ATKtip = c.ATKtip + "����������Ĺ�������" + CStr((15 + i * 5)) + "%" + vbCrLf
                End If
                
        Case "��ֳ֮��"
            If InStr(2, CurrSkill, "e") > 0 Then
                AddBonus c, 16, 3, i, WeaponList(c.cWeapon, 1)
                AddBonus c, 6, 4, i, WeaponList(c.cWeapon, 1)
            End If
                
        Case "�Թ�����֮��"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            AddBonus c, 10, 3, i, WeaponList(c.cWeapon, 1)
            If j > 0 Then
                  If (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Or InStr(2, CurrSkill, "d") > 0) And (InStr(1, c.Bonustip, "֮��") <= 0) Then AddBonus c, 16, 3, i, "����֮��"
                  If InStr(1, c.ATKtip, "֮��") <= 0 Then AddBonus c, 20, 1, i, "����֮��"
            End If
                
        Case "�Դ��Թ�"
                
        Case "�S��֮����"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then c.Energy = c.Energy + 25 + i * 5
            t = c.Energy - 100
            t = t * (21 + i * 7) / 100
            If t > (70 + i * 10) Then t = (70 + i * 10)
            AddBonus c, t, 1, 1, "�S��֮����"
             
                
        Case "�ǽ�"
            j = GetBuffCount("�ǽ�")
                If j > 0 Then AddBonus c, 6 * j, 3, i, "�ǽ�"
                


                
                
        Case "�����Ż�"
                
        Case "����ն��"
             j = GetBuffCount("����ն��")
                If j > 0 Then
                    AddBonus c, 4 * j, 1, i, "����ն��"
                    AddBonus c, 4 * j, 8, i, "����ն��"
                End If
                
        Case "��������"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 And (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0) Then AddBonus c, 8 * j, 3, i, WeaponList(c.cWeapon, 1)

            
                
        Case "�������"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 36 * j, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "���֮��"
            c.SPower = c.SPower + 15 + i * 5
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
            If j >= 100 Then
                j = j - 100
                c.InSheild = True
            End If
                If j > 0 Then
                    AddBonus c, 4 * j * IIf(c.InSheild, 2, 1), 1, i, WeaponList(c.cWeapon, 1)
                End If
                
                
        Case "�ӽ�"
                If c.InSheild Then AddBonus c, 12, 3, i, "�ӽ����жܣ�"
                
        Case "���ֹ�"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 4 * j, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "��ì"
                
        Case "��Ӱ����"
            j = GetBuffCount("��Ӱ����")
                If j > 0 Then AddBonus c, 30, 3, i, "��Ӱ����"
                
                
                
        Case "�����"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 6 * j, 3, i, WeaponList(c.cWeapon, 1)
        
                
        Case "��Ī˹֮��"
            If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                AddBonus c, 12 + 8 * j, 3, i, WeaponList(c.cWeapon, 1)
            End If
        
                
        Case "���"
            If FrmMain.CheckState(4).Value = Checked Or FrmMain.CheckState(2).Value = Checked Then
                AddBonus c, 20, 3, i, "���"
            End If
            
        Case "ѩ�������"
                
        Case "����֮�ع�"
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                AddBonus c, 12, 3, i, WeaponList(c.cWeapon, 1), True
                If j > 0 Then
                    Select Case CharList(Val(FrmMain.AlphaImageChar.tag), 4)
                        Case "��"
                            c.GeoDMG = c.GeoDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "��������֮�ع����Ԫ���˺��ӳɣ�" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "��"
                            c.AnemoDMG = c.AnemoDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "��������֮�ع�ķ�Ԫ���˺��ӳɣ�" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "��"
                            c.ElectroDMG = c.ElectroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "��������֮�ع����Ԫ���˺��ӳɣ�" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "��"
                            c.PyroDMG = c.PyroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "��������֮�ع�Ļ�Ԫ���˺��ӳɣ�" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "��"
                            c.CryoDMG = c.CryoDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "��������֮�ع�ı�Ԫ���˺��ӳɣ�" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "ˮ"
                            c.HydroDMG = c.HydroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "��������֮�ع��ˮԪ���˺��ӳɣ�" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                        Case "��"
                            c.DendroDMG = c.DendroDMG + IIf(j = 3, 21 + i * 7, j * (6 + i * 2))
                            c.Bonustip = c.Bonustip + "��������֮�ع�Ĳ�Ԫ���˺��ӳɣ�" + CStr(IIf(j = 3, 21 + i * 7, j * (6 + i * 2))) + "%" + vbCrLf
                    End Select
                End If
                
                
                
        Case "�绨֮��"
             If InStr(2, CurrSkill, "e") > 0 Then
                 AddBonus c, 16, 1, i, WeaponList(c.cWeapon, 1)
             Else
                    j = GetBuffCount(WeaponList(c.cWeapon, 1))
                        If j > 0 Then AddBonus c, 16, 1, i, WeaponList(c.cWeapon, 1)
             End If
             
        Case "��ӥ��"
            AddBonus c, 20, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "���������"
            j = GetBuffCount("��Ӱ����")
                If j > 0 Then AddBonus c, 6 * j, 1, i, "���������"
                
        Case "��������"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12, 1, i, WeaponList(c.cWeapon, 1)
                
                
        Case "����֮����"
                j = GetBuffCount(WeaponList(c.cWeapon, 1))
                AddBonus c, 20, 1, i, WeaponList(c.cWeapon, 1), True
                If InStr(2, CurrSkill, "a") > 0 And j > 0 Then AddBonus c, IIf(j = 3, 40, j * 12), 3, i, WeaponList(c.cWeapon, 1)
        
                
        Case "ħ������"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(4).Value = Checked Then
                Call Jug2(c, 9 + i * 3)
                c.Bonustip = c.Bonustip + "����ħ�����۵����ˣ�" + CStr(9 + i * 3) + "%" + vbCrLf
            End If
            
        Case "ѻ��"
            If FrmMain.CheckState(2).Value = Checked Or FrmMain.CheckState(1).Value = Checked Then
                AddBonus c, 12, 3, i, "ѻ��"
            End If
            
            
        Case "������"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
               AddBonus c, 14, 4, i, WeaponList(c.cWeapon, 1)
                
                
        Case "�ڽ�"
            If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 20, 3, i, WeaponList(c.cWeapon, 1)
                
        Case "���Ҵ�ǹ"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12 * j, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "����ս��"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12 * j, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "����ն��"
            j = GetBuffCount("����ն��")
                If j > 0 Then
                    c.ATKBonus = c.ATKBonus + (9 + i * 3) * j
                    c.ATKtip = c.ATKtip + "���Ժ���ն���Ĺ�������" + CStr((9 + i * 3) * j) + "%" + vbCrLf
                End If
                
        Case "�������"
            j = GetBuffCount("�������")
                If j > 0 Then
                    c.ATKBonus = c.ATKBonus + (9 + i * 3) * j
                    c.ATKtip = c.ATKtip + "���Ժ������Ĺ�������" + CStr((9 + i * 3) * j) + "%" + vbCrLf
                End If
        Case "���ҳ���"
            j = GetBuffCount(WeaponList(c.cWeapon, 1))
                If j > 0 Then AddBonus c, 12 * j, 1, i, WeaponList(c.cWeapon, 1)
                
        Case "��ӧǹ"
            If FrmMain.BuffComboBox1.Text = "ʷ��ķ" Then AddBonus c, 40, 3, i, WeaponList(c.cWeapon, 1)
            
                
        Case "������ǹ"

        Case "�����"
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
            If FrmMain.BuffCheck(i).Visible = False Then '����ģʽ
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
        c.ATKtip = c.ATKtip + "����" + tip + "�Ĺ�������" + CStr(M) + "%" + vbCrLf
    End If
    
    If Atype = 2 Then
        c.ATKFlag = c.ATKFlag + M
        c.ATKtip = c.ATKtip + "����" + tip + "�Ĺ�������" + CStr(M) + vbCrLf
    End If
    
    
    If Atype = 3 Then
        If Not IsMissing(ele) Then
            Call Jug2(c, M, ele)
         Else
            Call Jug2(c, M)
         End If
        c.Bonustip = c.Bonustip + "����" + tip + "���˺��ӳɣ�" + CStr(M) + "%" + vbCrLf
    End If
    
    If Atype = 4 Then
        c.CritRate = c.CritRate + M
        c.CritRatetip = c.CritRatetip + "����" + tip + "�ı����ʣ�" + CStr(M) + "%" + vbCrLf
    End If
        
    If Atype = 5 Then
        c.CritDmg = c.CritDmg + M
        c.CritDMGtip = c.CritDMGtip + "����" + tip + "�ı����˺���" + CStr(M) + "%" + vbCrLf
    End If
    
    If Atype = 6 Then
        c.HPBonus = c.HPBonus + M
        c.HPtip = c.HPtip + "����" + tip + "������ֵ��" + CStr(M) + "%" + vbCrLf
     End If
     
     If Atype = 7 Then
        c.HPFlag = c.HPFlag + M
        c.HPtip = c.HPtip + "����" + tip + "������ֵ��" + CStr(M) + vbCrLf
     End If
     
     If Atype = 8 Then
        c.DEFBonus = c.DEFBonus + M
        c.DEFtip = c.DEFtip + "����" + tip + "�ķ�������" + CStr(M) + "%" + vbCrLf
     End If
     
     If Atype = 9 Then
        c.DEFFlag = c.DEFFlag + M
        c.DEFtip = c.DEFtip + "����" + tip + "�ķ�������" + CStr(M) + vbCrLf
     End If
     
     If Atype = 10 Then
        c.EM = c.EM + M
        c.EMtip = c.EMtip + "����" + tip + "��Ԫ�ؾ�ͨ��" + CStr(M) + vbCrLf
     End If
End Sub
