Attribute VB_Name = "Module3"
Public Sub SolveCharBonus(ByRef c As Chars)
Dim i%, t As Single, temp As Variant, j As Integer
    i = CBoxFlag + 1 '�����ȼ�
        Select Case CharList(c.cNumber, 1)
            Case "������"
                If CurrSkill = "c1e2" Then
                j = GetBuffCount("�������츳2")
                       If j > 0 Then AddBonus c, 25, 3, 1, CharList(c.cNumber, 1) + "�츳2"
                End If
                If InStr(2, CurrSkill, "d") > 0 Then
                j = GetBuffCount("����������4")
                       If j > 0 Then AddBonus c, 30, 3, 1, CharList(c.cNumber, 1) + "����4"
                End If
                
            Case "����"
                j = GetBuffCount("�����츳2")
                If j > 0 Then c.PyroDMG = c.PyroDMG + 2 * j: c.Bonustip = c.Bonustip + "���������츳2�Ļ�Ԫ���˺��ӳɣ�" + CStr(2 * j) + "%" + vbCrLf
                j = GetBuffCount("��������1")
                If j > 0 Then AddBonus c, 20, 1, 1, CharList(c.cNumber, 1) + "����1"
                j = GetBuffCount("��������2")
                If j > 0 Then c.PyroDMG = c.PyroDMG + 25: c.Bonustip = c.Bonustip + "������������2�Ļ�Ԫ���˺��ӳɣ�25%" + vbCrLf
                
            Case "�Ű���"
                j = GetBuffCount("�Ű�������2")
                If j > 0 Then c.HydroDMG = c.HydroDMG + 15: c.Bonustip = c.Bonustip + "���԰Ű�������2��ˮԪ���˺��ӳɣ�15%" + vbCrLf
                
                
            Case "��"
                j = GetBuffCount("�̣�����")
                If j > 0 Then j = (Int((j - 1) / 3) + 1) * 5
                If (CurrSkill = "c2d1" Or CurrSkill = "c2d2" Or CurrSkill = "c2d3" Or CurrSkill = "c2a1" Or CurrSkill = "c2c1") Then
                If j > 0 Then
                    FrmMain.Label2(2).Caption = "��"
                    temp = Array(58.45, 61.95, 65.45, 70, 73.5, 77, 81.55, 86.1, 90.65, 95.2, 99.75, 104.3, 108.85, 113.4, 117.95)
                    c.AnemoDMG = c.AnemoDMG + temp(FrmMain.LevelBox(2).ListIndex - 1) + j
                    c.Bonustip = c.Bonustip + "����Q���ܵ����ˣ�" + CStr(temp(FrmMain.LevelBox(2).ListIndex - 1)) + "%" + vbCrLf
                    c.Bonustip = c.Bonustip + "�������츳2�����ˣ�" + CStr(j) + "%" + vbCrLf
                Else
                    FrmMain.Label2(2).Caption = "����"
                End If
                End If
                
            Case "����"
                If CurrSkill = "c5q1" And c.Level >= 42 Then
                    c.CritRate = c.CritRate + 10
                    c.CritRatetip = c.CritRatetip + "���԰����츳2�ı����ʣ�10%" + vbCrLf
                    If CBoxFlag = 5 Then AddBonus c, 15, 1, 1, "��������6"
                End If
                
            Case "����"
                If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then
                    j = GetBuffCount("�����츳2")
                    If j > 0 Then AddBonus c, 15, 3, 1, "�����츳2"
                End If
                
            Case "����"
                j = GetBuffCount("���ң��Ƿ�")
                temp = Array(3.84, 4.07, 4.3, 4.6, 4.83, 5.06, 5.36, 5.66, 5.96, 6.26, 6.56, 6.85, 7.15, 7.45, 7.75)
                t = c.MaxHP * (1 + c.HPBonus / 100) + c.HPFlag
                t = t * temp(FrmMain.LevelBox(1).ListIndex - 1) / 100
                If j > 0 Then AddBonus c, t, 2, 1, "����Ԫ��ս��"

                If c.lowHP Then
                    c.PyroDMG = c.PyroDMG + 33
                    c.Bonustip = c.Bonustip + "���Ժ����츳3�Ļ�Ԫ���˺��ӳɣ�33%" + vbCrLf
                End If
            
            
            Case "��¬��"
                j = GetBuffCount("��¬�ˣ��Ƿ���")
                If j > 0 Then
                    c.PyroDMG = c.PyroDMG + 20
                    c.Bonustip = c.Bonustip + "���Ե�¬���츳3�Ļ�Ԫ���˺��ӳɣ�20%" + vbCrLf
                End If
                j = GetBuffCount("��¬������1")
                    If j > 0 Then AddBonus c, 15, 3, 1, "��¬������1"
                j = GetBuffCount("��¬������2")
                    If j > 0 Then AddBonus c, 10 * j, 1, 1, "��¬������2"
                j = GetBuffCount("��¬������4")
                    If j > 0 Then AddBonus c, 40, 3, 1, "��¬������4"
                j = GetBuffCount("��¬������6")
                    If j > 0 Then AddBonus c, 30, 3, 1, "��¬������6"
                
                
            
            Case "����类�"
                j = GetBuffCount("����类�����6")
                    If j > 0 And InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 298, 3, 1, "����类�����6"
                j = GetBuffCount("����类����Ƿ�ʹ��")
                    If j > 0 Then
                        c.CryoDMG = c.CryoDMG + 18
                        c.Bonustip = c.Bonustip + "��������类��츳3�ı�Ԫ���˺��ӳɣ�18%" + vbCrLf
                End If
                j = GetBuffCount("����类����츳2")
                If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddBonus c, 30, 3, 1, "����类��츳2"
            Case "�׵罫��"
                        t = (c.Energy - 100) * 0.4
                        c.ElectroDMG = c.ElectroDMG + t
                        c.Bonustip = c.Bonustip + "�����׵罫���츳2����Ԫ���˺��ӳɣ�" + CStr(t) + "%" + vbCrLf
        End Select
        
        If c.��Ե4 = True Then
                t = Round(c.Energy * 0.25, 2)
                If t > 75 Then t = 75
                Call Jug2(c, t)
                c.Bonustip = c.Bonustip + "���Ծ�Ե4���׵����ˣ�" + CStr(t) + "%" + vbCrLf
        End If
End Sub
