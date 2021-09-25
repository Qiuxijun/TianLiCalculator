Attribute VB_Name = "Module1"

Public CharList() As String, WeaponList() As String, CurrCharSkill() As String, CurrSkill As String, Enemy() As String, ArtList() As String, SelectCount As Integer, BuffListTip(1 To 6) As String
Public CBoxFlag As Integer, RBoxFlag As Integer
Public DMGTypetext(1 To 8) As String, ArtTypetext(1 To 18) As String
Public test As Chars
Public ReloadTip As Boolean
Public Const Text1Bound = 24
Public Const Check1Bound = 44

Public BYCTt(1 To 7) As Single, BYCTM As Integer, BYCTa(0 To 7) As Single, BYCTc As Chars, BYCTfct() As String, BYCTzct(1 To 3) As String, BYCTans As Long, BYCTnow As Long, BYCTzcta(0 To 10) As String


Public Type Chars
    cName As String
    cNumber As Integer
    cWeapon As Integer
    Level As Integer
    MaxHP As Single
    DEF As Single
    ATK As Single
    EM As Single
    CritRate As Single
    CritDmg As Single
    HealingBonus As Single
    Energy As Single
    SPower As Single
    PyroDMG As Single
    HydroDMG As Single
    DendroDMG As Single
    ElectroDMG As Single
    AnemoDMG As Single
    CryoDMG As Single
    GeoDMG As Single
    PhysicalDMG As Single
    ATKBonus As Single
    ATKFlag As Single
    HPBonus As Single
    HPFlag As Single
    DEFBonus As Single
    DEFFlag As Single
    UsedE As Integer
    UsedEA As Integer
    InSheild As Boolean
    WeaponType As Integer
    ħŮ4 As Single
    lowHP As Boolean
    ATKtip As String
    DEFtip As String
    HPtip As String
    EMtip As String
    CritRatetip As String
    CritDMGtip As String
    Bonustip As String
    ��ħ  As String
    ��Ե4 As Boolean
End Type




Public Sub CreatChar(ByRef Char As Chars, Level As Integer, WeaponLevel As Integer)
'On Error Resume Next
Dim i%, j%, sumi%, sumj%, temp() As Single, temp2() As String, t As Integer
Dim tall As String, tempc() As String, tempR() As String, tempAll() As String

Char.MaxHP = 0
Char.DEF = 0
Char.ATK = 0
Char.EM = 0
Char.CritRate = 0
Char.CritDmg = 0
Char.HealingBonus = 0
Char.Energy = 0
Char.SPower = 0
Char.PyroDMG = 0
Char.HydroDMG = 0
Char.DendroDMG = 0
Char.ElectroDMG = 0
Char.AnemoDMG = 0
Char.CryoDMG = 0
Char.GeoDMG = 0
Char.PhysicalDMG = 0
Char.ATKBonus = 0
Char.ATKFlag = 0
Char.HPBonus = 0
Char.HPFlag = 0
Char.DEFBonus = 0
Char.DEFFlag = 0
Char.ATKtip = ""
Char.DEFtip = ""
Char.HPtip = ""
Char.EMtip = ""
Char.CritRatetip = "��ɫ�Դ��ı����ʣ�5%" + vbCrLf
Char.CritDMGtip = "��ɫ�Դ��ı����˺���50%" + vbCrLf
Char.Bonustip = ""
Char.��ħ = ""
Char.UsedE = 0
Char.UsedEA = 0
Char.InSheild = False
Char.ħŮ4 = 0
Char.lowHP = False
    

For i = 1 To SelectCount
    If FrmMain.BuffCheck(i).Caption = "����ֵ����50%" Then
        If FrmMain.BuffCheck(i).Value = Checked And FrmMain.BuffCheck(i).Visible = True Then
            Char.lowHP = True
        Else
            Char.lowHP = False
        End If
    End If
Next

For i = 1 To SelectCount
    If FrmMain.BuffCheck(i).Caption = "ʩ��Ԫ��ս����" Then
        If FrmMain.BuffCheck(i).Value = Checked And FrmMain.BuffCheck(i).Visible = True Then
            Char.UsedE = 1
        Else
            Char.UsedE = 0
        End If
    End If
Next
    If CurrSkill = "c3a3" Or CurrSkill = "c3a4" Or CurrSkill = "c9a4" Or CurrSkill = "c9c2" Then test.UsedE = 1
    
    For i = 1 To SelectCount
        If Mid(FrmMain.SelectBuffLabel(i), 1, 3) = "ʥ����" And FrmMain.SelectBuffBox(i).Visible = True Then
            If FrmMain.SelectBuffBar(i).Visible = True Then
                Char.UsedE = Val(FrmMain.BuffLabel(i).tag)
            Else
                Char.UsedE = IIf(FrmMain.BuffCheck(i).Value = Checked, 1, 0)
            End If
        End If
        
         If Mid(FrmMain.SelectBuffLabel(i), 1, 4) = "�԰�֮��" And FrmMain.SelectBuffBox(i).Visible = True Then
                Char.UsedEA = Val(FrmMain.BuffLabel(i).tag)
        End If
        
        If Mid(FrmMain.SelectBuffLabel(i), 1, 5) = "��ɵ�����" And FrmMain.SelectBuffBox(i).Visible = True Then Char.InSheild = IIf(FrmMain.BuffCheck(i).Value = Checked, True, False)
    Next



        

    
                Open App.Path + "\Data\Data\C" + CStr(Char.cNumber) + "" For Binary As #1
                   tall = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   tempc = Split(tall, vbCrLf)
                   sumi = 15
                   sumj = UBound(Split(tempc(0), vbTab)) + 1
                   
                    ReDim temp(1 To sumi, 1 To sumj) As Single
                    ReDim temp2(1 To sumj)
                    
                        For i = 1 To sumi
                            tempR = Split(tempc(i - 1), vbTab)
                            For j = 1 To sumj
                                If i = 1 Then
                                    temp2(j) = tempR(j - 1)
                                Else
                                    If InStr(1, tempR(j - 1), "%") > 0 Then
                                        tempR(j - 1) = Mid(tempR(j - 1), 1, InStr(1, tempR(j - 1), "%") - 1)
                                    End If
                                    temp(i, j) = Val(tempR(j - 1))
                                End If
                            Next
                        Next
        
    
    

    

    

    Char.WeaponType = Val(CharList(Char.cNumber, 2))
    
    If Level <= 20 Then
        Char.MaxHP = temp(2, 2) + (temp(3, 2) - temp(2, 2)) / 19 * (Level - 1)
        Char.ATK = temp(2, 3) + (temp(3, 3) - temp(2, 3)) / 19 * (Level - 1)
        Char.DEF = temp(2, 4) + (temp(3, 4) - temp(2, 4)) / 19 * (Level - 1)
        Char.Level = Level
    End If
    
    If Level = 21 Then
        Char.MaxHP = temp(4, 2)
        Char.ATK = temp(4, 3)
        Char.DEF = temp(4, 4)
        Char.Level = 20
    End If
    
    If Level > 21 And Level <= 41 Then
        Char.MaxHP = temp(4, 2) + (temp(5, 2) - temp(4, 2)) / 20 * (Level - 21)
        Char.ATK = temp(4, 3) + (temp(5, 3) - temp(4, 3)) / 20 * (Level - 21)
        Char.DEF = temp(4, 4) + (temp(5, 4) - temp(4, 4)) / 20 * (Level - 21)
        Char.Level = Level - 1
    End If
    
    If Level = 42 Then
        Char.MaxHP = temp(6, 2)
        Char.ATK = temp(6, 3)
        Char.DEF = temp(6, 4)
        t = 6
        Char.Level = 40
    End If
    
    If Level > 42 And Level <= 52 Then
        Char.MaxHP = temp(6, 2) + (temp(7, 2) - temp(6, 2)) / 10 * (Level - 42)
        Char.ATK = temp(6, 3) + (temp(7, 3) - temp(6, 3)) / 10 * (Level - 42)
        Char.DEF = temp(6, 4) + (temp(7, 4) - temp(6, 4)) / 10 * (Level - 42)
        t = 7
        Char.Level = Level - 2
    End If
    
    If Level = 53 Then
        Char.MaxHP = temp(8, 2)
        Char.ATK = temp(8, 3)
        Char.DEF = temp(8, 4)
        t = 8
        Char.Level = 50
    End If
    
    If Level > 53 And Level <= 63 Then
        Char.MaxHP = temp(8, 2) + (temp(9, 2) - temp(8, 2)) / 10 * (Level - 53)
        Char.ATK = temp(8, 3) + (temp(9, 3) - temp(8, 3)) / 10 * (Level - 53)
        Char.DEF = temp(8, 4) + (temp(9, 4) - temp(8, 4)) / 10 * (Level - 53)
        t = 9
        Char.Level = Level - 3
    End If
    
    If Level = 64 Then
        Char.MaxHP = temp(10, 2)
        Char.ATK = temp(10, 3)
        Char.DEF = temp(10, 4)
        t = 10
        Char.Level = 60
    End If
    
    
    If Level > 64 And Level <= 74 Then
        Char.MaxHP = temp(10, 2) + (temp(11, 2) - temp(10, 2)) / 10 * (Level - 64)
        Char.ATK = temp(10, 3) + (temp(11, 3) - temp(10, 3)) / 10 * (Level - 64)
        Char.DEF = temp(10, 4) + (temp(11, 4) - temp(10, 4)) / 10 * (Level - 64)
        t = 11
        Char.Level = Level - 4
    End If
    
    
    If Level = 75 Then
        Char.MaxHP = temp(12, 2)
        Char.ATK = temp(12, 3)
        Char.DEF = temp(12, 4)
        t = 12
        Char.Level = 70
    End If
    
    
    If Level > 75 And Level <= 85 Then
        Char.MaxHP = temp(12, 2) + (temp(13, 2) - temp(12, 2)) / 10 * (Level - 75)
        Char.ATK = temp(12, 3) + (temp(13, 3) - temp(12, 3)) / 10 * (Level - 75)
        Char.DEF = temp(12, 4) + (temp(13, 4) - temp(12, 4)) / 10 * (Level - 75)
        t = 13
        Char.Level = Level - 5
    End If
    
    
    If Level = 86 Then
        Char.MaxHP = temp(14, 2)
        Char.ATK = temp(14, 3)
        Char.DEF = temp(14, 4)
        t = 14
        Char.Level = 80
    End If
    
    If Level > 86 And Level <= 96 Then
        Char.MaxHP = temp(14, 2) + (temp(15, 2) - temp(14, 2)) / 10 * (Level - 86)
        Char.ATK = temp(14, 3) + (temp(15, 3) - temp(14, 3)) / 10 * (Level - 86)
        Char.DEF = temp(14, 4) + (temp(15, 4) - temp(14, 4)) / 10 * (Level - 86)
        t = 15
        Char.Level = Level - 6
    End If
    

    
    If Char.Energy = 0 Then Char.Energy = 100
    
    
                Open App.Path + "\Data\Data\" + WeaponList(Char.cWeapon, 5) + "_" + WeaponList(Char.cWeapon, 4) For Binary As #1
                   tall = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   tempc = Split(tall, vbCrLf)
                   sumi = 9
                   sumj = UBound(Split(tempc(0), vbTab)) + 1
                   ReDim tempAll(1 To sumi, 1 To sumj) As String
                    For i = 1 To sumi
                        tempR = Split(tempc(i - 1), vbTab)
                        For j = 1 To sumj
                            tempAll(i, j) = tempR(j - 1)
                        Next
                    Next
        
    
    
    Char.ATK = Char.ATK + Val(tempAll(WeaponLevel + 1, 2))
    
    
    
    
    
    Char.ATKtip = "��ɫ�Ĺ�������ֵ��" + CStr(Char.ATK) + vbCrLf
    Char.DEFtip = "��ɫ�ķ�������ֵ��" + CStr(Char.DEF) + vbCrLf
    Char.HPtip = "��ɫ������ֵ��ֵ��" + CStr(Char.DEF) + vbCrLf
    
    If t >= 6 Then Jug Char, temp2(5), temp(t, 5), False
    
    If Char.CritRate = 0 Then Char.CritRate = 5
    If Char.CritDmg = 0 Then Char.CritDmg = 50
    
    
        For i = 3 To 11
            If tempAll(1, i) = WeaponList(Char.cWeapon, 6) Then
                Jug Char, tempAll(1, i), Val(tempAll(WeaponLevel + 1, i)), True
                Exit For
            End If
        Next

End Sub

Private Sub Jug(ByRef c As Chars, s As String, v As Single, Optional Weapon As Boolean)
If v < 1 Then v = v * 100
    Select Case s
        Case "������"
            c.ATKBonus = c.ATKBonus + v
            If Weapon Then
                c.ATKtip = c.ATKtip + "���������ԵĹ�������" + CStr(v) + "%" + vbCrLf
            Else
                c.ATKtip = c.ATKtip + "��ɫͻ�ƼӳɵĹ�������" + CStr(v) + "%" + vbCrLf
            End If
        Case "������"
            c.DEFBonus = c.DEFBonus + v
            If Weapon Then
                c.DEFtip = c.DEFtip + "���������Եķ�������" + CStr(v) + "%" + vbCrLf
            Else
                c.DEFtip = c.DEFtip + "��ɫͻ�Ƽӳɵķ�������" + CStr(v) + "%" + vbCrLf
            End If
        Case "����ֵ"
            c.HPBonus = c.HPBonus + v
            If Weapon Then
                c.HPtip = c.HPtip + "���������Ե�����ֵ��" + CStr(v) + "%" + vbCrLf
            Else
                c.HPtip = c.HPtip + "��ɫͻ�Ƽӳɵ�����ֵ��" + CStr(v) + "%" + vbCrLf
            End If
        Case "������"
            c.CritRate = c.CritRate + v
            If Weapon Then
                c.CritRatetip = c.CritRatetip + "���������Եı����ʣ�" + CStr(v) + "%" + vbCrLf
            Else
                c.CritRatetip = c.CritRatetip + "��ɫͻ�Ƽӳɵı����ʣ�" + CStr(v - 5) + "%" + vbCrLf
            End If
        Case "�����˺�"
            c.CritDmg = c.CritDmg + v
            If Weapon Then
                c.CritDMGtip = c.CritDMGtip + "���������Եı����˺���" + CStr(v) + "%" + vbCrLf
            Else
                c.CritDMGtip = c.CritDMGtip + "��ɫͻ�Ƽӳɵı����˺���" + CStr(v - 50) + "%" + vbCrLf
            End If
        Case "Ԫ�س���Ч��"
            c.Energy = c.Energy + v
        Case "Ԫ�ؾ�ͨ"
            c.EM = c.EM + v
            If Weapon Then
                c.EMtip = c.EMtip + "���������Ե�Ԫ�ؾ�ͨ��" + CStr(v) + "" + vbCrLf
            Else
                c.EMtip = c.EMtip + "��ɫͻ�Ƽӳɵ�Ԫ�ؾ�ͨ��" + CStr(v) + "" + vbCrLf
            End If
        Case "��Ԫ���˺��ӳ�"
            c.GeoDMG = c.GeoDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ��ң���" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ��ң���" + CStr(v) + "" + vbCrLf
            End If
        Case "��Ԫ���˺��ӳ�"
            c.PyroDMG = c.PyroDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ��𣩣�" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ��𣩣�" + CStr(v) + "" + vbCrLf
            End If
        Case "ˮԪ���˺��ӳ�"
            c.HydroDMG = c.HydroDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ�ˮ����" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ�ˮ����" + CStr(v) + "" + vbCrLf
            End If
        Case "��Ԫ���˺��ӳ�"
            c.CryoDMG = c.CryoDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ�������" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ�������" + CStr(v) + "" + vbCrLf
            End If
        Case "��Ԫ���˺��ӳ�"
            c.ElectroDMG = c.ElectroDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ��ף���" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ��ף���" + CStr(v) + "" + vbCrLf
            End If
        Case "��Ԫ���˺��ӳ�"
            c.AnemoDMG = c.AnemoDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ��磩��" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ��磩��" + CStr(v) + "" + vbCrLf
            End If
        Case "�����˺��ӳ�"
            c.PhysicalDMG = c.PhysicalDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ�������" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ�������" + CStr(v) + "" + vbCrLf
            End If
        Case "��Ԫ���˺��ӳ�"
            c.DendroDMG = c.DendroDMG + v
            If Weapon Then
                c.Bonustip = c.Bonustip + "���������Ե����ˣ��ݣ���" + CStr(v) + "" + vbCrLf
            Else
                c.Bonustip = c.Bonustip + "��ɫͻ�Ƽӳɵ����ˣ��ݣ���" + CStr(v) + "" + vbCrLf
            End If
        Case ���Ƽӳ�
            c.HealingBonus = c.HealingBonus + v
    End Select
End Sub

Private Sub GetWeaponData()
    
End Sub





Private Function AddArt1(ByRef c As Chars, Index As Integer) As Integer
Dim n%, s As String, i As Single
If Index = 0 Then
    AddArt1 = 0
    Exit Function
End If
    For n = 1 To 5
        s = ArtList(Index, n * 2)
        i = Val(ArtList(Index, n * 2 + 1))
            If s = "����ֵ%" Then c.HPBonus = c.HPBonus + i
            If s = "������%" Then c.ATKBonus = c.ATKBonus + i
            If s = "������%" Then c.DEFBonus = c.DEFBonus + i
            If s = "Ԫ�ؾ�ͨ" Then c.EM = c.EM + i
            If s = "��Ԫ���˺�%" Then c.PyroDMG = c.PyroDMG + i: c.Bonustip = c.Bonustip + "����ʥ��������ˣ��𣩣�46.6%" + vbCrLf
            If s = "ˮԪ���˺�%" Then c.HydroDMG = c.HydroDMG + i: c.Bonustip = c.Bonustip + "����ʥ��������ˣ�ˮ����46.6%" + vbCrLf
            If s = "��Ԫ���˺�%" Then c.CryoDMG = c.CryoDMG + i: c.Bonustip = c.Bonustip + "����ʥ��������ˣ�������46.6%" + vbCrLf
            If s = "��Ԫ���˺�%" Then c.AnemoDMG = c.AnemoDMG + i: c.Bonustip = c.Bonustip + "����ʥ��������ˣ��磩��46.6%" + vbCrLf
            If s = "��Ԫ���˺�%" Then c.GeoDMG = c.GeoDMG + i: c.Bonustip = c.Bonustip + "����ʥ��������ˣ��ң���46.6%" + vbCrLf
            If s = "��Ԫ���˺�%" Then c.ElectroDMG = c.ElectroDMG + i: c.Bonustip = c.Bonustip + "����ʥ��������ˣ��ף���46.6%" + vbCrLf
            If s = "�����˺�%" Then c.PhysicalDMG = c.PhysicalDMG + i: c.Bonustip = c.Bonustip + "����ʥ��������ˣ�������46.6%" + vbCrLf
            If s = "Ԫ�س���Ч��%" Then c.Energy = c.Energy + i
            If s = "������%" Then c.CritRate = c.CritRate + i
            If s = "�����˺�%" Then c.CritDmg = c.CritDmg + i
            If s = "���Ƽӳ�%" Then c.HealingBonus = c.HealingBonus + i
            If s = "����ֵ" Then c.HPFlag = c.HPFlag + i
            If s = "������" Then c.ATKFlag = c.ATKFlag + i
            If s = "������" Then c.DEFFlag = c.DEFFlag + i
    Next
    s = ArtList(Index, 1)
    AddArt1 = Val(Mid(s, 2, InStr(1, s, "_") - 2))
End Function
Public Function AddArt(ByRef c As Chars, Index As Integer, Optional Selft As String) As String
Dim t As String, i As Single, temp() As Integer, n%, j%

If IsMissing(Selft) = False And Selft <> "" Then
t = Selft
GoTo begindo
End If

    If FrmMain.SetSwitch(Index).Value = True Then
    Dim tempc As Chars
    tempc.ATKBonus = c.ATKBonus
    tempc.ATKFlag = c.ATKFlag
    tempc.HPBonus = c.HPBonus
    tempc.HPFlag = c.HPFlag
    tempc.DEFBonus = c.DEFBonus
    tempc.DEFFlag = c.DEFFlag
    tempc.EM = c.EM
    tempc.CritRate = c.CritRate
    tempc.CritDmg = c.CritDmg
    
        n = UBound(ArtTypetext)
        ReDim temp(0 To n) As Integer
        j = AddArt1(c, Val(FrmMain.SetPic1(Index).tag))
        temp(j) = temp(j) + 1
        j = AddArt1(c, Val(FrmMain.SetPic2(Index).tag))
        temp(j) = temp(j) + 1
        j = AddArt1(c, Val(FrmMain.SetPic3(Index).tag))
        temp(j) = temp(j) + 1
        j = AddArt1(c, Val(FrmMain.SetPic4(Index).tag))
        temp(j) = temp(j) + 1
        j = AddArt1(c, Val(FrmMain.SetPic5(Index).tag))
        temp(j) = temp(j) + 1

        For j = 1 To n
            If temp(j) >= 2 Then
                If temp(j) = 4 Then
                    t = ArtTypetext(j) + "4"
                Else
                    t = t + ArtTypetext(j) + "2"
                End If
            End If
        Next
    If tempc.ATKBonus < c.ATKBonus Then c.ATKtip = c.ATKtip + "����ʥ����Ĺ�����" + CStr(c.ATKBonus - tempc.ATKBonus) + "%" + vbCrLf
    If tempc.ATKFlag < c.ATKFlag Then c.ATKtip = c.ATKtip + "����ʥ����Ĺ�����" + CStr(c.ATKFlag - tempc.ATKFlag) + "" + vbCrLf
    If tempc.HPBonus < c.HPBonus Then c.HPtip = c.HPtip + "����ʥ���������ֵ��" + CStr(c.HPBonus - tempc.HPBonus) + "%" + vbCrLf
    If tempc.HPFlag < c.HPFlag Then c.HPtip = c.HPtip + "����ʥ���������ֵ��" + CStr(c.HPFlag - tempc.HPFlag) + "" + vbCrLf
    If tempc.DEFBonus < c.DEFBonus Then c.DEFtip = c.DEFtip + "����ʥ����ķ�������" + CStr(c.DEFBonus - tempc.DEFBonus) + "%" + vbCrLf
    If tempc.DEFFlag < c.DEFFlag Then c.DEFtip = c.DEFtip + "����ʥ����ķ�������" + CStr(c.DEFFlag - tempc.DEFFlag) + "" + vbCrLf
    If tempc.EM < c.EM Then c.EMtip = c.EMtip + "����ʥ�����Ԫ�ؾ�ͨ��" + CStr(c.EM - tempc.EM) + "" + vbCrLf
    If tempc.CritRate < c.CritRate Then c.CritRatetip = c.CritRatetip + "����ʥ����ı����ʣ�" + CStr(c.CritRate - tempc.CritRate) + "%" + vbCrLf
    If tempc.CritDmg < c.CritDmg Then c.CritDMGtip = c.CritDMGtip + "����ʥ����ı����˺���" + CStr(c.CritDmg - tempc.CritDmg) + "%" + vbCrLf
    
        
    Else
        With FrmMain
            c.HPFlag = c.HPFlag + Val(.SetText1(Index))
            c.HPtip = c.HPtip + "����ʥ���������ֵ��" + .SetText1(Index) + vbCrLf
            c.ATKFlag = c.ATKFlag + Val(.SetText2(Index))
            c.ATKtip = c.ATKtip + "����ʥ����Ĺ�������" + .SetText2(Index) + vbCrLf
            c.DEFFlag = c.DEFFlag + Val(.SetText3(Index))
            c.DEFtip = c.DEFtip + "����ʥ����ķ�������" + .SetText3(Index) + vbCrLf
            c.EM = c.EM + Val(.SetText4(Index))
            c.EMtip = c.EMtip + "����ʥ�����Ԫ�ؾ�ͨ��" + .SetText4(Index) + vbCrLf
            c.CritRate = c.CritRate + Val(.SetText5(Index))
            c.CritRatetip = c.CritRatetip + "����ʥ����ı����ʣ�" + .SetText5(Index) + vbCrLf
            c.CritDmg = c.CritDmg + Val(.SetText6(Index))
            c.CritDMGtip = c.CritDMGtip + "����ʥ����ı����˺���" + .SetText6(Index) + vbCrLf
            c.Energy = c.Energy + Val(.SetText7(Index))
            
            
            Select Case .SetCombo2(Index).Text
        Case "�������˺��ӳ�"
            c.GeoDMG = c.GeoDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ��ң���46.6%" + vbCrLf
        Case "�������˺��ӳ�"
            c.PyroDMG = c.PyroDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ��𣩣�46.6%" + vbCrLf
        Case "ˮ�����˺��ӳ�"
            c.HydroDMG = c.HydroDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ�ˮ����46.6%" + vbCrLf
        Case "�������˺��ӳ�"
            c.CryoDMG = c.CryoDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ�������46.6%" + vbCrLf
        Case "�������˺��ӳ�"
            c.ElectroDMG = c.ElectroDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ��ף���46.6%" + vbCrLf
        Case "�������˺��ӳ�"
            c.AnemoDMG = c.AnemoDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ��磩��46.6%" + vbCrLf
        Case "�����˺��ӳ�"
            c.PhysicalDMG = c.PhysicalDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ�������46.6%" + vbCrLf
        Case "�������˺��ӳ�"
            c.DendroDMG = c.DendroDMG + 46.6
            c.Bonustip = c.Bonustip + "����ʥ��������ˣ��ݣ���46.6%" + vbCrLf
        End Select
        
        If .SetCombo3(Index).Text = "���Ƽӳ�" Then c.HealingBonus = c.HealingBonus + 39.5
        t = .SetTipLabel13(Index).Caption
        End With
    End If
    
    AddArt = t
    
begindo:
    
            If InStr(1, t, "ˮ��") > 0 Then
                c.HydroDMG = c.HydroDMG + 15
                c.Bonustip = c.Bonustip + "����ˮ��2���׵�ˮԪ���˺��ӳɣ�15%" + vbCrLf
                If Right(t, 1) = "4" And c.UsedE > 0 And (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0) Then Call Jug2(c, 30): c.Bonustip = c.Bonustip + "����ˮ��4���׵����ˣ�30%" + vbCrLf
            End If
            
            If InStr(1, t, "����") > 0 Then
                c.ElectroDMG = c.ElectroDMG + 15
                c.Bonustip = c.Bonustip + "��������2���׵���Ԫ���˺��ӳɣ�15%" + vbCrLf
            End If
            
            If InStr(1, t, "����") > 0 Then
                If InStr(2, CurrSkill, "q") > 0 Then
                    Call Jug2(c, 20)
                    c.Bonustip = c.Bonustip + "��������2���׵����ˣ�20%" + vbCrLf
                    If Right(t, 1) = "4" Then
                        AddBonus c, 20, 1, 1, "����4����"
                    End If
                End If
            End If
            
            If InStr(1, t, "����") > 0 Then
                c.GeoDMG = c.GeoDMG + 15
                c.Bonustip = c.Bonustip + "��������2���׵���Ԫ���˺��ӳɣ�15%" + vbCrLf
            End If
        
            If InStr(1, t, "�԰�") > 0 Then
                c.PhysicalDMG = c.PhysicalDMG + 25
                c.Bonustip = c.Bonustip + "���Բ԰�2���׵������˺��ӳɣ�25%" + vbCrLf
                If Right(t, 1) = "4" And c.UsedEA > 0 Then
                    If c.UsedEA = 1 Then
                        c.ATKBonus = c.ATKBonus + 9
                        c.ATKtip = c.ATKtip + "���Բ԰�4���׵Ĺ�����9%" + vbCrLf
                    Else
                        c.PhysicalDMG = c.PhysicalDMG + 25
                        c.Bonustip = c.Bonustip + "���Բ԰�4���׵������˺��ӳɣ�25%" + vbCrLf
                        c.ATKBonus = c.ATKBonus + 18
                        c.ATKtip = c.ATKtip + "���Բ԰�4���׵Ĺ�����18%" + vbCrLf

                    End If
                End If
            End If
        
            If InStr(1, t, "����") > 0 Then
                c.AnemoDMG = c.AnemoDMG + 15
                c.Bonustip = c.Bonustip + "���Է���2���׵ķ�Ԫ���˺��ӳɣ�15%" + vbCrLf
            End If
        
            If InStr(1, t, "���") > 0 Then
                c.SPower = c.SPower + 35
                If Right(t, 1) = "4" And c.InSheild And (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0) Then Call Jug2(c, 40): c.Bonustip = c.Bonustip + "�������4���׵����ˣ�40%" + vbCrLf
            End If
            
            If InStr(1, t, "����") > 0 Then
                c.EM = c.EM + 80
                c.EMtip = c.EMtip + "��������2���׵ľ�ͨ�ӳɣ�80" + vbCrLf
                If Right(t, 1) = "4" And InStr(2, CurrSkill, "c") > 0 And c.WeaponType >= 4 Then Call Jug2(c, 35): c.Bonustip = c.Bonustip + "��������4���׵����ˣ�35%" + vbCrLf
            End If
            
            If InStr(1, t, "ƽ��") > 0 Then
                If Right(t, 1) = "4" And FrmMain.CheckState(4).Value = Checked Then Call Jug2(c, 35): c.Bonustip = c.Bonustip + "����ƽ��4���׵����ˣ�35%" + vbCrLf
            End If
            
            If InStr(1, t, "ǧ��") > 0 Then
                c.HPBonus = c.HPBonus + 20
                c.HPtip = c.HPtip + "����ǧ��2���׵�����ֵ��20%" + vbCrLf
            End If
            
            If InStr(1, t, "�ɻ�") > 0 Then
                If Right(t, 1) = "4" And FrmMain.CheckState(1).Value = Checked Then Call Jug2(c, 35): c.Bonustip = c.Bonustip + "���Զɻ�4���׵����ˣ�35%" + vbCrLf
            End If
            
            If InStr(1, t, "��Ů") > 0 Then
                c.HealingBonus = c.HealingBonus + 15
            End If
            
            If InStr(1, t, "��Ե") > 0 Then
                c.Energy = c.Energy + 20
                If Right(t, 1) = "4" And InStr(2, CurrSkill, "q") > 0 Then c.��Ե4 = True
            End If
            
            If InStr(1, t, "ħŮ") > 0 Then
                c.PyroDMG = c.PyroDMG + 15
                c.Bonustip = c.Bonustip + "����ħŮ2���׵Ļ�Ԫ���˺��ӳɣ�15%" + vbCrLf
                If Right(t, 1) = "4" Then
                    c.ħŮ4 = c.ħŮ4 + 0.15
                    If c.UsedE > 0 Then
                        n = c.UsedE
                        If n > 3 Then n = 3
                        c.PyroDMG = c.PyroDMG + 7.5 * n: c.Bonustip = c.Bonustip + "����ħŮ4���׵Ļ�Ԫ���˺��ӳɣ�" + CStr(7.5 * n) + "%" + vbCrLf

                    End If
                End If
            End If
            
            If InStr(1, t, "�Ƕ�") > 0 Then
                c.ATKBonus = c.ATKBonus + 18
                c.ATKtip = c.ATKtip + "���ԽǶ�2���׵Ĺ�����18%" + vbCrLf
                If Right(t, 1) = "4" And InStr(2, CurrSkill, "a") > 0 And c.WeaponType < 4 Then Call Jug2(c, 35): c.Bonustip = c.Bonustip + "���ԽǶ�4���׵����ˣ�35%" + vbCrLf
            End If
            
            If InStr(1, t, "׷��") > 0 Then
                c.ATKBonus = c.ATKBonus + 18
                c.ATKtip = c.ATKtip + "����׷��2���׵Ĺ�����18%" + vbCrLf
                If Right(t, 1) = "4" And c.UsedE > 0 And (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Or InStr(2, CurrSkill, "d") > 0) Then Call Jug2(c, 50): c.Bonustip = c.Bonustip + "����׷��4���׵����ˣ�50%" + vbCrLf
            End If
            
            If InStr(1, t, "����") > 0 Then
                c.CryoDMG = c.CryoDMG + 15
                c.Bonustip = c.Bonustip + "���Ա���2���׵ı�Ԫ���˺��ӳɣ�15%" + vbCrLf
                If Right(t, 1) = "4" Then
                    If FrmMain.CheckState(7).Value = Checked Then
                        c.CritRate = c.CritRate + 40
                    Else
                            If FrmMain.CheckState(3).Value = Checked Then c.CritRate = c.CritRate + 20
                    End If
                End If
            End If
            
            If InStr(1, t, "��ʿ") > 0 Then
                c.PhysicalDMG = c.PhysicalDMG + 25
                If Right(t, 1) = "4" And InStr(2, CurrSkill, "c") > 0 And GetBuffCount("ȾѪ����ʿ��") > 0 Then
                    Call Jug2(c, 50)
                    c.Bonustip = c.Bonustip + "������ʿ4���׵��ػ����ˣ�50%" + vbCrLf
                End If
            End If
            
              
End Function
Public Sub Jug2(ByRef c As Chars, v As Single, Optional ele As Boolean)
    Select Case FrmMain.Label2(2).Caption
        Case "��"
            c.GeoDMG = c.GeoDMG + v
        Case "��"
            c.PyroDMG = c.PyroDMG + v
        Case "ˮ"
            c.HydroDMG = c.HydroDMG + v
        Case "��"
            c.CryoDMG = c.CryoDMG + v
        Case "��"
            c.ElectroDMG = c.ElectroDMG + v
        Case "��"
            c.AnemoDMG = c.AnemoDMG + v
        Case "����"
            If ele <> True Then c.PhysicalDMG = c.PhysicalDMG + v
        Case "��"
            c.DendroDMG = c.DendroDMG + v
    End Select
End Sub


Public Sub AddBuffListBonus(ByRef c As Chars)
With FrmMain
    If c.cNumber = 7 And InStr(2, CurrSkill, "q") > 0 Then '�����س��Լ���Q
        .CheckState(18).Value = Checked
        .LoadBuff (4)
    End If


    c.ATKBonus = c.ATKBonus + Val(.Label2(8).tag)
    If Val(.Label2(8).tag) <> 0 Then c.ATKtip = c.ATKtip + BuffListTip(1)
    c.ATKFlag = c.ATKFlag + Val(.Label2(10).tag)
    If Val(.Label2(10).tag) <> 0 Then c.ATKtip = c.ATKtip + BuffListTip(2)
    
    Call Jug2(c, Val(.Label2(12).tag))
    If Val(.Label2(12).tag) <> 0 Then c.Bonustip = c.Bonustip + BuffListTip(3)
    
    c.EM = c.EM + Val(.Label2(14).tag)
    If Val(.Label2(14).tag) <> 0 Then c.EMtip = c.EMtip + BuffListTip(4)
    
    c.CritRate = c.CritRate + Val(.Label2(16).tag)
    If Val(.Label2(16).tag) <> 0 Then c.CritRatetip = c.CritRatetip + BuffListTip(5)
    
    c.CritDmg = c.CritDmg + Val(.Label2(18).tag)
    If Val(.Label2(18).tag) <> 0 Then c.CritDMGtip = c.CritDMGtip + BuffListTip(6)
    
    c.Energy = c.Energy + Val(.Label2(20).tag)
End With
End Sub






Public Function GetBonus(SkillCode As String) As Single
Dim j%, i As Integer, v As Variant, t As String, flag As Boolean, temp As Integer
    flag = False
    For j = 1 To UBound(CurrCharSkill)
        If CurrCharSkill(j, 1) = SkillCode Then
            i = j
            Exit For
        End If
    Next
            If InStr(2, SkillCode, "a") > 0 Then j = 0
            If InStr(2, SkillCode, "c") > 0 Then j = 0
            If InStr(2, SkillCode, "d") > 0 Then j = 0
            If InStr(2, SkillCode, "e") > 0 Then j = 1
            If InStr(2, SkillCode, "q") > 0 Then j = 2
            If SkillCode = "c9a4" Or SkillCode = "c9c2" Then j = 1
            
            
            If SkillCode = "c8s1" Then j = 1 '���ܺ���
            temp = FrmMain.LevelBox(j).ListIndex
            
            t = CurrCharSkill(i, 3 + temp)
            t = Mid(t, 1, InStr(1, t, "%") - 1)
            
            
            GetBonus = Round(Val(t), 2)

If SkillCode = "c3a3" Or SkillCode = "c3a4" Then '��������
    v = Array(137.91, 140.18, 142.45, 145.4, 147.67, 149.94, 152.89, 155.84, 158.79, 161.74, 164.7, 167.65, 170.6, 173.55, 176.5)
    GetBonus = Round(GetBonus * v(FrmMain.LevelBox(1).ListIndex - 1) / 100, 2)
End If

If SkillCode = "c15q1" And test.lowHP = True Then '���Һ���
     v = Array(379.09, 401.79, 424.49, 454, 476.7, 499.4, 528.91, 558.42, 587.93, 617.44, 646.95, 676.46, 705.97, 735.48, 764.99)
     GetBonus = Round(v(FrmMain.LevelBox(2).ListIndex - 1), 2)
End If

If SkillCode = "c15q1" And test.lowHP = True Then '���Һ���
     v = Array(379.09, 401.79, 424.49, 454, 476.7, 499.4, 528.91, 558.42, 587.93, 617.44, 646.95, 676.46, 705.97, 735.48, 764.99)
     GetBonus = Round(v(FrmMain.LevelBox(2).ListIndex - 1), 2)
End If
    


End Function

Public Function Calc(c As Chars, DMGtype As String, KX As Single, JF As Single, DJ As Integer, mode As Integer) As String
Dim finalATK As Single, finalDEF As Single, finalHP As Single, t2, t1, fCR As Single, fBonus As Single, fKX As Single, fFY As Single, Bonus As Single, EMBonus As Single, finalEM As Single
Dim ans As Single, s As String
Dim tip1 As String, tip2 As String, tip3 As String
Dim v As Variant, temp As Long

finalATK = c.ATK * (1 + c.ATKBonus / 100) + c.ATKFlag
finalDEF = c.DEF * (1 + c.DEFBonus / 100) + c.DEFFlag
finalHP = c.MaxHP * (1 + c.HPBonus / 100) + c.HPFlag




tip1 = "���չ�������"


tip2 = "���������˺���"
tip3 = ""

'---------����---------
If CurrSkill = "c1e2" Then
finalATK = finalDEF
tip1 = "���շ�������"
c.ATKtip = c.DEFtip

End If





    Select Case DMGtype
        Case "��"
            fBonus = c.GeoDMG / 100 + 1
        Case "��"
            fBonus = c.PyroDMG / 100 + 1
            If FrmMain.CheckState(2).Value = Checked Then
                EMBonus = 1.5
                tip3 = "�����������ˮ��"
                If FrmMain.CheckState(39).Value = Checked Then c.ħŮ4 = c.ħŮ4 + 0.15
            End If
            If FrmMain.CheckState(3).Value = Checked Or FrmMain.CheckState(7).Value = Checked Then
                EMBonus = 2
                tip3 = "���ڻ���������"
            End If

        Case "ˮ"
            fBonus = c.HydroDMG / 100 + 1
            If FrmMain.CheckState(1).Value = Checked Then
                EMBonus = 2
                tip3 = "��������ˮ���"
                If FrmMain.CheckState(39).Value = Checked Then c.ħŮ4 = c.ħŮ4 + 0.15
            End If
            
        Case "��"
            fBonus = c.CryoDMG / 100 + 1
            If FrmMain.CheckState(1).Value = Checked Then
                EMBonus = 1.5
                tip3 = "���ڻ��������"
            End If
            
        Case "��"
            fBonus = c.ElectroDMG / 100 + 1
        Case "��"
            fBonus = c.AnemoDMG / 100 + 1
        Case "����"
            fBonus = c.PhysicalDMG / 100 + 1
        Case "��"
            fBonus = c.DendroDMG / 100 + 1
    End Select
    



t1 = c.CritRate / 100
t2 = c.CritDmg / 100
If t1 > 1 Then t1 = 1
fCR = 1 + t1 * t2


Bonus = GetBonus(CurrSkill) / 100



    ans = finalATK * Bonus '�����Ĺ�����*����
    tip2 = tip2 + CStr(Int(ans))
    
    If c.cNumber = 1 And CBoxFlag >= 1 And InStr(2, CurrSkill, "q") > 0 Then '������2��
        ans = ans + finalDEF * 1.2
        tip2 = tip2 + "+" + CStr(Int(finalDEF * 1.2)) + "��������2����"
    End If
    
    If c.cNumber = 5 And CBoxFlag >= 1 And InStr(2, CurrSkill, "e") > 0 Then '����2��
        tip2 = tip2 + "+" + CStr(Int(ans * 2)) + "������2����"
        ans = ans * 3
    End If
    
    If c.cNumber = 13 And InStr(2, CurrSkill, "q") > 0 Then
        If CurrSkill = "c13q1" Then
            v = Array(3.89, 4.18, 4.47, 4.86, 5.15, 5.44, 5.83, 6.22, 6.61, 7, 7.39, 7.78, 8.26, 8.75, 9.23)
        Else
            v = Array(0.73, 0.78, 0.84, 0.91, 0.96, 1.02, 1.09, 1.16, 1.23, 1.31, 1.38, 1.45, 1.54, 1.63, 1.72)
        End If
        temp = GetBuffCount("�׵罫������Ը")
        If temp > 0 Then
            temp = temp * v(FrmMain.LevelBox(2).ListIndex - 1) * finalATK / 100
            ans = ans + temp
            tip2 = tip2 + "+" + CStr(Int(temp)) + "���׵罫��Ը���ӳɣ�"
        End If

    End If
    If c.cNumber = 13 And CBoxFlag >= 1 Then JF = JF + 60
    
    
    
    If KX <= 0 Then fKX = 1 - KX / 200
    If KX > 0 And KX <= 75 Then fKX = 1 - KX / 100
    If KX > 75 Then fKX = 1 / (1 + 4 * KX / 100)
    

    
    fFY = (c.Level + 100) / ((1 - JF / 100) * (DJ + 100) + c.Level + 100)
    
    
    If EMBonus <> 0 Then '��������Ӧ
        finalEM = c.EM * 2.78 / (c.EM + 1400)
        ans = ans * EMBonus * (1 + finalEM + c.ħŮ4)
    End If
    
    ans = ans * fBonus '����������

    ans = ans * fKX * fFY '���Կ������ͷ�����

    

    If mode = 1 Then
c.ATKtip = "-" + Replace(c.ATKtip, vbCrLf, vbCrLf + "-")
c.ATKtip = Mid(c.ATKtip, 1, Len(c.ATKtip) - 1)
c.DEFtip = "-" + Replace(c.DEFtip, vbCrLf, vbCrLf + "-")
c.DEFtip = Mid(c.DEFtip, 1, Len(c.DEFtip) - 1)
c.HPtip = "-" + Replace(c.HPtip, vbCrLf, vbCrLf + "-")
c.HPtip = Mid(c.HPtip, 1, Len(c.HPtip) - 1)
c.Bonustip = "-" + Replace(c.Bonustip, vbCrLf, vbCrLf + "-")
c.Bonustip = Mid(c.Bonustip, 1, Len(c.Bonustip) - 1)
c.CritRatetip = "-" + Replace(c.CritRatetip, vbCrLf, vbCrLf + "-")
c.CritRatetip = Mid(c.CritRatetip, 1, Len(c.CritRatetip) - 1)
c.CritDMGtip = "-" + Replace(c.CritDMGtip, vbCrLf, vbCrLf + "-")
c.CritDMGtip = Mid(c.CritDMGtip, 1, Len(c.CritDMGtip) - 1)
c.EMtip = "-" + Replace(c.EMtip, vbCrLf, vbCrLf + "-")
c.EMtip = Mid(c.EMtip, 1, Len(c.EMtip) - 1)

    
                        s = "�˺�����ѧ������" + CStr(Round(ans * fCR)) + tip3 + vbCrLf + "�˺��ı������֣�" + CStr(Round(ans * (1 + t2))) + tip3 + vbCrLf
                        s = s + vbCrLf + tip1 + CStr(finalATK)
                        s = s + vbCrLf + c.ATKtip
                        s = s + vbCrLf + "���ʣ�" + CStr(Bonus * 100) + "%" + vbCrLf + tip2 + vbCrLf
                        
                        If EMBonus <> 0 Then
                            s = s + vbCrLf + "Ԫ�ؾ�ͨ����" + CStr(c.EM)
                            s = s + vbCrLf + "Ԫ�ؾ�ͨ�ӳɣ�" + CStr(Round(finalEM * 100, 2)) + "%��" + "������Ӧ�ӳ�" + CStr(c.ħŮ4 * 100) + "%"
                            s = s + vbCrLf + c.EMtip + vbCrLf
                        End If

                        s = s + vbCrLf + "���˹���" + CStr((fBonus - 1) * 100) + "%"
                        s = s + vbCrLf + c.Bonustip + vbCrLf
                        
                        s = s + vbCrLf + "���￹�ԣ�" + DMGtype + "����" + CStr(KX)
                        s = s + vbCrLf + "���������" + CStr(JF) + vbCrLf
                        
                        
                        s = s + vbCrLf + "�����ʣ�" + CStr(t1 * 100) + "%"
                        s = s + vbCrLf + c.CritRatetip
                        s = s + vbCrLf + "�����˺���" + CStr(t2 * 100) + "%"
                        s = s + vbCrLf + c.CritDMGtip
                        
                        
                        Calc = s
    
    
        ans = ans * fCR
    Else
        ans = ans * fCR
        Calc = CStr(Round(ans))
    End If
 


'Calc = Round(ans)
End Function





