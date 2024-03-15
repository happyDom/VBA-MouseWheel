Attribute VB_Name = "MouseWheel"
Option Explicit

Rem ͨ��API ���Ӳ�׽��������������Ϣ

#If Win64 Then
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#Else
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#End If

Rem ���� wParam ֵ����
Private Const WH_MOUSE_LL As Long = 14

Private Const WM_MOUSEMOVE As Long = &H200      '��� [�ƶ�]

Private Const WM_LBUTTONDOWN As Long = &H201    '��� [���] ����
Private Const WM_LBUTTONUP As Long = &H202      '��� [���] ̧��
Private Const WM_LBUTTONDBLCLK As Long = &H203  '��� [���] ˫��

Private Const WM_RBUTTONDOWN As Long = &H204    '��� [�Ҽ�] ����
Private Const WM_RBUTTONUP As Long = &H205      '��� [�Ҽ�] ̧��
Private Const WM_RBUTTONDBLCLK As Long = &H206  '��� [�Ҽ�] ˫��

Private Const WM_MBUTTONDOWN As Long = &H207    '��� [����] ����
Private Const WM_MBUTTONUP As Long = &H208      '��� [����] ̧��
Private Const WM_MBUTTONDBLCLK As Long = &H209  '��� [����] ˫��
Private Const WM_MOUSEWHEEL As Long = &H20A     '��� [����] ��������ʱ��Ҫ���� lParam ������ mouseData ��Աֵ���жϹ�������

Private Const WM_KEYDOWN As Long = &H100
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28

Private hHook As Long

Private comBox As MSForms.ComboBox
Private listBox As MSForms.listBox

Rem ����һ������ṹ
Type POINT
    X As Long
    Y As Long
End Type

Rem �����������¼��ṹ
Rem ���� MSLLHOOKSTRUCT �ṹ�Ķ��壬��ο���https://blog.csdn.net/linux7985/article/details/39644669
Type MSLLHOOKSTRUCT
    pt As POINT
    mouseData As Long '��4λ��ʾ�����ƶ�����>0 ��ʾ��ǰ������<0 ��ʾ������
    flags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

Private Function MouseProcOnComBox(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MSLLHOOKSTRUCT) As Long
    Rem ��׽�������¼������� combox �� ListIndex ��ֵ
    On Error GoTo ErrLine

    If comBox Is Nothing Then
        Rem ��� ComBox ����û�и�ֵ�����˳�
        Exit Function
    End If
    If 0 = hHook Then
        Rem ��� hHook ����û��ֵ�����˳�
        Exit Function
    End If

    Dim currentIdx As Long
    currentIdx = comBox.ListIndex '��ǰλ��

    If nCode >= 0 And wParam = WM_MOUSEWHEEL Then
        If lParam.mouseData < 0 Then
            Rem ��������
            If comBox.ListCount > currentIdx Then comBox.ListIndex = currentIdx + 1
        Else
            Rem ������ǰ��
            If currentIdx > 0 Then comBox.ListIndex = currentIdx - 1
        End If
    End If

    Rem �ѹ��Ӵ��ݸ���һ�����Ƿ񴫵�ȡ���� nCode ��ֵ��
    MouseProcOnComBox = CallNextHookEx(hHook, nCode, wParam, lParam)

    Rem �˳�����
    Exit Function
ErrLine:
    Debug.Print "MouseProcOnComBox is called"
    Debug.Print "error is: "
    Debug.Print Err.description
End Function

Private Function MouseProcOnListBox(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MSLLHOOKSTRUCT) As Long
    Rem ��׽�������¼������� listbox �� ListIndex ��ֵ
    On Error GoTo ErrLine

    If listBox Is Nothing Then
        Rem ��� listBox ����û�и�ֵ�����˳�
        Exit Function
    End If
    If 0 = hHook Then
        Rem ��� hHook ����û��ֵ�����˳�
        Exit Function
    End If

    Dim currentIdx As Long
    currentIdx = listBox.ListIndex '��ǰλ��

    If nCode >= 0 And wParam = WM_MOUSEWHEEL Then
        If lParam.mouseData < 0 Then
            Rem ��������
            If listBox.ListCount > currentIdx Then listBox.ListIndex = currentIdx + 1
        Else
            Rem ������ǰ��
            If currentIdx > 0 Then listBox.ListIndex = currentIdx - 1
        End If
    End If

    Rem �ѹ��Ӵ��ݸ���һ�����Ƿ񴫵�ȡ���� nCode ��ֵ��
    MouseProcOnListBox = CallNextHookEx(hHook, nCode, wParam, lParam)

    Rem �˳�����
    Exit Function
ErrLine:
    Debug.Print "MouseProcOnListBox is called"
    Debug.Print "error is: "
    Debug.Print Err.description
End Function

Public Function ChooseHook_ComBox(ByRef Box As MSForms.ComboBox)
    Rem һ���� Enter �¼��У����йҹ��Ӳ���������
    Rem Private Sub ComboBox_Enter()
    Rem     MouseWheel.ChooseHook_Combox me.ComBox
    Rem End Sub

    If 0 <> hHook Then
        uHook
    End If

    If 0 = hHook Then
        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnComBox, 0, 0)

        If 0 <> hHook Then
            Set comBox = Box
        End If
    End If
End Function

Public Function ChooseHook_ListBox(ByRef Box As MSForms.listBox)
    Rem һ���� Enter �¼��У����йҹ��Ӳ���������
    Rem Private Sub ListBox_Enter()
    Rem     MouseWheel.ChooseHook_Combox me.ListBox
    Rem End Sub

    If 0 <> hHook Then
        uHook
    End If

    If 0 = hHook Then
        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProcOnListBox, 0, 0)

        If 0 <> hHook Then
            Set listBox = Box
        End If
    End If
End Function

Public Function uHook()
    Rem һ���� Exit �¼��У�����ȡ���Ӳ���������
    Rem Private Sub ComboBox_Exit()
    Rem     MouseWheel.uHook
    Rem End Sub

    If 0 <> hHook Then
        UnhookWindowsHookEx hHook

        Rem ��λ hHook
        hHook = 0

        Rem ��λ comBoxHook
        Set comBox = Nothing

        Rem ��λ listBox
        Set listBox = Nothing
    End If
End Function
