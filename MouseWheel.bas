Attribute VB_Name = "MouseWheel"
Option Explicit

Rem 通过API 钩子捕捉并处理鼠标滚轮消息

#If Win64 Then
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#Else
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#End If

Rem 定义 wParam 值常量
Private Const WH_MOUSE_LL As Long = 14

Private Const WM_MOUSEMOVE As Long = &H200      '鼠标 [移动]

Private Const WM_LBUTTONDOWN As Long = &H201    '鼠标 [左键] 按下
Private Const WM_LBUTTONUP As Long = &H202      '鼠标 [左键] 抬起
Private Const WM_LBUTTONDBLCLK As Long = &H203  '鼠标 [左键] 双击

Private Const WM_RBUTTONDOWN As Long = &H204    '鼠标 [右键] 按下
Private Const WM_RBUTTONUP As Long = &H205      '鼠标 [右键] 抬起
Private Const WM_RBUTTONDBLCLK As Long = &H206  '鼠标 [右键] 双击

Private Const WM_MBUTTONDOWN As Long = &H207    '鼠标 [滚轮] 按下
Private Const WM_MBUTTONUP As Long = &H208      '鼠标 [滚轮] 抬起
Private Const WM_MBUTTONDBLCLK As Long = &H209  '鼠标 [滚轮] 双击
Private Const WM_MOUSEWHEEL As Long = &H20A     '鼠标 [滚轮] 滚动，此时需要根据 lParam 参数的 mouseData 成员值来判断滚动方向

Private Const WM_KEYDOWN As Long = &H100
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28

Private hHook As Long

Private comBox As MSForms.ComboBox
Private listBox As MSForms.listBox

Rem 定义一个坐标结构
Type POINT
    X As Long
    Y As Long
End Type

Rem 定义鼠标滚轮事件结构
Rem 关于 MSLLHOOKSTRUCT 结构的定义，请参考：https://blog.csdn.net/linux7985/article/details/39644669
Type MSLLHOOKSTRUCT
    pt As POINT
    mouseData As Long '高4位表示滚轮移动量：>0 表示向前滚动，<0 表示向后滚动
    flags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

Private Function MouseProcOnComBox(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MSLLHOOKSTRUCT) As Long
    Rem 捕捉鼠标滚轮事件，调整 combox 中 ListIndex 的值
    On Error GoTo ErrLine

    If comBox Is Nothing Then
        Rem 如果 ComBox 对象没有赋值，则退出
        Exit Function
    End If
    If 0 = hHook Then
        Rem 如果 hHook 钩子没有值，则退出
        Exit Function
    End If

    Dim currentIdx As Long
    currentIdx = comBox.ListIndex '当前位置

    If nCode >= 0 And wParam = WM_MOUSEWHEEL Then
        If lParam.mouseData < 0 Then
            Rem 滚轮向后滚
            If comBox.ListCount > currentIdx Then comBox.ListIndex = currentIdx + 1
        Else
            Rem 滚轮向前滚
            If currentIdx > 0 Then comBox.ListIndex = currentIdx - 1
        End If
    End If

    Rem 把钩子传递给下一级（是否传递取决于 nCode 的值）
    MouseProcOnComBox = CallNextHookEx(hHook, nCode, wParam, lParam)

    Rem 退出程序
    Exit Function
ErrLine:
    Debug.Print "MouseProcOnComBox is called"
    Debug.Print "error is: "
    Debug.Print Err.description
End Function

Private Function MouseProcOnListBox(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MSLLHOOKSTRUCT) As Long
    Rem 捕捉鼠标滚轮事件，调整 listbox 中 ListIndex 的值
    On Error GoTo ErrLine

    If listBox Is Nothing Then
        Rem 如果 listBox 对象没有赋值，则退出
        Exit Function
    End If
    If 0 = hHook Then
        Rem 如果 hHook 钩子没有值，则退出
        Exit Function
    End If

    Dim currentIdx As Long
    currentIdx = listBox.ListIndex '当前位置

    If nCode >= 0 And wParam = WM_MOUSEWHEEL Then
        If lParam.mouseData < 0 Then
            Rem 滚轮向后滚
            If listBox.ListCount > currentIdx Then listBox.ListIndex = currentIdx + 1
        Else
            Rem 滚轮向前滚
            If currentIdx > 0 Then listBox.ListIndex = currentIdx - 1
        End If
    End If

    Rem 把钩子传递给下一级（是否传递取决于 nCode 的值）
    MouseProcOnListBox = CallNextHookEx(hHook, nCode, wParam, lParam)

    Rem 退出程序
    Exit Function
ErrLine:
    Debug.Print "MouseProcOnListBox is called"
    Debug.Print "error is: "
    Debug.Print Err.description
End Function

Public Function ChooseHook_ComBox(ByRef Box As MSForms.ComboBox)
    Rem 一般在 Enter 事件中，进行挂钩子操作，例如
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
    Rem 一般在 Enter 事件中，进行挂钩子操作，例如
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
    Rem 一般在 Exit 事件中，进行取钩子操作，例如
    Rem Private Sub ComboBox_Exit()
    Rem     MouseWheel.uHook
    Rem End Sub

    If 0 <> hHook Then
        UnhookWindowsHookEx hHook

        Rem 复位 hHook
        hHook = 0

        Rem 复位 comBoxHook
        Set comBox = Nothing

        Rem 复位 listBox
        Set listBox = Nothing
    End If
End Function
