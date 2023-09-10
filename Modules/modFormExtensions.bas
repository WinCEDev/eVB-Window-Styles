Attribute VB_Name = "FormExtensions"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Declarations for Windows API functions.

Public Declare Function FormExtensions_SetWindowLong _
               Lib "Coredll" _
               Alias "SetWindowLongW" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long

Public Declare Function FormExtensions_GetWindowLong _
               Lib "Coredll" _
               Alias "GetWindowLongW" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long) As Long
                            
Public Declare Function FormExtensions_ShowWindow _
               Lib "Coredll" _
               Alias "ShowWindow" (ByVal hwnd As Long, _
                                   ByVal nCmdShow As Long) As Long

Public Declare Function FormExtensions_SetWindowPos _
               Lib "Coredll" _
               Alias "SetWindowPos" (ByVal hwnd As Long, _
                                     ByVal hWndInsertAfter As Long, _
                                     ByVal x As Long, _
                                     ByVal y As Long, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long, _
                                     ByVal wFlags As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Constants for getting and setting window styles and extended window styles.

Public Const FormExtensions_GWL_STYLE          As Long = (-16) 'Window Styles.

Public Const FormExtensions_GWL_EXSTYLE        As Long = (-20) 'Extended Window Styles.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Window Styles.

'The values for WS_MAXIMIZEBOX and WS_MINIMIZEBOX are reversed in Windows CE.
Public Const FormExtensions_WS_MAXIMIZEBOX     As Long = &H20000 'Normally &H10000.

Public Const FormExtensions_WS_MINIMIZEBOX     As Long = &H10000 'Normally &H20000.

Public Const FormExtensions_WS_THICKFRAME      As Long = &H40000

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Extended Window Styles.

Public Const FormExtensions_WS_EX_APPWINDOW    As Long = &H40000

Public Const FormExtensions_WS_EX_TOOLWINDOW   As Long = &H80

Public Const FormExtensions_WS_EX_CAPTIONOKBTN As Long = &H80000000

Private Const FormExtensions_WS_EX_TOPMOST     As Long = &H8

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Constants for ShowWindow function.

Public Const FormExtensions_SW_HIDE            As Long = 0

Public Const FormExtensions_SW_SHOW            As Long = 5

Public Const FormExtensions_FLAGS              As Long = 3

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Constants for SetWindowPos function.

Public Const FormExtensions_HWND_TOPMOST       As Long = -1

Public Const FormExtensions_HWND_NOTOPMOST     As Long = -2

Public Function FormExtensions_GetExStyle(ByVal Form As Form, _
                                          ByVal ExStyleFlag As Long) As Boolean

    FormExtensions_GetExStyle = (FormExtensions_GetWindowLong(Form.hwnd, FormExtensions_GWL_EXSTYLE) And ExStyleFlag) <> 0
End Function

Public Function FormExtensions_LetExStyle(ByVal Form As Form, _
                                          ByVal ExStyleFlag As Long, _
                                          ByVal State As Boolean) As Long

    Dim lngStyle As Long

    lngStyle = FormExtensions_GetWindowLong(Form.hwnd, FormExtensions_GWL_EXSTYLE)
    
    If State Then
        'Add the extended style flag.
        lngStyle = lngStyle Or ExStyleFlag
    Else
        'Remove the extended style flag.
        lngStyle = lngStyle And Not ExStyleFlag
    End If
    
    FormExtensions_LetExStyle = FormExtensions_SetWindowLong(Form.hwnd, FormExtensions_GWL_EXSTYLE, lngStyle)
End Function

Private Function FormExtensions_GetStyle(ByVal Form As Form, _
                                         ByVal StyleFlag As Long) As Boolean

    FormExtensions_GetStyle = (FormExtensions_GetWindowLong(Form.hwnd, FormExtensions_GWL_STYLE) And StyleFlag) <> 0
End Function

Private Function FormExtensions_LetStyle(ByVal Form As Form, _
                                         ByVal StyleFlag As Long, _
                                         ByVal State As Boolean) As Long

    Dim lngStyle As Long

    lngStyle = FormExtensions_GetWindowLong(Form.hwnd, FormExtensions_GWL_STYLE)
    
    If State Then
        'Add the style flag.
        lngStyle = lngStyle Or StyleFlag
    Else
        'Remove the style flag.
        lngStyle = lngStyle And Not StyleFlag
    End If
    
    FormExtensions_LetStyle = FormExtensions_SetWindowLong(Form.hwnd, FormExtensions_GWL_STYLE, lngStyle)
End Function

Public Function FormExtensions_GetToolWindow(ByVal Form As Form) As Boolean

    FormExtensions_GetToolWindow = FormExtensions_GetExStyle(Form, FormExtensions_WS_EX_TOOLWINDOW)
End Function

Public Function FormExtensions_LetToolWindow(ByVal Form As Form, _
                                             ByVal State As Boolean) As Long

    FormExtensions_ShowWindow Form.hwnd, FormExtensions_SW_HIDE
    FormExtensions_LetExStyle Form, FormExtensions_WS_EX_APPWINDOW, Not State
    FormExtensions_LetToolWindow = FormExtensions_LetExStyle(Form, FormExtensions_WS_EX_TOOLWINDOW, State)

    FormExtensions_ShowWindow Form.hwnd, FormExtensions_SW_SHOW
End Function

Public Function FormExtensions_GetOKButton(ByVal Form As Form) As Boolean

    FormExtensions_GetOKButton = FormExtensions_GetExStyle(Form, FormExtensions_WS_EX_CAPTIONOKBTN)
End Function

Public Function FormExtensions_LetOKButton(ByVal Form As Form, _
                                           ByVal State As Boolean) As Long

    FormExtensions_LetOKButton = FormExtensions_LetExStyle(Form, FormExtensions_WS_EX_CAPTIONOKBTN, State)
End Function

Public Function FormExtensions_GetMinimizeBox(ByVal Form As Form) As Boolean

    FormExtensions_GetMinimizeBox = FormExtensions_GetStyle(Form, FormExtensions_WS_MINIMIZEBOX)
End Function

Public Function FormExtensions_LetMinimizeBox(ByVal Form As Form, _
                                              ByVal State As Boolean) As Long

    FormExtensions_LetMinimizeBox = FormExtensions_LetStyle(Form, FormExtensions_WS_MINIMIZEBOX, State)
End Function

Public Function FormExtensions_GetMaximizeBox(ByVal Form As Form) As Boolean

    FormExtensions_GetMaximizeBox = FormExtensions_GetStyle(Form, FormExtensions_WS_MAXIMIZEBOX)
End Function

Public Function FormExtensions_LetMaximizeBox(ByVal Form As Form, _
                                              ByVal State As Boolean) As Long

    FormExtensions_LetMaximizeBox = FormExtensions_LetStyle(Form, FormExtensions_WS_MAXIMIZEBOX, State)
End Function

Public Function FormExtensions_GetResizable(ByVal Form As Form) As Boolean

    FormExtensions_GetResizable = FormExtensions_GetStyle(Form, FormExtensions_WS_THICKFRAME)
End Function

Public Function FormExtensions_LetResizable(ByVal Form As Form, _
                                            ByVal State As Boolean) As Long

    FormExtensions_LetResizable = FormExtensions_LetStyle(Form, FormExtensions_WS_THICKFRAME, State)
End Function

Public Function FormExtensions_GetTopMost(ByVal Form As Form) As Boolean

    Dim lngStyle As Long

    lngStyle = FormExtensions_GetWindowLong(Form.hwnd, FormExtensions_GWL_EXSTYLE)
    
    'Check if the WS_EX_TOPMOST style flag is set.
    FormExtensions_GetTopMost = (lngStyle And FormExtensions_WS_EX_TOPMOST) <> 0
End Function

Public Function FormExtensions_LetTopMost(ByVal Form As Form, ByVal State As Boolean) As Long

    If State Then

        FormExtensions_LetTopMost = FormExtensions_SetWindowPos(Form.hwnd, FormExtensions_HWND_TOPMOST, 0, 0, 0, 0, FormExtensions_FLAGS)
    Else

        FormExtensions_LetTopMost = FormExtensions_SetWindowPos(Form.hwnd, FormExtensions_HWND_NOTOPMOST, 0, 0, 0, 0, FormExtensions_FLAGS)
    End If
     
End Function



