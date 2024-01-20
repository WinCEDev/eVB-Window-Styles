# eVB-Window-Styles

This module lets you set various styles and other window properties in your eVB applications.

## Usage

Add the `FormExtensions.bas` module to your project, then use any of the included functions. For example, if you want to add a maximize and minimize button to your form, add the following code:

```vb
Private Sub Form_Load()
    FormExtensions_LetMaximizeBox Me, True
    FormExtensions_LetMinimizeBox Me, True
End Sub
```

## Screenshots

![A screenshot of the example application, showing all options as inactive.](https://github.com/WinCEDev/eVB-Window-Styles/blob/main/Screenshots/app.png?raw=1)

![A screenshot of the example application, with the Minimize, Maximize and Resizable options selected.](https://github.com/WinCEDev/eVB-Window-Styles/blob/main/Screenshots/app_minmaxresize.png?raw=1)

## Functions

### `FormExtensions_GetExStyle`

```vb
Public Function FormExtensions_GetExStyle(ByVal Form As Form, _
                                          ByVal ExStyleFlag As Long) As Boolean
```

This function retrieves a specific extended window style (`ExStyleFlag`) for the given `Form` and returns `True` if the style is set and `False` if it is not.

### `FormExtensions_LetExStyle`

```vb
Public Function FormExtensions_LetExStyle(ByVal Form As Form, _
                                          ByVal ExStyleFlag As Long, _
                                          ByVal State As Boolean) As Long
```

This function allows you to set or clear a specific extended window style (`ExStyleFlag`) for the given `Form` based on the provided State (`True` for set, `False` for clear). It returns the new extended style after modification.

### `FormExtensions_GetStyle`

```vb
Private Function FormExtensions_GetStyle(ByVal Form As Form, _
                                         ByVal StyleFlag As Long) As Boolean
```

This function retrieves a specific standard window style (`StyleFlag`) for the given `Form` and returns `True` if the style is set and `False` if it is not.

### `FormExtensions_LetStyle`

```vb
Private Function FormExtensions_LetStyle(ByVal Form As Form, _
                                         ByVal StyleFlag As Long, _
                                         ByVal State As Boolean) As Long
```

This function allows you to set or clear a specific standard window style (`StyleFlag`) for the given `Form` based on the provided `State` (`True` for set, `False` for clear). It returns the new style after modification.

### `FormExtensions_GetToolWindow`

```vb
Public Function FormExtensions_GetToolWindow(ByVal Form As Form) As Boolean
```

This function checks whether the `Form` has the tool window extended style (`WS_EX_TOOLWINDOW`) set and returns `True` if it is set.

### `FormExtensions_LetToolWindow`

```vb
Public Function FormExtensions_LetToolWindow(ByVal Form As Form, _
                                             ByVal State As Boolean) As Long
```

This function allows you to set or clear the tool window extended style for the given `Form` based on the provided `State` (`True` for set, `False` for clear). Requires the form to be redisplayed.

### `FormExtensions_GetOKButton`

```vb
Public Function FormExtensions_GetOKButton(ByVal Form As Form) As Boolean
```

This function checks whether the `Form` has the caption OK button extended style (`WS_EX_CAPTIONOKBTN`) set and returns `True` if it is set.

### `FormExtensions_LetOKButton`

```vb
Public Function FormExtensions_LetOKButton(ByVal Form As Form, _
                                           ByVal State As Boolean) As Long
```

This function allows you to set or clear the caption OK button extended style for the given `Form` based on the provided `State` (`True` for set, `False` for clear).

**Note:** You will need an external component to subclass the form if you want to handle events when the user taps on the OK button.

### `FormExtensions_GetMinimizeBox`

```vb
Public Function FormExtensions_GetMinimizeBox(ByVal Form As Form) As Boolean
```

This function checks whether the `Form` has the minimize box style (`WS_MINIMIZEBOX`) set and returns `True` if it is set.

### `FormExtensions_LetMinimizeBox`

```vb
Public Function FormExtensions_LetMinimizeBox(ByVal Form As Form, _
                                              ByVal State As Boolean) As Long
```

This function allows you to set or clear the minimize box style for the given `Form` based on the provided `State` (`True` for set, `False` for clear).

### `FormExtensions_GetMaximizeBox`

```vb
Public Function FormExtensions_GetMaximizeBox(ByVal Form As Form) As Boolean
```

This function checks whether the `Form` has the maximize box style (`WS_MAXIMIZEBOX`) set and returns `True` if it is set.

### `FormExtensions_LetMaximizeBox`

```vb
Public Function FormExtensions_LetMaximizeBox(ByVal Form As Form, _
                                              ByVal State As Boolean) As Long
```

This function allows you to set or clear the maximize box style for the given `Form` based on the provided State (`True` for set, `False` for clear).

### `FormExtensions_GetResizable`

```vb
Public Function FormExtensions_GetResizable(ByVal Form As Form) As Boolean
```

This function checks whether the `Form` has the thick frame style (`WS_THICKFRAME`) set and returns `True` if it is set. This indicates whether the `Form` is resizable.

### `FormExtensions_LetResizable`

```vb
Public Function FormExtensions_LetResizable(ByVal Form As Form, _
                                            ByVal State As Boolean) As Long
```

This function allows you to set or clear the thick frame style for the given `Form` based on the provided `State` (`True` for set, `False` for clear). This can control whether the `Form` is resizable.

On some devices the resize area might be hard to hit with the stylus, so it's recommended to also provide an alternative way to resize the form (like a dialog box).

### `FormExtensions_GetTopMost`

```vb
Public Function FormExtensions_GetTopMost(ByVal Form As Form) As Boolean
```

This function checks whether the `Form` has the topmost extended style (`WS_EX_TOPMOST`) set and returns `True` if it is set.

### `FormExtensions_LetTopMost`

```vb
Public Function FormExtensions_LetTopMost(ByVal Form As Form, ByVal State As Boolean) As Long
```

This function allows you to set or clear the topmost attribute for the given `Form` based on the provided `State` (`True` for set, `False` for clear).

## Notes
- It appears that the values of `WS_MAXIMIZEBOX` and `WS_MINIMIZEBOX` are reversed in Windows CE. The constants have been updated to reflect this.
