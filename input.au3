#include <Array.au3>

HotKeySet("^!t", "ProcessClipboard") ; Register Ctrl+Alt+T as the hotkey

While 1
    Sleep(100)
WEnd

Func ProcessClipboard()
    ; Read the tab-separated values from the clipboard
    Local $clipboardData = ClipGet()

    ; Check if clipboard has data
    If StringLen($clipboardData) > 0 Then
        ; Split the string into an array using tabs as the delimiter
        Local $valuesArray = StringSplit($clipboardData, @TAB)

        ; Loop through the array and send each value with a down arrow press
        For $i = 1 To $valuesArray[0]
            ; Send the current value
            Send($valuesArray[$i])

            ; Send a down arrow press (you can adjust the sleep duration if needed)
            Send("{DOWN}")
            Sleep(50) ; Sleep for 1 second (adjust as needed)
        Next
    Else
        MsgBox(16, "Error", "Clipboard is empty or does not contain tab-separated values.")
    EndIf
EndFunc

