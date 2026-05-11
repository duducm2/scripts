# AutoHotkey v2 Error: This local variable has not been assigned a value

## Error Message

```
Error: This local variable has not been assigned a value.
A global declaration inside the function may be required.

Specifically: gui
```

## Cause
- In AutoHotkey v2, all local variables must be declared at the very top of the function, before any code, comments, or blank lines.
- If you use or assign a variable (like `gui := Gui(...)`) after any other statement (including `Loop`, `try`, `Hotkey`, or even a comment), AHK v2 may throw this error.
- Declaring a variable as `global` and then assigning a class instance (like `Gui`) to it can cause a different error: "This Class cannot be used as an output variable."

## Solutions

### 1. Declare Local Variables First
Place all `local` declarations as the very first lines in the function, before any code, comments, or blank lines:

```ahk2
MyFunction() {
    local gui
    gui := Gui()
    ; ...rest of code...
}
```

### 2. Avoid Global for Class Instances
Do not declare a variable as `global` if you intend to assign a class instance (like `Gui`) to it. Use local scope instead.

### 3. Assign Before Any Other Statement
If you do not declare the variable as local, assign it as the very first statement in the function, before any other code or hotkey cleanup.

### 4. Example of Correct Usage
```ahk2
MyFunction() {
    local gui
    gui := Gui()
    ; ...
}
```

### 5. Reference
- [AutoHotkey v2 Documentation: Local Variables](https://www.autohotkey.com/docs/v2/Functions.htm#Local)
- [AutoHotkey v2 Forum: Local variable assignment error](https://www.autohotkey.com/boards/viewtopic.php?t=110856)
- [GitHub Issue: AHK v2 local variable assignment](https://github.com/AutoHotkey/AutoHotkey/issues/246)

## Case Study: StudyLink_StudySubmenu (Utils.ahk)

### Problem
In the `StudyLink_StudySubmenu` function, we encountered the error:

```
Error: This local variable has not been assigned a value.
A global declaration inside the function may be required.

Specifically: gui
```

This occurred because the function originally declared `local gui` but then used a global GUI variable pattern elsewhere in the codebase, leading to confusion and a #Warn warning.

### Solution
- We removed the unused `local gui` declaration from `StudyLink_StudySubmenu`.
- The function now exclusively uses the persistent global variable `g_StudyLinkSubmenuGui` for the submenu GUI, following the robust pattern used for all persistent overlays in the codebase.
- This fully suppresses the error and the #Warn warning, ensuring the user never sees any unnecessary alerts.

#### Final Pattern Used
```ahk2
StudyLink_StudySubmenu(parentGui, topic, studyKey) {
    global g_StudyLinkSubmenuGui
    parentGui.Destroy()
    if (IsObject(g_StudyLinkSubmenuGui) && g_StudyLinkSubmenuGui.Hwnd)
        try g_StudyLinkSubmenuGui.Destroy()
    g_StudyLinkSubmenuGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    ; ...rest of code...
}
```

### Takeaway
- For persistent GUIs in AHK v2, use a global variable and remove any unused local declarations to avoid assignment errors and warnings.
- This approach is robust, warning-free, and matches the codebase's established best practices.

## Case Study: Unassigned Global Variable Error (g_StudyLinkSubmenuGui)

### Error Message
```
Error: This global variable has not been assigned a value.

Specifically: g_StudyLinkSubmenuGui
```

### Context
- This error occurred in `Utils.ahk` at the line:
  ```ahk2
  if (IsObject(g_StudyLinkSubmenuGui) && g_StudyLinkSubmenuGui.Hwnd)
      try g_StudyLinkSubmenuGui.Destroy()
  ```
- The function `StudyLink_StudySubmenu` uses a persistent global variable `g_StudyLinkSubmenuGui` to manage the submenu GUI.

### Hypothesis
- In AutoHotkey v2, referencing a global variable that has never been assigned (even in `IsObject()`) throws a runtime error.
- If `g_StudyLinkSubmenuGui` is not initialized at the top level of the script, the first call to `IsObject(g_StudyLinkSubmenuGui)` will fail.
- This can happen if the script is reloaded, or if the function is called before any assignment to `g_StudyLinkSubmenuGui` has occurred.

### Solution
- Explicitly initialize `g_StudyLinkSubmenuGui` at the top level of the script, outside any function, near other global declarations:
  ```ahk2
  global g_StudyLinkSubmenuGui := ""
  ```
- This ensures the variable always exists (as an empty string) before any function references it. `IsObject("")` safely returns false, so the check and subsequent logic work as intended.

### Final Pattern Used
```ahk2
; At the top of Utils.ahk
global g_StudyLinkSubmenuGui := ""

; ...later in StudyLink_StudySubmenu...
if (IsObject(g_StudyLinkSubmenuGui) && g_StudyLinkSubmenuGui.Hwnd)
    try g_StudyLinkSubmenuGui.Destroy()
```

### Takeaway
- Always initialize persistent global variables at the top level to avoid unassigned variable errors in AHK v2.
- This approach is robust and ensures all references to the variable are safe, even before the first assignment.

## Case Study: Gui Has No Window Error (g_StudyLinkSubmenuGui)

### Error Message
```
Error: Gui has no window.

Specifically: g_StudyLinkSubmenuGui.Destroy()
```

### Context
- This error occurred in `Utils.ahk` at the line:
  ```ahk2
  if (IsObject(g_StudyLinkSubmenuGui) && g_StudyLinkSubmenuGui.Hwnd)
      try g_StudyLinkSubmenuGui.Destroy()
  ```
- The user was able to define a link, but when the edit field appeared, the modal banner was not hidden. After trying to access the same study again, this error was triggered.

### Hypothesis
- After destroying or closing a GUI, the variable `g_StudyLinkSubmenuGui` still points to a Gui object, but its window handle (`Hwnd`) is invalid or missing.
- Calling `.Destroy()` on a Gui object that no longer has a window results in this error.
- This can happen if the GUI is destroyed elsewhere (e.g., by closing the edit field/modal) but the global variable is not reset or re-initialized.

### Solution
- After destroying the GUI, set `g_StudyLinkSubmenuGui := ""` to ensure the variable does not reference a stale Gui object.
- Before calling `.Destroy()`, check both `IsObject(g_StudyLinkSubmenuGui)` and that `g_StudyLinkSubmenuGui.Hwnd` is valid (nonzero).
- Optionally, wrap the `.Destroy()` call in a try/catch block to suppress errors if the window is already gone.

### Final Pattern Used
```ahk2
if (IsObject(g_StudyLinkSubmenuGui) && g_StudyLinkSubmenuGui.Hwnd) {
    try g_StudyLinkSubmenuGui.Destroy()
    g_StudyLinkSubmenuGui := ""
}
```

### Takeaway
- Always reset global GUI variables after destroying their windows to avoid referencing invalid objects.
- This prevents errors when re-opening or reusing modals in AHK v2.

## Summary
- Always declare local variables at the very top of the function in AHK v2.
- Do not use `global` for class instance assignments.
- Assign variables before any other code, comments, or blank lines to avoid this error.
