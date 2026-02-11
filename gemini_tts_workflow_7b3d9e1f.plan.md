---
name: Implement Gemini TTS Workflow
overview: Implement a new hotkey (Win+Alt+Shift+7) that copies selected text, sends it to Gemini with a strict "repeat exactly" instruction, and triggers Gemini's Text-to-Speech feature upon completion.
todos:
  - id: refactor_read_aloud
    content: Refactor the existing logic inside the `#!+o` hotkey in `Gemini.ahk` into a standalone, reusable function named `GeminiTriggerReadAloud()`. Update `#!+o` to call this new function.
    status: completed
  - id: define_tts_class
    content: Define a new class `GeminiAsyncTTS` in `Gemini.ahk`. This class should mirror the structure of `GeminiAsyncLookup` but be tailored for the TTS workflow.
    status: completed
  - id: implement_start_method
    content: Implement the `Start()` method in `GeminiAsyncTTS`. It must copy the current selection, activate Gemini, and send the prompt.
    status: completed
  - id: implement_completion_logic
    content: Implement the `CheckCompletion()` method in `GeminiAsyncTTS` to poll for the "Stop streaming" button. Once finished, it should call `GeminiTriggerReadAloud()`.
    status: completed
  - id: update_hotkey
    content: Update the existing placeholder `#!+7` hotkey in `Gemini.ahk` to instantiate and start the `GeminiAsyncTTS` class.
    status: completed
---

# Implement Gemini TTS Workflow

## Analysis / Context
The user requires a workflow to read selected text aloud using Gemini's high-quality TTS. The current `Win+Alt+Shift+O` hotkey reads the *last* response but doesn't handle inputting new text. The new `Win+Alt+Shift+7` hotkey will bridge this by feeding selected text to Gemini and then triggering the read-aloud action.

## Proposed Changes
1.  **Refactor `#!+o`**: The logic to find the "Show more options" button and click "Text to speech" is complex and currently locked inside the `#!+o` hotkey. This needs to be extracted into `GeminiTriggerReadAloud()` so it can be reused by the new workflow without code duplication.
2.  **New Class `GeminiAsyncTTS`**: A dedicated class to handle the asynchronous nature of the request (submit -> wait -> trigger).
    *   **Prompt**: "Repeat the following text exactly as it is. Do not add any introduction, explanation, or markdown formatting. Just output the text itself:\n\n[COPIED_TEXT]"
3.  **Hotkey Assignment**: Replace the TODO in `#!+7`.

## Files to Modify
*   `c:\Users\fie7ca\Documents\scripts\Gemini.ahk`

## Implementation Strategy

1.  **Refactoring**: Move the body of `#!+o` (from `SetTitleMatchMode` down to `Send "!{Tab}"`) into a function `GeminiTriggerReadAloud()`.
2.  **Class Definition**:
    *   Create `GeminiAsyncTTS`.
    *   `Start()`: Save `OriginalHwnd`, copy text, activate Gemini, paste prompt + text, submit, restore `OriginalHwnd`, start timer.
    *   `CheckCompletion()`: Use `UIA` to check for "Stop streaming". When gone, stop timer and call `GeminiTriggerReadAloud()`.
3.  **Integration**:
    ```ahk
    #!+7:: {
        (GeminiAsyncTTS()).Start()
    }
    ```