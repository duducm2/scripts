#Requires AutoHotkey v2.0+

; Win+Alt+Shift+5 to check grammar
#!+o::
{
    A_Clipboard := ""  ; Start off empty to allow ClipWait to detect when the text has arrived.

    ; Copy content
    Send "^c"

    ClipWait  ; Wait for the clipboard to contain text.

    SetTitleMatchMode 2
    WinActivate "chatgpt"

    Sleep 250
    Send "{Esc}" ; Focus main area

    Sleep 250
    Send "+{Esc}" ; Focus main area

    ; Search String
    searchString :=
        "Below, you will find content. This content can be either a word or a sentence, in Portuguese or English. I would like you to correct the sentence based on grammar, taking cohesion and coherence into consideration. Please, don't prompt any additional comment, neither put your answer into quotation marks. Remember, we are not playing thumbs up or thumbs down now."

    ; Set the clipboard to the new text
    A_Clipboard := searchString . "`n`nContent: " . A_Clipboard

    ; Paste the content using Ctrl + V
    Sleep 100
    Send("^a") ;
    Sleep 500
    Send("^v") ; Send Ctrl + V to paste the content

    ; Press Enter to send
    Sleep 500
    Send("{Enter}")

    Sleep 500
}
