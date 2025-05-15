#Requires AutoHotkey v2

CapsLock & 6::
{

    A_Clipboard := ""  ; Start off empty to allow ClipWait to detect when the text has arrived.

    ; Copy content
    Send "^c"

    ClipWait  ; Wait for the clipboard to contain text.

    SetTitleMatchMode 2
    WinActivate "chatgpt - transcription"

    Sleep 250
    Send "{Esc}" ; Focus main area

    Sleep 250
    Send "+{Esc}" ; Focus main area

    ; Search String
    searchString :=
        "Below, you will find a word or phrase. I'd like you to answer in five sections: the 1st section you will repeat the word twice. For each time you repeat, use a point to finish the phrase. The 2nd section should have the definition of the word (You should also say each part of speech does the different definitions belong to). The 3d section should have the pronunciation of this word using the Internation Phonetic Alphabet characters (for American English).The 4th section should have the same word applied in a real sentence (put that in quotations, so I can identify that). In the 5th, Write down the translation of the word into Portuguese. Please, do not title any section. Thanks!"

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

    ; Back to previous window
    Send "!{Tab}"

    Send("{CapsLock}")
}
