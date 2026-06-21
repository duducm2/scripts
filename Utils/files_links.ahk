; =============================================================================
; Utils module: files_links.ahk
; Files and links quick-open (InitQuickOpenFiles)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Files & Links System
; =============================================================================

; Global variables for quick open files
global g_QuickOpenFiles := []
global g_QuickOpenFileCharMap := Map()  ; Maps character to file path

; Register a file for quick opening
RegisterQuickOpenFile(filePath, title) {
    global g_QuickOpenFiles
    g_QuickOpenFiles.Push({ filePath: filePath, title: title, category: "Files & Links" })
}

; Initialize quick open files
InitQuickOpenFiles() {
    ; Register dissertation Power BI file with character 'y'
    RegisterQuickOpenFile(
        "C:\Users\eduev\Meu Drive\04 - Pós-graduação\01 - Mestrado\26-ai-experiment\infoVis\Dissertation InfoVis  - PowerBI - Charts.pbix",
        "📊 Dissertation InfoVis"
    )

    ; Register radio-tiso exercises YouTube link
    RegisterQuickOpenFile(
        "https://www.youtube.com/watch?v=I6ZRH9Mraqw&t=2s",
        "📻 Radio-Tiso Exercises"
    )

    ; Register GS_UX core team_UX and CIP Integration Miro
    RegisterQuickOpenFile(
        "https://miro.com/app/board/uXjVJdbNFkA=/",
        "🎨 GS_UX core team_UX and CIP Integration Miro"
    )

    ; Register GS_E&S_CIP Dashboard research and design Miro
    RegisterQuickOpenFile(
        "https://miro.com/app/board/uXjVJVZSXvk=/",
        "📊 GS_E&S_CIP Dashboard research and design Miro"
    )
}
InitQuickOpenFiles()

