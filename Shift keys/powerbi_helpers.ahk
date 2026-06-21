; =============================================================================
; Shift keys module: powerbi_helpers.ahk
; Power BI drawer config helpers
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

PowerBI_GetDrawerConfigs() {
    return [{
        label: "Visualizations",
        names: ["Visualizations", "Visualizações"],
        classContains: ["toggle-button"]
    }, {
        label: "Data",
        names: ["Data", "Dados"],
        classContains: ["toggle-button"]
    }, {
        label: "Properties",
        names: ["Properties", "Propriedades"],
        classContains: ["toggle-button"]
    }, {
        label: "Filter pane",
        names: [
            "Collapse or expand the filter pane while editing. This also determines how report readers see it",
            "Filter pane",
            "Pane de filtros"
        ],
        classContains: ["pbi-glyph-doublechevronleft", "pbi-glyph-doublechevronright"]
    }]
}

PowerBI_FindDrawerButton(root, config) {
    try {
        if config.HasOwnProp("names") {
            for , name in config.names {
                if !name
                    continue
                for typeVariant in ["Button", 50000] {
                    btn := root.FindFirst({ Type: typeVariant, Name: name })
                    if btn
                        return btn
                    btn := root.FindFirst({ Type: typeVariant, Name: name, matchmode: "Substring" })
                    if btn
                        return btn
                }
            }
        }

        if config.HasOwnProp("classNames") {
            for , className in config.classNames {
                if !className
                    continue
                for typeVariant in ["Button", 50000] {
                    btn := root.FindFirst({ Type: typeVariant, ClassName: className })
                    if btn
                        return btn
                }
            }
        }

        if config.HasOwnProp("classContains") {
            classNeedles := config.classContains
            if (Type(classNeedles) != "Array")
                classNeedles := [classNeedles]
            allButtons := ""
            try allButtons := root.FindAll({ Type: "Button" })
            if !allButtons
                try allButtons := root.FindAll({ Type: 50000 })
            if allButtons {
                for btn in allButtons {
                    if !btn
                        continue
                    btnClass := ""
                    try btnClass := btn.ClassName
                    for , needle in classNeedles {
                        if needle && InStr(btnClass, needle)
                            return btn
                    }
                }
            }
        }
    } catch Error {
    }
    return ""
}

PowerBI_CollapseDrawerElement(element) {
    current := element
    loop 4 {
        if !current
            break
        result := PowerBI_AttemptCollapse(current)
        if result != -1
            return result
        try current := UIA.TreeWalkerTrue.GetParentElement(current)
        catch {
            current := ""
        }
    }
    return -1
}

PowerBI_AttemptCollapse(element) {
    try {
        hasPattern := element.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
        if hasPattern {
            pat := element.ExpandCollapsePattern
            state := pat.ExpandCollapseState
            if state != UIA.ExpandCollapseState.Collapsed {
                pat.Collapse()
                Sleep 40
                return 1
            }
            return 0
        }
    } catch Error {
    }
    return -1
}

PowerBI_ExpandDrawerElement(element) {
    current := element
    loop 4 {
        if !current
            break
        result := PowerBI_AttemptExpand(current)
        if result != -1
            return result
        try current := UIA.TreeWalkerTrue.GetParentElement(current)
        catch {
            current := ""
        }
    }
    return -1
}

PowerBI_AttemptExpand(element) {
    try {
        hasPattern := element.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
        if hasPattern {
            pat := element.ExpandCollapsePattern
            state := pat.ExpandCollapseState
            if state != UIA.ExpandCollapseState.Expanded {
                pat.Expand()
                Sleep 40
                return 1
            }
            return 0
        }
    } catch Error {
    }
    return -1
}
