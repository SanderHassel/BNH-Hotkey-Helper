#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

; ============================================================================
; BNH HOTKEY HELPER v6.1.2 - BLACKBOX EDITION
; Sander Hasselberg - Birger N. Haug AS
; Sist oppdatert: 2025-10-29
; ============================================================================

; --- KONFIGURASJON ---
global SCRIPT_VERSION := "6.1.2"  ; Oppdatert fra "6.1.1"
global APP_TITLE := "BNH Hotkey Helper"
global STATS_FILE := A_ScriptDir "\BNH_stats.ini"

; --- AUTO-UPDATE KONFIGURASJON ---
global UPDATE_URL := "https://raw.githubusercontent.com/SanderHassel/BNH-Hotkey-Helper/refs/heads/main/BNH_Hotkey_V5_DEMO.ahk"
global UPDATE_INTERVAL := 7200000  ; 120 minutter i millisekunder (60 * 120 * 1000)
global LAST_UPDATE_FILE := A_ScriptDir "\last_update.txt"

; Start auto-update timer
SetTimer(CheckForUpdates, UPDATE_INTERVAL)

; Sjekk også ved oppstart (etter 10 sekunder)
SetTimer(() => CheckForUpdates(), -10000)

; Farger
global COLORS := {
    BG_DARK: "0x1a1a1a",
    BG_MEDIUM: "0x2d2d2d",
    TEXT_WHITE: "0xFFFFFF",
    TEXT_GRAY: "0xCCCCCC",
    BLUE: "0x3399FF",
    RED: "Red",
    GREEN: "0x00CC66",
    ORANGE: "0xFF9900",
    PURPLE: "0x9B59B6",
    DARK_RED: "0xE74C3C",
    CYAN: "0x00BCD4"  ;
}

; Dekk prisliste paths (oppdater årlig)
global DEKK_PATHS := {
    Continental: "O:\Verksted\Felles priser BRUK DENNE\Dekkprisliste Continental høst 2025.xlsx",
    Nokian: "O:\Verksted\Felles priser BRUK DENNE\Dekkprisliste Nokian høst 2025.xlsx"
}

; ============================================================================
; AUTO-UPDATE SYSTEM - v6.0 (MED SEMANTISK VERSJONERING)
; ============================================================================

ExtractVersionFromFile(filePath) {
    try {
        fileContent := FileRead(filePath)
        if RegExMatch(fileContent, 'global SCRIPT_VERSION := "([^"]+)"', &match) {
            return match[1]
        }
        return ""
    } catch {
        return ""
    }
}

CompareVersions(version1, version2) {
    try {
        parts1 := StrSplit(version1, ".")
        parts2 := StrSplit(version2, ".")
        
        while (parts1.Length < 3)
            parts1.Push("0")
        while (parts2.Length < 3)
            parts2.Push("0")
        
        Loop 3 {
            num1 := Integer(parts1[A_Index])
            num2 := Integer(parts2[A_Index])
            
            if (num1 > num2)
                return 1
            else if (num1 < num2)
                return -1
        }
        return 0
    } catch {
        return (version1 = version2) ? 0 : ((version1 > version2) ? 1 : -1)
    }
}

UpdateScript(newFilePath, newVersion) {
    try {
        backupFile := A_ScriptDir "\BNH_Hotkey_Helper_BACKUP_v" SCRIPT_VERSION ".ahk"
        
        try {
            FileCopy(A_ScriptFullPath, backupFile, 1)
        }
        
        FileCopy(newFilePath, A_ScriptFullPath, 1)
        FileDelete(newFilePath)
        
        TrayTip("🎉 Oppdatering fullført!`n`nOppdatert fra v" SCRIPT_VERSION " til v" newVersion "`n`nReloader om 3 sekunder...", "BNH Auto-Update", 0x1)
        
        timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        logEntry := timestamp " - Oppdatert fra v" SCRIPT_VERSION " til v" newVersion "`n"
        FileAppend(logEntry, LAST_UPDATE_FILE)
        
        Sleep(3000)
        Reload()
    } catch as e {
        MsgBox("❌ Oppdateringsfeil:`n`n" e.Message, "Auto-Update", "Icon!")
    }
}

CheckForUpdates() {
    try {
        TrayTip("🔄 Sjekker for oppdateringer...", "BNH Auto-Update", 0x1 | 0x10)
        
        tempFile := A_Temp "\BNH_Hotkey_Helper_Update.ahk"
        
        try {
            Download(UPDATE_URL, tempFile)
        } catch as e {
            TrayTip("❌ Kunne ikke sjekke oppdateringer`n`nKontroller internettforbindelsen.", "BNH Auto-Update", 0x3)
            return
        }
        
        newVersion := ExtractVersionFromFile(tempFile)
        
        if (newVersion = "") {
            FileDelete(tempFile)
            TrayTip("⚠️ Kunne ikke lese versjonsnummer fra oppdateringsfil", "BNH Auto-Update", 0x2)
            return
        }
        
        versionComparison := CompareVersions(newVersion, SCRIPT_VERSION)
        
        if (versionComparison > 0) {
            TrayTip("🎉 Ny versjon funnet!`n`nOppdaterer fra v" SCRIPT_VERSION " til v" newVersion "...", "BNH Auto-Update", 0x1)
            UpdateScript(tempFile, newVersion)
        } else if (versionComparison = 0) {
            FileDelete(tempFile)
            TrayTip("✅ Du har nyeste versjon (v" SCRIPT_VERSION ")", "BNH Auto-Update", 0x1 | 0x10)
        } else {
            FileDelete(tempFile)
            TrayTip("ℹ️ Du bruker en nyere versjon (v" SCRIPT_VERSION ") enn publisert versjon (v" newVersion ")", "BNH Auto-Update", 0x1 | 0x10)
        }
    } catch as e {
        TrayTip("❌ Oppdateringsfeil:`n`n" e.Message, "BNH Auto-Update", 0x3)
    }
}

; ============================================================================
; HOTKEYS - CLIPBOARD OPERASJONER
; ============================================================================

^Space:: {
    try {
        TrackUsage("Fjern mellomrom")
        if !ValidateClipboard()
            return
        A_Clipboard := RemoveSpaces(A_Clipboard)
        Send(A_Clipboard)
    } catch as e {
        ShowError("Fjern mellomrom", e)
    }
}

^!T:: {
    try {
        TrackUsage("Formater telefonnumre")
        if !ValidateClipboard()
            return
        A_Clipboard := FormatPhoneNumbers(A_Clipboard)
        ClipWait(1)
        Send("^v")
    } catch as e {
        ShowError("Formater telefonnumre", e)
    }
}

^+Q:: {
    try {
        TrackUsage("Rabatt kalkulator")
        ; ✅ Åpne kalkulatoren uten å kreve clipboard
        ; Hvis clipboard inneholder et gyldig tall, bruk det - ellers tomt
        ShowDiscountDialog(A_Clipboard)
    } catch as e {
        ShowError("Rabatt kalkulator", e)
    }
}

^!B:: {
    try {
        TrackUsage("Væske-søk")
        ShowFluidSearchDialog()
    } catch as e {
        ShowError("Væske-søk", e)
    }
}

+F5:: {
    try {
        TrackUsage("Clear Cookies + Refresh (Shift+F5)")
        
        if !WinActive("ahk_exe msedge.exe") {
            ShowQuietNotification("⚠️ Fungerer kun i Microsoft Edge")
            return
        }
        
        ; Åpne clear-dialog og la brukeren fullføre
        Send("^+{Delete}")
        
        TrayTip("👆 Trykk 'Clear now' i dialogen`n`nSiden vil refreshe automatisk når du lukker dialogen", "Slett Cookies", 0x1 | 0x10)
        
        ; Vent til dialogen lukkes
        WinWaitNotActive("ahk_exe msedge.exe",, 5)
        
        ; Refresh
        Send("^w")  ; Lukk settings-tab
        Sleep(300)
        Send("^r")  ; Hard refresh
        ShowQuietNotification("✅ Siden refreshet!")
        
    } catch as e {
        ShowError("Clear Cookies", e)
    }
}

; ============================================================================
; HOTKEYS - HURTIG TEKST
; ============================================================================

^+B:: {
    try {
        TrackUsage("Nissan Bremsevæske")
        SendText("00000-01B00")
    } catch as e {
        ShowError("Nissan Bremsevæske", e)
    }
}

^+F:: {
    try {
        TrackUsage("Nissan Frostvæske")
        SendText("00000-01F00")
    } catch as e {
        ShowError("Nissan Frostvæske", e)
    }
}

^+W:: {
    try {
        TrackUsage("20% Rabatt tekst")
        SendText("(20`% Rabatt Inkludert)")
    } catch as e {
        ShowError("Rabatt tekst", e)
    }
}

^+M:: {
    try {
        TrackUsage("MDH Bestilt Levert")
        SendText("MDH Bestilt Levert")
    } catch as e {
        ShowError("MDH tekst", e)
    }
}

^+A:: {
    try {
        TrackUsage("Nøkkelautomat tekst")
        SendText("(eller bruk nøkkelautomat utenfor åpningstid)")
    } catch as e {
        ShowError("Nøkkelautomat tekst", e)
    }
}

^+N:: {
    try {
        TrackUsage("Nøkkelautomat tekst")
        SendText("(eller bruk nøkkelautomat utenfor åpningstid)")
    } catch as e {
        ShowError("Nøkkelautomat tekst", e)
    }
}

^+S:: {
    try {
        TrackUsage("Serviceavtale")
        SendText("Serviceavtale")
    } catch as e {
        ShowError("Serviceavtale", e)
    }
}

^+X:: {
    try {
        TrackUsage("Vetner på bilen")
        SendText("Venter på bilen")
    } catch as e {
        ShowError("Venter på bilen", e)
    }
}

; ============================================================================
; HOTKEYS - SYSTEM
; ============================================================================

^+J:: {
    try {
        TrackUsage("Easter Egg: Ctrl+Shift+J")
        
        ; Åpne i kiosk-modus (Chrome)
        Run('chrome.exe --kiosk "https://ih1.redbubble.net/image.5166131951.3868/flat,750x,075,f-pad,750x1000,f8f8f8.u9.jpg"')
        
    } catch as e {
        ; Fallback til normal åpning hvis Chrome ikke funnet
        try {
            Run("https://ih1.redbubble.net/image.5166131951.3868/flat,750x,075,f-pad,750x1000,f8f8f8.u9.jpg")
            Sleep(1500)
            Send("{F11}")
        } catch {
            ShowError("Open Link", e)
        }
    }
}

^+R:: {
    try {
        TrackUsage("Reload script")
        Reload()
    } catch as e {
        ShowError("Reload", e)
    }
}

^+H:: {
    try {
        TrackUsage("Vis hjelp")
        ShowHelpDialog()
    } catch as e {
        ShowError("Hjelp", e)
    }
}

^+D:: {
    try {
        TrackUsage("Dekktilbud meny")
        ShowDekkDialog()
    } catch as e {
        ShowError("Dekktilbud", e)
    }
}

; ============================================================================
; AUTOFACET SETUP HUB - Ctrl+Shift+P
; ============================================================================

^+P:: {
    try {
        TrackUsage("Autofacet Setup Hub")
        ShowAutofacetSetupHub()
    } catch as e {
        ShowError("Autofacet Setup Hub", e)
    }
}

ShowAutofacetSetupHub() {
    setupGui := Gui("+AlwaysOnTop", "⚙️ Autofacet - Setup Hub")
    setupGui.BackColor := COLORS.BG_DARK
    setupGui.MarginX := 30
    setupGui.MarginY := 30
    
    ; Header med Tesla-stil
    headerBox := setupGui.Add("Text", "w640 h70 Center c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM, "")
    titleText := setupGui.Add("Text", "xp+10 yp+10 w620 h50 Center c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM, "⚙️ AUTOFACET SETUP")
    titleText.SetFont("s18 Bold", "Segoe UI")
    
    infoText := setupGui.Add("Text", "x30 y90 w640 h30 c" COLORS.TEXT_GRAY, "Konfigurer hurtigtaster for ulike Autofacet-funksjoner:")
    infoText.SetFont("s10", "Segoe UI")
    
    ; NULLPUNKT-KORT (ALLTID FØRST)
    nullpunktModule := {
        name: "NULLPUNKT", 
        shortcut: "Klikkes etter hver handling", 
        icon: "🎯", 
        color: "0x2C3E50",  ; Mørk grå
        desc: "Fjerner blå markering", 
        x: 190, 
        y: 130
    }
    CreateAutofacetModuleCard(setupGui, nullpunktModule)
    
    ; Grid layout (3x2) med Tesla-kort
    modules := [
        {name: "LAGRE", shortcut: "Ctrl+Shift+1", icon: "💾", color: COLORS.GREEN, desc: "Lagre endringer", x: 30, y: 280},
        {name: "PLANNER", shortcut: "Ctrl+Shift+2", icon: "📅", color: COLORS.BLUE, desc: "Åpne planner-visning", x: 350, y: 280},
        {name: "KOMMUNIKASJON", shortcut: "Ctrl+Shift+3", icon: "💬", color: COLORS.ORANGE, desc: "Send melding til kunde", x: 30, y: 430},
        {name: "HISTORIKK", shortcut: "Ctrl+Shift+4", icon: "📋", color: COLORS.PURPLE, desc: "Vis kundehistorikk", x: 350, y: 430},
        {name: "OPPDATERINGER", shortcut: "Ctrl+Shift+5", icon: "🔄", color: COLORS.DARK_RED, desc: "Hent nye data", x: 30, y: 580},
        {name: "ARBEIDSORDRE", shortcut: "Ctrl+Shift+|", icon: "📝", color: COLORS.CYAN, desc: "Åpne arbeidsordre", x: 350, y: 580}
    ]
    
    for module in modules {
        CreateAutofacetModuleCard(setupGui, module)
    }
    
    ; Footer med status
    statusText := setupGui.Add("Text", "x30 y730 w640 h25 Center c" COLORS.TEXT_GRAY, "✅ Konfigurer NULLPUNKT først, deretter de andre modulene")
    statusText.SetFont("s9 Italic", "Segoe UI")
    
    closeBtn := CreateStyledButton(setupGui, "x240 y770 w220 h45", "✅ Lukk", COLORS.BLUE, 12)
    closeBtn.OnEvent("Click", (*) => setupGui.Destroy())
    
    setupGui.OnEvent("Close", (*) => setupGui.Destroy())
    setupGui.OnEvent("Escape", (*) => setupGui.Destroy())
    setupGui.Show("w700 h840")
}

; Lag visuelt kort for hver modul
CreateAutofacetModuleCard(gui, module) {
    ; Bakgrunnsboks
    cardBg := gui.Add("Text", "x" module.x " y" module.y " w300 h130 Background" module.color " Border", "")
    
    ; Ikon + tittel
    iconText := gui.Add("Text", "x" (module.x + 15) " y" (module.y + 15) " w60 h60 Center c" COLORS.TEXT_WHITE " Background" module.color, module.icon)
    iconText.SetFont("s32", "Segoe UI")
    
    nameText := gui.Add("Text", "x" (module.x + 85) " y" (module.y + 20) " w200 h30 c" COLORS.TEXT_WHITE " Background" module.color, module.name)
    nameText.SetFont("s14 Bold", "Segoe UI")
    
    ; Beskrivelse
    descText := gui.Add("Text", "x" (module.x + 85) " y" (module.y + 50) " w200 h25 c" COLORS.TEXT_WHITE " Background" module.color, module.desc)
    descText.SetFont("s8", "Segoe UI")
    
    ; Hurtigtast badge - ✅ SENTRERT
    shortcutBox := gui.Add("Text", "x" (module.x + 15) " y" (module.y + 95) " w270 h25 Center c" COLORS.TEXT_WHITE " Background" COLORS.BG_DARK " 0x201", module.shortcut)
    shortcutBox.SetFont("s9 Bold", "Consolas")
    
    ; Klikkhendelse - åpner setup for denne modulen
    cardBg.OnEvent("Click", (*) => StartModuleSetup(module.name, module.shortcut))
    iconText.OnEvent("Click", (*) => StartModuleSetup(module.name, module.shortcut))
    nameText.OnEvent("Click", (*) => StartModuleSetup(module.name, module.shortcut))
    descText.OnEvent("Click", (*) => StartModuleSetup(module.name, module.shortcut))
    shortcutBox.OnEvent("Click", (*) => StartModuleSetup(module.name, module.shortcut))
    
    ; Endre musepeker til hånd
    for ctrl in [cardBg, iconText, nameText, descText, shortcutBox] {
        DllCall("SetClassLongPtr", "Ptr", ctrl.Hwnd, "Int", -12, 
                "Ptr", DllCall("LoadCursor", "Ptr", 0, "Int", 32649, "Ptr"))
    }
}

; Start setup-prosess for valgt modul - FORENKLET (KUN KOORDINATER)
StartModuleSetup(moduleName, shortcut) {
    try {
        TrackUsage("Setup: " moduleName)
        
        ; SPESIELL INSTRUKSJON FOR NULLPUNKT
        if (moduleName = "NULLPUNKT") {
            setupText := "🎯 Konfigurer NULLPUNKT:`n`n"
            setupText .= "NULLPUNKTET brukes til å fjerne blå markering etter klikk.`n`n"
            setupText .= "VIKTIG: Velg et sted på siden som IKKE er en knapp!`n"
            setupText .= "For eksempel: Et tomt område ved siden av logoer eller meny.`n`n"
            setupText .= "1. Trykk OK`n"
            setupText .= "2. Hold musen over et TOMT område (5 sek)`n"
            setupText .= "3. Posisjonen lagres automatisk"
        } else {
            setupText := "🎯 Konfigurer " moduleName "-knapp:`n`n"
            setupText .= "1. Trykk OK for å fortsette`n"
            setupText .= "2. Du har 5 sekunder til å holde musen over knappen`n"
            setupText .= "3. Musposisjonen lagres automatisk`n`n"
            setupText .= "⚠️ VIKTIG: Hold musen HELT STILLE over knappen!"
        }
        
        result := MsgBox(setupText, "Setup: " moduleName, "OKCancel Icon!")
        
        if (result = "Cancel")
            return
        
        ; Nedtelling (5 sekunder)
        Loop 5 {
            remaining := 6 - A_Index
            ToolTip("Lagrer posisjon om " remaining " sekunder...`n`nHold musen stille!", A_ScreenWidth/2, A_ScreenHeight/2)
            Sleep(1000)
        }
        ToolTip()
        
        ; Hent musposisjon
        MouseGetPos(&mx, &my)
        
        ; Lagre koordinater
        configFile := A_ScriptDir "\autofacet_config.ini"
        configSection := moduleName
        
        IniWrite(mx, configFile, configSection, "X")
        IniWrite(my, configFile, configSection, "Y")
        
        ; BEKREFTELSE MED EMOJI
        if (moduleName = "NULLPUNKT") {
            MsgBox("✅ Nullpunkt lagret!`n`n🎯 X: " mx "`n🎯 Y: " my "`n`n💡 Dette punktet klikkes etter hver handling for å fjerne blå markering.", "Suksess", "Iconi T4")
        } else {
            MsgBox("✅ Posisjon lagret!`n`nX: " mx "`nY: " my "`n`nHurtigtast: " shortcut, "Suksess", "Iconi T3")
        }
        
    } catch as e {
        ShowError("StartModuleSetup", e)
    }
}

; ============================================================================
; AUTOFACET MODULER - Ctrl+Shift+1 til 5 (FEEDBACK KUN VED FEIL)
; ============================================================================

^+|:: {
    try {
        TrackUsage("Autofacet ARBEIDSORDRE")
        
        if !WinActive("ahk_exe chrome.exe") && !WinActive("ahk_exe msedge.exe") && !WinActive("ahk_exe brave.exe") {
            ShowQuietNotification("⚠️ Denne funksjonen fungerer kun i Chrome/Edge/Brave")
            return
        }
        
        MouseGetPos(&origX, &origY)
        
        if FindAndClickButton("ARBEIDSORDRE") {
            Sleep(50)
            MouseMove(origX, origY)
        } else {
            ShowQuietNotification("❌ ARBEIDSORDRE ikke konfigurert. Trykk Ctrl+Shift+P for setup.")
        }
    } catch as e {
        ShowError("Autofacet ARBEIDSORDRE", e)
    }
}

^+1:: {
    try {
        TrackUsage("Autofacet LAGRE")
        
        if !WinActive("ahk_exe chrome.exe") && !WinActive("ahk_exe msedge.exe") && !WinActive("ahk_exe brave.exe") {
            ShowQuietNotification("⚠️ Denne funksjonen fungerer kun i Chrome/Edge/Brave")
            return
        }
        
        MouseGetPos(&origX, &origY)
        
        if FindAndClickButton("LAGRE") {
            ; Suksess - ingen feedback
            Sleep(50)
            MouseMove(origX, origY)
        } else {
            ShowQuietNotification("❌ LAGRE ikke konfigurert. Trykk Ctrl+Shift+P for setup.")
        }
    } catch as e {
        ShowError("Autofacet LAGRE", e)
    }
}

^+2:: {
    try {
        TrackUsage("Autofacet PLANNER")
        
        if !WinActive("ahk_exe chrome.exe") && !WinActive("ahk_exe msedge.exe") && !WinActive("ahk_exe brave.exe") {
            ShowQuietNotification("⚠️ Denne funksjonen fungerer kun i Chrome/Edge/Brave")
            return
        }
        
        MouseGetPos(&origX, &origY)
        
        if FindAndClickButton("PLANNER") {
            Sleep(50)
            MouseMove(origX, origY)
        } else {
            ShowQuietNotification("❌ PLANNER ikke konfigurert. Trykk Ctrl+Shift+P for setup.")
        }
    } catch as e {
        ShowError("Autofacet PLANNER", e)
    }
}

^+3:: {
    try {
        TrackUsage("Autofacet KOMMUNIKASJON")
        
        if !WinActive("ahk_exe chrome.exe") && !WinActive("ahk_exe msedge.exe") && !WinActive("ahk_exe brave.exe") {
            ShowQuietNotification("⚠️ Denne funksjonen fungerer kun i Chrome/Edge/Brave")
            return
        }
        
        MouseGetPos(&origX, &origY)
        
        if FindAndClickButton("KOMMUNIKASJON") {
            Sleep(50)
            MouseMove(origX, origY)
        } else {
            ShowQuietNotification("❌ KOMMUNIKASJON ikke konfigurert. Trykk Ctrl+Shift+P for setup.")
        }
    } catch as e {
        ShowError("Autofacet KOMMUNIKASJON", e)
    }
}

^+4:: {
    try {
        TrackUsage("Autofacet HISTORIKK")
        
        if !WinActive("ahk_exe chrome.exe") && !WinActive("ahk_exe msedge.exe") && !WinActive("ahk_exe brave.exe") {
            ShowQuietNotification("⚠️ Denne funksjonen fungerer kun i Chrome/Edge/Brave")
            return
        }
        
        MouseGetPos(&origX, &origY)
        
        if FindAndClickButton("HISTORIKK") {
            Sleep(50)
            MouseMove(origX, origY)
        } else {
            ShowQuietNotification("❌ HISTORIKK ikke konfigurert. Trykk Ctrl+Shift+P for setup.")
        }
    } catch as e {
        ShowError("Autofacet HISTORIKK", e)
    }
}

^+5:: {
    try {
        TrackUsage("Autofacet OPPDATERINGER")
        
        if !WinActive("ahk_exe chrome.exe") && !WinActive("ahk_exe msedge.exe") && !WinActive("ahk_exe brave.exe") {
            ShowQuietNotification("⚠️ Denne funksjonen fungerer kun i Chrome/Edge/Brave")
            return
        }
        
        MouseGetPos(&origX, &origY)
        
        if FindAndClickButton("OPPDATERINGER") {
            Sleep(50)
            MouseMove(origX, origY)
        } else {
            ShowQuietNotification("❌ OPPDATERINGER ikke konfigurert. Trykk Ctrl+Shift+P for setup.")
        }
    } catch as e {
        ShowError("Autofacet OPPDATERINGER", e)
    }
}

; ============================================================================
; AUTOFACET KLIKK-FUNKSJON - v6.0 (MED NULLPUNKT)
; ============================================================================

FindAndClickButton(configSection) {
    try {
        configFile := A_ScriptDir "\autofacet_config.ini"
        
        ; Les lagrede koordinater for målknappen
        if FileExist(configFile) {
            savedX := IniRead(configFile, configSection, "X", "")
            savedY := IniRead(configFile, configSection, "Y", "")
            
            if (savedX != "" && savedY != "") {
                targetX := Integer(savedX)
                targetY := Integer(savedY)
                
                ; STEG 1: Klikk på målknappen
                MouseMove(targetX, targetY, 0)
                Sleep(15)
                Click("Left")
                
                ; STEG 2: Les nullpunkt-koordinater
                nullX := IniRead(configFile, "NULLPUNKT", "X", "")
                nullY := IniRead(configFile, "NULLPUNKT", "Y", "")
                
                ; STEG 3: Hvis nullpunkt finnes, klikk der for å fjerne markering
                if (nullX != "" && nullY != "") {
                    Sleep(600)  ; Kort pause før nullpunkt-klikk
                    MouseMove(Integer(nullX), Integer(nullY), 0)
                    Sleep(10)
                    Click("Left")
                }
                
                return true
            }
        }
        
        return false
        
    } catch {
        return false
    }
}

; ============================================================================
; STILLE NOTIFIKASJON - KUN VED FEIL (INGEN LYD)
; ============================================================================

ShowQuietNotification(message, duration := 2500) {
    try {
        ; Tooltip nederst til høyre på skjermen
        x := A_ScreenWidth - 350
        y := A_ScreenHeight - 100
        
        ToolTip(message, x, y)
        SetTimer(() => ToolTip(), -duration)
    } catch {
        ; Silent fail
    }
}

; ============================================================================
; FUNKSJONER - REGNUMMER VALIDERING
; ============================================================================

ValidateLicensePlate(text) {
    try {
        ; Trim og fjern mellomrom
        cleaned := Trim(text)
        cleaned := StrReplace(cleaned, " ", "")
        cleaned := StrReplace(cleaned, Chr(160), "")  ; Non-breaking space
        
        ; MØNSTER: 2 bokstaver + 5 tall (norsk regnummer)
        ; Eksempel: EB25688, AB12345
        if RegExMatch(cleaned, "^[A-ZÆØÅa-zæøå]{2}\d{5}$") {
            ; Konverter til store bokstaver
            return StrUpper(cleaned)
        }
        
        ; Ikke gyldig regnummer
        return ""
    } catch {
        return ""
    }
}

ProcessHotstringWithPlate(templateText) {
    try {
        ; Sjekk om clipboard inneholder et gyldig regnummer
        licensePlate := ValidateLicensePlate(A_Clipboard)
        
        ; Hvis gyldig, erstatt {LICENSEPLATE}
        if (licensePlate != "") {
            return StrReplace(templateText, "{LICENSEPLATE}", licensePlate)
        }
        
        ; Hvis ikke gyldig, behold {LICENSEPLATE} som placeholder
        return templateText
    } catch {
        return templateText
    }
}

; ============================================================================
; HOTSTRINGS - STANDARD TEKSTER
; ============================================================================

:*:*garanti::
{
    try {
        TrackUsage("Hotstring: *garanti")
        SendText("Verkstedet vil undersøke det du opplever og om det er berettiget mot garanti. Dersom det dekkes av garanti er det kostnadsfritt, men om det er grunnet ytre påvirkninger eller det ikke er dekket av garanti blir man belastet for feilsøk 1 990,- for inntil en time per punkt. Dersom det er behov for mer tid, vil du bli kontaktet av verkstedet for godkjennelse av videre feilsøk og pris.")
    }
}

:*:*info::
{
    try {
        TrackUsage("Hotstring: *info")
        SendText("Bilen leveres kl 08:00 og hentes kl 15:30. Du kan også levere/hente utenfor åpningstid i nøkkelautomat. Vi tilbyr leiebil fra 698,- pr døgn, må avtales minimum 48 timer før verkstedtimen. Du kan besvare denne e-posten eller ringe oss på tlf 400 10 400.")
    }
}

:*:*inntil::
{
    try {
        TrackUsage("Hotstring: *inntil")
        SendText("*Inntil 1 time feilsøk.")
    }
}

:*:*anbefalt::
{
    try {
        TrackUsage("Hotstring: *anbefalt")
        SendText("Vi anbefaler årlig skift av vindusviskere foran (XXX,-), batteri i nøklene (50,- pr stk) og at man utfører motorvask (390,- sammen med service). Ønsker du at vi utfører noen av disse punktene i tillegg? Hører fra deg om du ønsker dette.")
    }
}

:*:*avdnr::
{
    try {
        TrackUsage("Hotstring: *avdnr")
        SendText("01 - Røyken, 02 - Rud, 03 - Oslo, 05 - Follo, 06 - Drammen, 07 - Lillestrøm, 08 - Gardermoen.")
    }
}

:*:*opsms-::
{
    try {
        TrackUsage("Hotstring: *opsms-")
        template := "Hei, vi har forsøkt å ringe deg. Basert på våre opplysninger er det tid for service på din bil med regnr: {LICENSEPLATE}. Bestill time raskt og enkelt på nett: https://service.bnh.no/ Hilsen Birger N. Haug / 40010400."
        SendText(ProcessHotstringWithPlate(template))
    }
}

:*:*opsms+::
{
    try {
        TrackUsage("Hotstring: *opsms+")
        template := "Hei, vi har forsøkt å ringe deg. Det er på tide med service på {LICENSEPLATE}. Du har allerede en forhåndsbetalt serviceavtale. Bestill time här: https://service.bnh.no/ eller ring oss på 40010400. Hilsen Birger N. Haug."
        SendText(ProcessHotstringWithPlate(template))
    }
}

:*:*samtykke::
{
    try {
        TrackUsage("Hotstring: *samtykke")
        SendText("Vi har ikke registrert ditt samtykke til elektronisk kommunikasjon og markedsføring. Vil du motta nyttige oppdateringer om bilholdet ditt – som nyhetsbrev, eksklusive tilbud, konkurranser og relevant informasjon på e-post eller SMS?.")
    }
}

:*:*forkontroll::
{
    try {
        TrackUsage("Hotstring: *forkontroll")
        SendText("Vi ønsker å utføre en forkontroll av lyden først. På denne måten kan vi utelukke eventuelle feil, bestille nødvendige deler og estimere kostnad for reparasjon. Du blir da med mekaniker på kjøretur for å fremvise og forklare lyden. Dette tar ca en halvtime og koster 490,- dersom dette ikke dekkes av garanti. Passer det for deg den DATO kl. XX:XX? Du kan svare på denne e-posten..")
    }
}

:*:*leiebil::
{
    try {
        TrackUsage("Hotstring: *leiebil")
        SendText("Du vil få tilsendt en SMS i forbindelse med din leiebilbestilling. Vennligst gå inn på den snarest og registrer dine kontakt-og kortopplysninger. Beløpet vil bli reservert på kortet og ikke trukket før fakturering.")
    }
}

:*:*ikkevent::
{
    try {
        TrackUsage("Hotstring: *ikkevent")
        SendText("OBS: Vi kan dessverre ikke tilby ventetime på denne type jobb. Avtalen er endret til levering kl. 08 og hente kl. 16. Du kan også levere/hente utenfor åpningstid i nøkkelautomat. Du blir kontaktet når bilen er klar. Vi tilbyr leiebil fra 698,- pr døgn, må avtales minimum 48 timer før verkstedtimen.")
    }
}

; ============================================================================
; ERROR HANDLING
; ============================================================================

ShowError(functionName, errorObj) {
    errorMsg := "❌ Feil i: " functionName "`n`n"
    errorMsg .= "Melding: " errorObj.Message "`n"
    errorMsg .= "Linje: " errorObj.Line "`n"
    errorMsg .= "What: " errorObj.What "`n`n"
    errorMsg .= "Scriptet fortsetter å kjøre."
    
    MsgBox(errorMsg, "BNH Hotkey Helper - Feil", "Icon! T10")
    
    try {
        logFile := A_ScriptDir "\BNH_errors.log"
        timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        FileAppend(timestamp " - " functionName " - " errorObj.Message " (Linje " errorObj.Line ")`n", logFile)
    }
}

; ============================================================================
; STATISTIKK - TRACKING
; ============================================================================

TrackUsage(functionName) {
    try {
        global STATS_FILE
        currentCount := IniRead(STATS_FILE, "Usage", functionName, 0)
        newCount := Integer(currentCount) + 1
        IniWrite(newCount, STATS_FILE, "Usage", functionName)
    } catch {
        ; Silent fail
    }
}

GetUsageStats() {
    try {
        global STATS_FILE
        if !FileExist(STATS_FILE)
            return Map()
        
        stats := Map()
        section := IniRead(STATS_FILE, "Usage", "", "")
        
        if (section = "")
            return stats
        
        lines := StrSplit(section, "`n")
        for line in lines {
            if (Trim(line) = "")
                continue
            parts := StrSplit(line, "=")
            if (parts.Length >= 2) {
                key := Trim(parts[1])
                value := Integer(Trim(parts[2]))
                stats[key] := value
            }
        }
        return stats
    } catch {
        return Map()
    }
}

ShowStatsDialog() {
    try {
        statsGui := Gui("+AlwaysOnTop +Resize", "BNH - Bruksstatistikk")
        statsGui.BackColor := COLORS.BG_DARK
        statsGui.MarginX := 20
        statsGui.MarginY := 20
        
        titleText := statsGui.Add("Text", "w560 h40 Center c" COLORS.TEXT_WHITE " Section", "📊 Bruksstatistikk")
        titleText.SetFont("s16 Bold", "Segoe UI")
        
        infoText := statsGui.Add("Text", "w560 h30 c" COLORS.TEXT_GRAY " xs y+5", "Her ser du hvor mange ganger hver funksjon er brukt:")
        infoText.SetFont("s10", "Segoe UI")
        
        lv := statsGui.Add("ListView", "w560 h350 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " xs y+10", 
                          ["Rang", "Funksjon", "Antall brukt"])
        lv.SetFont("s10", "Segoe UI")
        
        stats := GetUsageStats()
        statsArray := []
        for funcName, count in stats
            statsArray.Push({name: funcName, count: count})
        
        statsArray := SortByCount(statsArray)
        
        rank := 1
        for item in statsArray {
            lv.Add(, rank, item.name, item.count)
            rank++
        }
        
        if (statsArray.Length = 0)
            lv.Add(, "-", "Ingen data ennå - bruk scriptet for å se statistikk!", "0")
        
        lv.ModifyCol(1, "60 Center")
        lv.ModifyCol(2, "350")
        lv.ModifyCol(3, "120 Center")
        
        resetBtn := CreateStyledButton(statsGui, "xs w270 h40 y+15", "🗑️ Nullstill statistikk", COLORS.RED, 11)
        resetBtn.OnEvent("Click", (*) => ResetStats())
        
        closeBtn := CreateStyledButton(statsGui, "x+20 w270 h40 yp", "✅ Lukk", COLORS.BLUE, 11)
        closeBtn.OnEvent("Click", (*) => statsGui.Destroy())
        
        statsGui.OnEvent("Close", (*) => statsGui.Destroy())
        statsGui.OnEvent("Escape", (*) => statsGui.Destroy())
        statsGui.OnEvent("Size", ResizeStatsHandler)
        
        statsGui.Show("w600 h550")
        
        ResetStats() {
            result := MsgBox("Er du sikker på at du vil nullstille all statistikk?`n`nDette kan ikke angres!", "Bekreft nullstilling", "YesNo Icon?")
            if (result = "Yes") {
                global STATS_FILE
                try {
                    if FileExist(STATS_FILE)
                        FileDelete(STATS_FILE)
                    lv.Delete()
                    lv.Add(, "-", "Statistikk nullstilt", "0")
                    TrayTip("All statistikk er slettet", "✅ Nullstilt")
                } catch as e {
                    MsgBox("Kunne ikke slette statistikk: " e.Message)
                }
            }
        }
        
        ResizeStatsHandler(gui, minMax, width, height) {
            if (minMax = -1)
                return
            titleText.Move(, , width - 40)
            infoText.Move(, , width - 40)
            lv.Move(, , width - 40, height - 180)
            buttonY := height - 80
            resetBtn.Move(, buttonY)
            closeBtn.Move(, buttonY)
        }
    } catch as e {
        ShowError("Vis statistikk", e)
    }
}

SortByCount(arr) {
    try {
        n := arr.Length
        if (n <= 1)
            return arr
            
        Loop n - 1 {
            i := A_Index
            Loop n - i {
                j := A_Index
                if (arr[j].count < arr[j + 1].count) {
                    temp := arr[j]
                    arr[j] := arr[j + 1]
                    arr[j + 1] := temp
                }
            }
        }
        return arr
    } catch {
        return arr
    }
}

; ============================================================================
; FUNKSJONER - CLIPBOARD VALIDERING
; ============================================================================

ValidateClipboard(customMsg := "") {
    try {
        if (A_Clipboard = "") {
            msg := customMsg != "" ? customMsg : "Kopier noe først."
            MsgBox(APP_TITLE " v" SCRIPT_VERSION "`n`n" msg)
            return false
        }
        return true
    } catch {
        return false
    }
}

RemoveSpaces(text) {
    try {
        text := StrReplace(text, " ", "")
        text := StrReplace(text, Chr(160), "")
        return text
    } catch {
        return text
    }
}

; ============================================================================
; FUNKSJONER - TELEFONNUMMER FORMATERING
; ============================================================================

FormatPhoneNumbers(text) {
    try {
        rows := StrSplit(text, ["`n", "`r"])
        result := []
        
        for row in rows {
            if (Trim(row) = "")
                continue
            cells := StrSplit(row, "`t")
            newRow := []
            for cell in cells {
                if (Trim(cell) = "")
                    continue
                digits := RegExReplace(cell, "[^\d]", "")
                if (StrLen(digits) > 8)
                    digits := SubStr(digits, -8)
                if (digits != "")
                    newRow.Push("+47" . digits)
            }
            if (newRow.Length > 0)
                result.Push(ArrayJoin(newRow, "`t"))
        }
        return ArrayJoin(result, "`n")
    } catch {
        return text
    }
}

ArrayJoin(arr, delimiter) {
    try {
        result := ""
        for item in arr {
            if (result != "")
                result .= delimiter
            result .= item
        }
        return result
    } catch {
        return ""
    }
}

; ============================================================================
; FUNKSJONER - GUI HELPERS
; ============================================================================

CreateStyledButton(gui, options, text, bgColor, fontSize) {
    try {
        btn := gui.Add("Text", options " Center c" COLORS.TEXT_WHITE " Background" bgColor " Border 0x200", text)
        btn.SetFont("s" fontSize " Bold", "Segoe UI")
        DllCall("SetClassLongPtr", "Ptr", btn.Hwnd, "Int", -12, 
                "Ptr", DllCall("LoadCursor", "Ptr", 0, "Int", 32649, "Ptr"))
        return btn
    } catch {
        return gui.Add("Button", options, text)
    }
}

; ============================================================================
; GUI - VÆSKE-SØK
; ============================================================================

ShowFluidSearchDialog() {
    try {
        fluidGui := Gui("+AlwaysOnTop +Resize", "BNH - Væske Søk")
        fluidGui.BackColor := COLORS.BG_DARK
        fluidGui.MarginX := 20
        fluidGui.MarginY := 20
        
        titleText := fluidGui.Add("Text", "w560 h40 Center c" COLORS.TEXT_WHITE " Section", "Søk etter væske varenummer")
        titleText.SetFont("s14 Bold", "Segoe UI")
        
        searchLabel := fluidGui.Add("Text", "w560 h20 c" COLORS.TEXT_GRAY " xs y+10", "Søk (gammel eller ny varenummer, merke, beskrivelse):")
        searchLabel.SetFont("s10", "Segoe UI")
        
        searchBox := fluidGui.Add("Edit", "w560 h35 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " xs y+5")
        searchBox.SetFont("s11", "Segoe UI")
        searchBox.OnEvent("Change", (*) => UpdateFluidListView())
        
        lv := fluidGui.Add("ListView", "w560 h300 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " xs y+15", 
                          ["Merke", "Beskrivelse", "Gammelt Varenr.", "Nytt Varenr."])
        lv.SetFont("s10", "Segoe UI")
        
        fluids := GetFluidDatabase()
        PopulateFluidListView()
        
        copyBtn := CreateStyledButton(fluidGui, "xs w270 h40 y+15", "📋 Kopier Varenummer", COLORS.GREEN, 11)
        copyBtn.OnEvent("Click", (*) => CopySelectedFluid())
        
        pasteBtn := CreateStyledButton(fluidGui, "x+20 w270 h40 yp", "✅ Lim Inn (Enter)", COLORS.BLUE, 11)
        pasteBtn.OnEvent("Click", (*) => PasteSelectedFluid())
        
        fluidGui.OnEvent("Size", ResizeFluidHandler)
        fluidGui.OnEvent("Close", (*) => CleanupAndClose())
        fluidGui.OnEvent("Escape", (*) => CleanupAndClose())
        lv.OnEvent("DoubleClick", (*) => PasteSelectedFluid())
        
        fluidGui.Show("w600 h570")
        searchBox.Focus()
        
        HotIfWinActive("ahk_id " fluidGui.Hwnd)
        Hotkey("Enter", (*) => PasteSelectedFluid(), "On")
        
        PopulateFluidListView() {
            try {
                lv.Delete()
                for item in fluids
                    lv.Add(, item.brand, item.description, item.oldNumber, item.newNumber)
                lv.ModifyCol(1, "100")
                lv.ModifyCol(2, "150")
                lv.ModifyCol(3, "140")
                lv.ModifyCol(4, "140")
            }
        }
        
        UpdateFluidListView() {
            try {
                searchText := Trim(searchBox.Text)
                lv.Delete()
                if (searchText = "") {
                    PopulateFluidListView()
                    return
                }
                for item in fluids {
                    if (InStr(item.brand, searchText, false) 
                        || InStr(item.description, searchText, false)
                        || InStr(item.oldNumber, searchText, false) 
                        || InStr(item.newNumber, searchText, false))
                        lv.Add(, item.brand, item.description, item.oldNumber, item.newNumber)
                }
                lv.ModifyCol(1, "100")
                lv.ModifyCol(2, "150")
                lv.ModifyCol(3, "140")
                lv.ModifyCol(4, "140")
            }
        }
        
        CopySelectedFluid() {
            try {
                rowNum := lv.GetNext(0)
                if (rowNum = 0) {
                    MsgBox("Velg en væske fra listen først!")
                    return
                }
                newNumber := lv.GetText(rowNum, 4)
                A_Clipboard := ""
                Sleep(50)
                A_Clipboard := newNumber
                ClipWait(1)
                HotIfWinActive("ahk_id " fluidGui.Hwnd)
                Hotkey("Enter", "Off")
                HotIfWinActive()
                fluidGui.Destroy()
                TrayTip(newNumber " er kopiert til utklippstavlen", "Kopiert!")
            }
        }
        
        PasteSelectedFluid() {
            try {
                rowNum := lv.GetNext(0)
                if (rowNum = 0) {
                    MsgBox("Velg en væske fra listen først!")
                    return
                }
                newNumber := lv.GetText(rowNum, 4)
                HotIfWinActive("ahk_id " fluidGui.Hwnd)
                Hotkey("Enter", "Off")
                HotIfWinActive()
                fluidGui.Destroy()
                Send(newNumber)
            }
        }
        
        CleanupAndClose() {
            try {
                HotIfWinActive("ahk_id " fluidGui.Hwnd)
                Hotkey("Enter", "Off")
                HotIfWinActive()
                fluidGui.Destroy()
            }
        }
        
        ResizeFluidHandler(gui, minMax, width, height) {
            try {
                if (minMax = -1)
                    return
                titleText.Move(, , width - 40)
                searchLabel.Move(, , width - 40)
                searchBox.Move(, , width - 40)
                lv.Move(, , width - 40, height - 230)
                buttonY := height - 80
                copyBtn.Move(, buttonY)
                pasteBtn.Move(, buttonY)
            }
        }
    } catch as e {
        ShowError("Væske-søk dialog", e)
    }
}

GetFluidDatabase() {
    try {
        static fluidDB := ""
        if (fluidDB = "") {
            fluidDB := [
                {brand: "NISSAN", description: "Bremsevæske", oldNumber: "KE90399932", newNumber: "00000-01B00"},
                {brand: "NISSAN", description: "Frostvæske", oldNumber: "KE90299945", newNumber: "00000-01F00"}
            ]
        }
        return fluidDB
    } catch {
        return []
    }
}

; ============================================================================
; GUI - RABATT KALKULATOR v2.0 (MED LIVE-BEREGNING)
; ============================================================================

ShowDiscountDialog(originalValue := "") {
    try {
        rabattGui := Gui("+AlwaysOnTop", "BNH - Rabatt Kalkulator")
        rabattGui.BackColor := COLORS.BG_DARK
        rabattGui.MarginX := 20
        rabattGui.MarginY := 20
        
        titleText := rabattGui.Add("Text", "w320 h35 Center c" COLORS.TEXT_WHITE " Section", "💰 Rabatt Kalkulator")
        titleText.SetFont("s14 Bold", "Segoe UI")
        
        originalLabel := rabattGui.Add("Text", "w320 h20 c" COLORS.TEXT_GRAY " xs y+15", "Original pris:")
        originalLabel.SetFont("s9", "Segoe UI")
        
        ; VALIDER CLIPBOARD - KUN BRUK HVIS DET ER ETT RENT TALL
        cleanValue := ""
        if (originalValue != "") {
            testValue := Trim(originalValue)
            
            if (StrLen(testValue) > 15) {
                cleanValue := ""
            }
            else if RegExMatch(testValue, "^[\s\d\.,\-]+$") {
                normalized := StrReplace(testValue, " ", "")
                normalized := StrReplace(normalized, Chr(160), "")
                
                if (InStr(normalized, ".") && InStr(normalized, ",")) {
                    normalized := StrReplace(normalized, ".", "")
                    normalized := StrReplace(normalized, ",", ".")
                }
                else if (InStr(normalized, ",")) {
                    normalized := StrReplace(normalized, ",", ".")
                }
                
                normalized := RegExReplace(normalized, "[,\.]\-$", "")
                
                if (IsNumber(normalized)) {
                    floatValue := Float(normalized)
                    finalValue := Integer(Ceil(floatValue))
                    
                    if (finalValue > 0 && finalValue <= 10000000) {
                        cleanValue := String(finalValue)
                    }
                }
            }
        }
        
        originalInput := rabattGui.Add("Edit", "w320 h35 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " xs y+5", cleanValue)
        originalInput.SetFont("s11", "Segoe UI")
        originalInput.OnEvent("Change", (*) => CleanAndCalculate())
        
        rabattLabel := rabattGui.Add("Text", "w320 h20 c" COLORS.TEXT_GRAY " xs y+15", "Rabatt (%):")
        rabattLabel.SetFont("s9", "Segoe UI")
        
        rabattInput := rabattGui.Add("Edit", "w300 h35 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " xs y+5", "20")
        rabattInput.SetFont("s11", "Segoe UI")
        rabattInput.OnEvent("Change", (*) => CalculateDiscount())
        
        rabattUpDown := rabattGui.Add("UpDown", "Range0-100", 20)
        
        buttonLabel := rabattGui.Add("Text", "w320 h20 c" COLORS.TEXT_GRAY " xs y+15", "Hurtigvalg:")
        buttonLabel.SetFont("s9", "Segoe UI")
        
        btn10 := rabattGui.Add("Button", "w73 h30 xs y+5", "10%")
        btn10.SetFont("s9", "Segoe UI")
        btn10.OnEvent("Click", (*) => SetRabatt(10))
        
        btn15 := rabattGui.Add("Button", "w73 h30 x+5", "15%")
        btn15.SetFont("s9", "Segoe UI")
        btn15.OnEvent("Click", (*) => SetRabatt(15))
        
        btn20 := rabattGui.Add("Button", "w73 h30 x+5", "20%")
        btn20.SetFont("s9", "Segoe UI")
        btn20.OnEvent("Click", (*) => SetRabatt(20))
        
        btn25 := rabattGui.Add("Button", "w73 h30 x+5", "25%")
        btn25.SetFont("s9", "Segoe UI")
        btn25.OnEvent("Click", (*) => SetRabatt(25))
        
        resultBox := rabattGui.Add("Text", "w320 h80 Center c" COLORS.TEXT_WHITE " Background" COLORS.GREEN " xs y+20 Border", "")
        resultBox.SetFont("s12 Bold", "Segoe UI")
        
        resultLabel := rabattGui.Add("Text", "xp+10 yp+10 w300 h20 Center c" COLORS.TEXT_WHITE " Background" COLORS.GREEN, "NY PRIS:")
        resultLabel.SetFont("s9", "Segoe UI")
        
        resultValue := rabattGui.Add("Text", "xp yp+25 w300 h40 Center c" COLORS.TEXT_WHITE " Background" COLORS.GREEN, "0 kr")
        resultValue.SetFont("s20 Bold", "Segoe UI")
        
        sendBtn := CreateStyledButton(rabattGui, "xs w155 h45 y+15", "📤 Send (Enter)", COLORS.BLUE, 11)
        sendBtn.OnEvent("Click", (*) => SendDiscountedPrice())
        
        copyBtn := CreateStyledButton(rabattGui, "x+10 w155 h45 yp", "📋 Kopier", COLORS.ORANGE, 11)
        copyBtn.OnEvent("Click", (*) => CopyDiscountedPrice())
        
        cancelBtn := CreateStyledButton(rabattGui, "xs w320 h35 y+10", "Lukk", COLORS.RED, 9)
        cancelBtn.OnEvent("Click", (*) => rabattGui.Destroy())
        
        rabattGui.OnEvent("Close", (*) => rabattGui.Destroy())
        rabattGui.OnEvent("Escape", (*) => rabattGui.Destroy())
        rabattGui.Show("w360 h590")
        originalInput.Focus()
        
        HotIfWinActive("ahk_id " rabattGui.Hwnd)
        Hotkey("Enter", (*) => SendDiscountedPrice(), "On")
        
        CalculateDiscount()
        
        SetRabatt(value) {
            try {
                rabattInput.Value := value
                rabattUpDown.Value := value
                CalculateDiscount()
            }
        }
        
        CleanAndCalculate() {
            try {
                currentValue := originalInput.Text
                
                ; Skip hvis allerede er rent tall
                if (RegExMatch(currentValue, "^\d+$")) {
                    CalculateDiscount()
                    return
                }
                
                ; Rens input
                if (currentValue != "" && RegExMatch(currentValue, "[\d\.,\-\s]+")) {
                    normalized := StrReplace(currentValue, " ", "")
                    normalized := StrReplace(normalized, Chr(160), "")
                    
                    if (InStr(normalized, ".") && InStr(normalized, ",")) {
                        normalized := StrReplace(normalized, ".", "")
                        normalized := StrReplace(normalized, ",", ".")
                    }
                    else if (InStr(normalized, ",")) {
                        normalized := StrReplace(normalized, ",", ".")
                    }
                    
                    normalized := RegExReplace(normalized, "[,\.]\-$", "")
                    
                    if (IsNumber(normalized)) {
                        floatValue := Float(normalized)
                        cleanedValue := Integer(Ceil(floatValue))
                        
                        if (String(cleanedValue) != currentValue) {
                            cursorPos := StrLen(String(cleanedValue))
                            originalInput.Value := String(cleanedValue)
                            SendMessage(0xB1, cursorPos, cursorPos, originalInput)
                        }
                    }
                }
                
                CalculateDiscount()
                
            } catch {
                CalculateDiscount()
            }
        }
        
        CalculateDiscount() {
            try {
                originalStr := Trim(originalInput.Text)
                rabattStr := Trim(StrReplace(rabattInput.Text, "%", ""))
                
                if (originalStr = "" || !IsNumber(originalStr)) {
                    resultValue.Text := "Skriv inn pris"
                    return
                }
                
                if (rabattStr = "" || !IsNumber(rabattStr)) {
                    resultValue.Text := "Ugyldig %"
                    return
                }
                
                originalPrice := Float(originalStr)
                rabattPercent := Float(rabattStr)
                newPrice := Integer(Ceil(originalPrice * (1 - rabattPercent / 100)))
                resultValue.Text := newPrice " kr"
                
                if (rabattPercent >= 25)
                    resultBox.Opt("Background" COLORS.DARK_RED)
                else if (rabattPercent >= 15)
                    resultBox.Opt("Background" COLORS.ORANGE)
                else
                    resultBox.Opt("Background" COLORS.GREEN)
                
            } catch {
                resultValue.Text := "Feil"
            }
        }
        
        SendDiscountedPrice() {
            try {
                resultText := Trim(StrReplace(resultValue.Text, " kr", ""))
                
                if (resultText = "" || resultText = "Skriv inn pris" || resultText = "Ugyldig %" || resultText = "Feil") {
                    MsgBox("Vennligst skriv inn gyldig pris og rabatt!", "Ugyldig input", "Icon!")
                    return
                }
                
                HotIfWinActive("ahk_id " rabattGui.Hwnd)
                Hotkey("Enter", "Off")
                HotIfWinActive()
                
                rabattGui.Destroy()
                Send(resultText)
                
            } catch as e {
                ShowError("Send rabattert pris", e)
            }
        }
        
        CopyDiscountedPrice() {
            try {
                resultText := Trim(StrReplace(resultValue.Text, " kr", ""))
                
                if (resultText = "" || resultText = "Skriv inn pris" || resultText = "Ugyldig %" || resultText = "Feil") {
                    MsgBox("Vennligst skriv inn gyldig pris og rabatt!", "Ugyldig input", "Icon!")
                    return
                }
                
                A_Clipboard := resultText
                TrayTip(resultText " kr kopiert til utklippstavlen", "Kopiert!", 0x1 | 0x10)
                
            } catch as e {
                ShowError("Kopier rabattert pris", e)
            }
        }
        
    } catch as e {
        ShowError("Rabatt kalkulator", e)
    }
}

; ============================================================================
; GUI - DEKKTILBUD MENY
; ============================================================================

ShowDekkDialog() {
    dekkGui := Gui("+AlwaysOnTop", "Velg dekktilbud")
    dekkGui.BackColor := COLORS.BG_DARK
    dekkGui.MarginX := 20
    dekkGui.MarginY := 20
    
    titleText := dekkGui.Add("Text", "w300 h30 Center c" COLORS.TEXT_WHITE, "Velg dekktilbud:")
    titleText.SetFont("s12 Bold", "Segoe UI")
    
    btnCont := dekkGui.Add("Button", "w140 y+10", "Continental")
    btnCont.OnEvent("Click", (*) => OpenDekkFile("Continental", dekkGui))
    
    btnNok := dekkGui.Add("Button", "w140 x+10", "Nokian")
    btnNok.OnEvent("Click", (*) => OpenDekkFile("Nokian", dekkGui))
    
    dekkGui.Show("w320")
}

OpenDekkFile(brand, gui) {
    path := DEKK_PATHS.%brand%
    try {
        if !FileExist(path)
            throw Error("Finner ikke fil: " path)
        Run(path)
        gui.Destroy()
    } catch as e {
        MsgBox("Feil ved åpning av " brand " prisliste:`n" e.Message)
    }
}

; ============================================================================
; GUI - HJELP DIALOG
; ============================================================================

ShowHelpDialog() {
    try {
        helpGui := Gui("+Resize", APP_TITLE " v" SCRIPT_VERSION)
        helpGui.BackColor := COLORS.BG_DARK
        helpGui.MarginX := 20
        helpGui.MarginY := 20
        
        titleText := helpGui.Add("Text", "w560 h40 Center c" COLORS.TEXT_WHITE " Section", APP_TITLE)
        titleText.SetFont("s18 Bold", "Segoe UI")
        
        searchLabel := helpGui.Add("Text", "w560 h20 c" COLORS.TEXT_GRAY " xs y+10", "Søk i hurtigtaster og tekster:")
        searchLabel.SetFont("s10", "Segoe UI")
        
        searchBox := helpGui.Add("Edit", "w560 h30 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " xs y+5")
        searchBox.SetFont("s10", "Segoe UI")
        searchBox.OnEvent("Change", (*) => UpdateHelpListView())
        
        lv := helpGui.Add("ListView", "w560 h350 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " xs y+15 Sort", 
                          ["Hurtigtast/Tekst", "Beskrivelse"])
        lv.SetFont("s9", "Segoe UI")
        
        hotkeys := GetHotkeysMap()
        hotstrings := GetHotstringsMap()
        PopulateHelpListView()
        
        helpGui.OnEvent("Size", ResizeHelpHandler)
        helpGui.OnEvent("Close", (*) => helpGui.Destroy())
        helpGui.Show("w600 h500")
        searchBox.Focus()
        
        PopulateHelpListView() {
            try {
                lv.Delete()
                for hotkeyText, desc in hotkeys
                    lv.Add(, hotkeyText, desc)
                for trigger, text in hotstrings {
                    shortText := SubStr(text, 1, 60) . (StrLen(text) > 60 ? "..." : "")
                    lv.Add(, trigger, shortText)
                }
                lv.ModifyCol(1, "AutoHdr")
                lv.ModifyCol(2, "AutoHdr")
            }
        }
        
        UpdateHelpListView() {
            try {
                searchText := searchBox.Text
                lv.Delete()
                if (searchText = "") {
                    PopulateHelpListView()
                    return
                }
                for hotkeyText, desc in hotkeys {
                    if (InStr(hotkeyText, searchText, false) || InStr(desc, searchText, false))
                        lv.Add(, hotkeyText, desc)
                }
                for trigger, text in hotstrings {
                    if (InStr(trigger, searchText, false) || InStr(text, searchText, false)) {
                        shortText := SubStr(text, 1, 60) . (StrLen(text) > 60 ? "..." : "")
                        lv.Add(, trigger, shortText)
                    }
                }
                lv.ModifyCol(1, "AutoHdr")
                lv.ModifyCol(2, "AutoHdr")
            }
        }
        
        ResizeHelpHandler(gui, minMax, width, height) {
            try {
                if (minMax = -1)
                    return
                titleText.Move(, , width - 40)
                searchLabel.Move(, , width - 40)
                searchBox.Move(, , width - 40)
                lv.Move(, , width - 40, height - 110)
            }
        }
    } catch as e {
        ShowError("Hjelp dialog", e)
    }
}

; ============================================================================
; DATA MAPS - DOKUMENTASJON
; ============================================================================

GetHotkeysMap() {
    static hotkeyMap := ""
    if (hotkeyMap = "") {
        hotkeyMap := Map(
            "Ctrl + Space", "Fjerner mellomrom fra nummer i utklippstavlen",
            "Ctrl + Alt + T", "Lager rene telefonnumre med +47 foran",
            "Ctrl + Alt + B", "Søk etter væske varenummer (gammel → ny)",
            "Ctrl + Shift + Q", "Åpner rabatt-meny for å velge rabatt prosent",
            "Ctrl + Shift + B", "Nissan Bremsevæske (00000-01B00)",
            "Ctrl + Shift + F", "Nissan Frostvæske (00000-01F00)",
            "Ctrl + Shift + W", "(20% Rabatt Inkludert)",
            "Ctrl + Shift + M", "MDH Bestilt Levert",
            "Ctrl + Shift + A", "(eller bruk nøkkelautomat utenfor åpningstid)",
            "Ctrl + Shift + N", "(eller bruk nøkkelautomat utenfor åpningstid)",
            "Ctrl + Shift + S", "Serviceavtale",
            "Ctrl + Shift + D", "Åpner dekktilbud-meny",
            "Ctrl + Shift + H", "Viser denne hjelpeboksen",
            "Ctrl + Shift + R", "Starter scriptet på nytt",
            "Ctrl + Shift + X", "Venter på bilen",
            "Ctrl + Shift + P", "⚙️ Autofacet Setup Hub (konfigurer alle moduler)",
            "Ctrl + Shift + 1", "💾 LAGRE i Autofacet",
            "Ctrl + Shift + 2", "📅 PLANNER i Autofacet",
            "Ctrl + Shift + 3", "💬 KOMMUNIKASJON i Autofacet",
            "Ctrl + Shift + 4", "📋 HISTORIKK i Autofacet",
            "Ctrl + Shift + 5", "🔄 OPPDATERINGER i Autofacet",
            "Ctrl + Shift + |", "📝 ARBEIDSORDRE i Autofacet"  ;
        )
    }
    return hotkeyMap
}

GetHotstringsMap() {
    static hotstringMap := ""
    if (hotstringMap = "") {
        hotstringMap := Map(
            "*garanti", "Verkstedet vil undersøke det du opplever og om det er berettiget mot garanti...",
            "*info", "Bilen leveres kl 08:00 og hentes kl 15:30. Du kan også levere/hente utenfor åpningstid...",
            "*anbefalt", "Vi anbefaler årlig skift av vindusviskere foran (XXX,-), batteri i nøklene...",
            "*avdnr", "01 - Røyken, 02 - Rud, 03 - Oslo, 05 - Follo, 06 - Drammen, 07 - Lillestrøm, 08 - Gardermoen.",
            "*opsms-", "Hei, vi har forsøkt å ringe deg. Basert på våre opplysninger er det tid for service...",
            "*opsms+", "Hei, vi har forsøkt å ringe deg. Det er på tide med service på XXXXX...",
            "*samtykke", "Vi har ikke registrert ditt samtykke til elektronisk kommunikasjon...",
            "*forkontroll", "Vi ønsker å utføre en forkontroll av lyden først. På denne måten kan vi utelukke...",
            "*leiebil", "Du vil få tilsendt en SMS i forbindelse med din leiebilbestilling...",
            "*ikkevent", "OBS: Vi kan dessverre ikke tilby ventetime på denne type jobb..."
        )
    }
    return hotstringMap
}

; ============================================================================
; GDI+ FUNKSJONER - ✅ ORIGINAL FUNGERENDE VERSJON
; ============================================================================

global gdipToken := 0

if (!gdipToken)
    gdipToken := Gdip_Startup()

OnExit(GdipShutdownHandler)

GdipShutdownHandler(*) {
    global gdipToken
    if (gdipToken) {
        try {
            DllCall("gdiplus\GdiplusShutdown", "Ptr", gdipToken)
        }
        gdipToken := 0
    }
}

Gdip_Startup() {
    si := Buffer(24, 0)
    NumPut("UInt", 1, si)
    
    pToken := 0
    DllCall("gdiplus\GdiplusStartup", "Ptr*", &pToken, "Ptr", si, "Ptr", 0)
    return pToken
}

Gdip_BitmapFromScreen(Screen := 0) {
    if (Screen = 0) {
        screenWidth := DllCall("GetSystemMetrics", "Int", 0)
        screenHeight := DllCall("GetSystemMetrics", "Int", 1)
        x := 0
        y := 0
        w := screenWidth
        h := screenHeight
    } else {
        parts := StrSplit(Screen, "|")
        x := Integer(parts[1])
        y := Integer(parts[2])
        w := Integer(parts[3])
        h := Integer(parts[4])
    }
    
    hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
    hbm := CreateDIBSection(w, h)
    hdc2 := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
    obm := DllCall("SelectObject", "Ptr", hdc2, "Ptr", hbm, "Ptr")
    
    DllCall("BitBlt", "Ptr", hdc2, "Int", 0, "Int", 0, "Int", w, "Int", h, "Ptr", hdc, "Int", x, "Int", y, "UInt", 0x00CC0020)
    
    pBitmap := 0
    DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr", hbm, "Ptr", 0, "Ptr*", &pBitmap)
    
    DllCall("SelectObject", "Ptr", hdc2, "Ptr", obm)
    DllCall("DeleteObject", "Ptr", hbm)
    DllCall("DeleteDC", "Ptr", hdc2)
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
    
    return pBitmap
}

CreateDIBSection(w, h, bpp := 32) {
    hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
    
    bi := Buffer(40, 0)
    NumPut("UInt", 40, bi, 0)
    NumPut("Int", w, bi, 4)
    NumPut("Int", h, bi, 8)
    NumPut("UShort", 1, bi, 12)
    NumPut("UShort", bpp, bi, 14)
    
    hbm := DllCall("CreateDIBSection", "Ptr", hdc, "Ptr", bi, "UInt", 0, "Ptr*", &ppvBits := 0, "Ptr", 0, "UInt", 0, "Ptr")
    
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
    return hbm
}

Gdip_SaveBitmapToFile(pBitmap, sOutput) {
    ext := SubStr(sOutput, -3)
    
    if (ext = ".png")
        pCodec := GetCodecClsid("image/png")
    else if (ext = ".jpg" || ext = "jpeg")
        pCodec := GetCodecClsid("image/jpeg")
    else if (ext = ".bmp")
        pCodec := GetCodecClsid("image/bmp")
    else
        return -1
    
    wFile := Buffer(StrLen(sOutput) * 2 + 2)
    DllCall("MultiByteToWideChar", "UInt", 0, "UInt", 0, "Str", sOutput, "Int", -1, "Ptr", wFile, "Int", StrLen(sOutput) + 1)
    
    return DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "Ptr", wFile, "Ptr", pCodec, "Ptr", 0)
}

GetCodecClsid(mimeType) {
    static codecs := Map(
        "image/png", "{557CF406-1A04-11D3-9A73-0000F81EF32E}",
        "image/jpeg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}",
        "image/bmp", "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
    )
    
    clsid := Buffer(16)
    DllCall("ole32\CLSIDFromString", "WStr", codecs[mimeType], "Ptr", clsid)
    return clsid
}

Gdip_DisposeImage(pBitmap) {
    return DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
}

; ============================================================================
; TRAY MENU
; ============================================================================

A_TrayMenu.Delete()
A_TrayMenu.Add("&Hjelp (Ctrl+Shift+H)", (*) => ShowHelpDialog())
A_TrayMenu.Add("📊 &Statistikk", (*) => ShowStatsDialog())
A_TrayMenu.Add("&Væske-søk (Ctrl+Alt+B)", (*) => ShowFluidSearchDialog())
A_TrayMenu.Add()
A_TrayMenu.Add("🔄 Sjekk oppdateringer", (*) => CheckForUpdates())
A_TrayMenu.Add()
A_TrayMenu.Add("⚙️ &Autofacet Setup Hub (Ctrl+Shift+P)", (*) => Send("^+P"))
A_TrayMenu.Add("💾 LAGRE (Ctrl+Shift+1)", (*) => Send("^+1"))
A_TrayMenu.Add("📅 PLANNER (Ctrl+Shift+2)", (*) => Send("^+2"))
A_TrayMenu.Add("💬 KOMMUNIKASJON (Ctrl+Shift+3)", (*) => Send("^+3"))
A_TrayMenu.Add("📋 HISTORIKK (Ctrl+Shift+4)", (*) => Send("^+4"))
A_TrayMenu.Add("🔄 OPPDATERINGER (Ctrl+Shift+5)", (*) => Send("^+5"))
A_TrayMenu.Add("📝 ARBEIDSORDRE (Ctrl+Shift+|)", (*) => Send("^+|"))
A_TrayMenu.Add()
A_TrayMenu.Add("&Reload (Ctrl+Shift+R)", (*) => Reload())
A_TrayMenu.Add("&Avslutt", (*) => ExitApp())
A_TrayMenu.Default := "&Hjelp (Ctrl+Shift+H)"

; Startup melding
TrayTip("✅ BNH v" SCRIPT_VERSION " Blackbox Edition startet! Auto-update aktivert.", APP_TITLE, 0x1)


