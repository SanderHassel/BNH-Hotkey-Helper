; ============================================================================
; BNH QUICK SMS MODULE v1.0
; Separat modul for Quick SMS-funksjonalitet
; Inkluderes av BNH_Hotkey_V5_DEMO.ahk via #Include
; ============================================================================

ExecuteAutofacetQuickSMS() {
    try {
        TrackUsage("Autofacet Quick SMS")
        
        ; Sjekk om vi er i riktig nettleser
        if !WinActive("ahk_exe chrome.exe") && !WinActive("ahk_exe msedge.exe") && !WinActive("ahk_exe brave.exe") {
            ShowQuietNotification("⚠️ Denne funksjonen fungerer kun i Chrome/Edge/Brave")
            return
        }
        
        ; Sjekk om punktene er konfigurert
        configFile := A_ScriptDir "\autofacet_config.ini"
        
        if !FileExist(configFile) {
            MsgBox("❌ Quick SMS er ikke konfigurert!`n`nTrykk Ctrl+Shift+P og klikk på QUICKSMS-kortet for å konfigurere.", "Quick SMS", "Icon!")
            return
        }
        
        ; Les punkt 1 for å verifisere at QUICKSMS er satt opp
        p1x := IniRead(configFile, "QUICKSMS_POINT1", "X", "")
        p1y := IniRead(configFile, "QUICKSMS_POINT1", "Y", "")
        
        if (p1x = "" || p1y = "") {
            MsgBox("❌ Quick SMS-punkter er ikke konfigurert!`n`nTrykk Ctrl+Shift+P og klikk på QUICKSMS-kortet for å konfigurere.", "Quick SMS", "Icon!")
            return
        }
        
        ; Vis popup-meny for å velge SMS-mal
        ShowQuickSMSMenu()
        
    } catch as e {
        ShowError("Quick SMS", e)
    }
}

ShowQuickSMSMenu() {
    try {
        smsGui := Gui("+AlwaysOnTop", "📱 Quick SMS - Velg mal")
        smsGui.BackColor := COLORS.BG_DARK
        smsGui.MarginX := 20
        smsGui.MarginY := 20
        
        ; Header
        titleText := smsGui.Add("Text", "w400 h40 Center c" COLORS.TEXT_WHITE, "📱 QUICK SMS")
        titleText.SetFont("s16 Bold", "Segoe UI")
        
        infoText := smsGui.Add("Text", "w400 h25 Center c" COLORS.TEXT_GRAY, "Velg en SMS-mal eller skriv egen melding:")
        infoText.SetFont("s9", "Segoe UI")
        
        ; SMS-maler som knapper
        smsTemplates := [
            {name: "📋 Service -Avtale",   text: "Hei, vi har forsøkt å ringe deg. Basert på våre opplysninger er det tid for service på din bil. Bestill time raskt og enkelt på nett: https://service.bnh.no/ Hilsen Birger N. Haug / 40010400."},
            {name: "📋 Service +Avtale",   text: "Hei, vi har forsøkt å ringe deg. Det er på tide med service på bilen din. Du har allerede en forhåndsbetalt serviceavtale. Bestill time her: https://service.bnh.no/ eller ring oss på 40010400. Hilsen Birger N. Haug."},
            {name: "🔧 Service Garanti",   text: "Hei, vi har forsøkt å ringe deg. Basert på våre opplysninger er det tid for service på bilen din. Ring oss gjerne tilbake på 40010400 slik at vi kan sette opp en time. Hilsen Birger N. Haug."},
            {name: "🚗 EU-kontroll",       text: "Hei, vi har forsøkt å ringe deg. Basert på våre opplysninger er det tid for EU-kontroll på bilen din. Bestill time her: https://service.bnh.no/ eller ring oss på 40010400. Hilsen Birger N. Haug."},
            {name: "📞 Ring oss tilbake",  text: "Hei! Vi har forsøkt å ringe deg angående din bil. Vennligst ring oss tilbake på 40010400. Hilsen Birger N. Haug."}
        ]
        
        for idx, template in smsTemplates {
            btn := CreateStyledButton(smsGui, "w400 h35", template.name, COLORS.BG_MEDIUM, 10)
            templateText := template.text
            btn.OnEvent("Click", ((t) => (*) => SelectSMSTemplate(t))(templateText))
        }
        
        ; Separator
        smsGui.Add("Text", "w400 h2 y+10 Background" COLORS.TEXT_GRAY)
        
        ; Egendefinert melding
        customLabel := smsGui.Add("Text", "w400 h20 y+10 c" COLORS.TEXT_GRAY, "Eller skriv egen melding:")
        customLabel.SetFont("s9", "Segoe UI")
        
        customInput := smsGui.Add("Edit", "w400 h60 c" COLORS.TEXT_WHITE " Background" COLORS.BG_MEDIUM " Multi")
        customInput.SetFont("s10", "Segoe UI")
        
        ; Knapper
        sendCustomBtn := CreateStyledButton(smsGui, "w195 h40 y+10", "📤 Send egendefinert", COLORS.GREEN, 10)
        sendCustomBtn.OnEvent("Click", (*) => SendCustomSMS())
        
        cancelBtn := CreateStyledButton(smsGui, "x+10 w195 h40 yp", "❌ Avbryt", COLORS.RED, 10)
        cancelBtn.OnEvent("Click", (*) => smsGui.Destroy())
        
        smsGui.OnEvent("Close", (*) => smsGui.Destroy())
        smsGui.OnEvent("Escape", (*) => smsGui.Destroy())
        smsGui.Show("w440")
        
        SelectSMSTemplate(templateText) {
            smsGui.Destroy()
            Sleep(300)
            RunQuickSMSSequence(templateText)
        }
        
        SendCustomSMS() {
            customText := customInput.Value
            if (Trim(customText) = "") {
                MsgBox("Skriv inn en melding først!", "Quick SMS", "Icon!")
                return
            }
            smsGui.Destroy()
            Sleep(300)
            RunQuickSMSSequence(customText)
        }
        
    } catch as e {
        ShowError("Quick SMS Menu", e)
    }
}

RunQuickSMSSequence(smsText) {
    try {
        configFile := A_ScriptDir "\autofacet_config.ini"
        
        ; Les alle konfigurerte punkter
        points := []
        pointCount := 0
        
        Loop 4 {
            px := IniRead(configFile, "QUICKSMS_POINT" A_Index, "X", "")
            py := IniRead(configFile, "QUICKSMS_POINT" A_Index, "Y", "")
            
            if (px != "" && py != "") {
                points.Push({x: Integer(px), y: Integer(py)})
                pointCount++
            }
        }
        
        if (pointCount = 0) {
            MsgBox("❌ Ingen Quick SMS-punkter konfigurert!`n`nGå til Setup Hub (Ctrl+Shift+P) og konfigurer QUICKSMS.", "Quick SMS", "Icon!")
            return
        }
        
        ; Lagre original musposisjon
        MouseGetPos(&origX, &origY)
        
        ; Utfør klikk-sekvensen
        Loop pointCount {
            idx := A_Index
            MouseMove(points[idx].x, points[idx].y, 0)
            Sleep(100)
            Click("Left")
            
            if (idx < pointCount)
                Sleep(500)
        }
        
        ; Vent litt, deretter paste SMS-teksten
        Sleep(500)
        
        ; Paste meldingen
        prevClip := A_Clipboard
        A_Clipboard := smsText
        Sleep(100)
        Send("^v")
        Sleep(300)
        
        ; Gjenopprett original clipboard
        A_Clipboard := prevClip
        
        ; Flytt musen tilbake
        MouseMove(origX, origY, 0)
        
        ShowQuietNotification("✅ SMS-tekst limt inn!")
        
    } catch as e {
        ShowError("Quick SMS Sequence", e)
    }
}

SetupQuickSMSPoints() {
    try {
        setupText := "📱 Konfigurer Quick SMS (4 punkter):`n`n"
        setupText .= "Du må konfigurere 4 klikk-punkter:`n`n"
        setupText .= "PUNKT 1: Første klikk`n"
        setupText .= "PUNKT 2: Andre klikk`n"
        setupText .= "PUNKT 3: Tredje klikk`n"
        setupText .= "PUNKT 4: Fjerde klikk (før meny)`n`n"
        setupText .= "Trykk OK for å starte"
        
        result := MsgBox(setupText, "Setup: Quick SMS", "OKCancel Icon!")
        if (result = "Cancel")
            return
        
        configFile := A_ScriptDir "\autofacet_config.ini"
        points := ["PUNKT 1", "PUNKT 2", "PUNKT 3", "PUNKT 4"]
        sections := ["QUICKSMS_POINT1", "QUICKSMS_POINT2", "QUICKSMS_POINT3", "QUICKSMS_POINT4"]
        
        Loop 4 {
            idx := A_Index
            MsgBox("Klar for " points[idx] "`n`nDu har 5 sekunder til å plassere musen.", points[idx], "Iconi T3")
            
            Loop 5 {
                remaining := 6 - A_Index
                ToolTip("Lagrer " points[idx] " om " remaining " sekunder...`n`nHold musen stille!", A_ScreenWidth/2, A_ScreenHeight/2)
                Sleep(1000)
            }
            ToolTip()
            
            MouseGetPos(&mx, &my)
            IniWrite(mx, configFile, sections[idx], "X")
            IniWrite(my, configFile, sections[idx], "Y")
            
            MsgBox("✅ " points[idx] " lagret!`n`nX: " mx "`nY: " my, "Suksess", "Iconi T2")
        }
        
        MsgBox("🎉 Quick SMS fullstendig konfigurert!`n`nDobbel-tap Ctrl for å bruke.", "Ferdig", "Iconi T4")
        
    } catch as e {
        ShowError("Setup Quick SMS", e)
    }
}
