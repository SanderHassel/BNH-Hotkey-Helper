~y Up::
~+y Up:: {
    static lastYTime := 0
    static yCount := 0
    currentTime := A_TickCount
    
    if !GetKeyState("Shift", "P") {
        lastYTime := 0
        yCount := 0
        return
    }
    
    if (currentTime - lastYTime > 300) {
        yCount := 1
        lastYTime := currentTime
        return
    }
    
    if (yCount = 1 && (currentTime - lastYTime) <= 300) {
        yCount := 0
        lastYTime := 0
        
        Send("{Backspace 2}")
        ShowQuickTilbudYLoopDialog()
        return
    }
    
    lastYTime := currentTime
    yCount := 1
}

ShowQuickTilbudYLoopDialog() {
    ; Show GUI asking for loop count
    ; Call ExecuteAutofacetQuickTilbudY(loopCount) when user confirms
}

ExecuteAutofacetQuickTilbudY(loopCount) {
    ; Validate browser
    ; Read all 8 points + sleep times from INI
    ; Execute loop logic with point 8 conditional
}

SetupQuickTilbudYPoints() {
    ; Setup wizard for 8 points
    ; For each point: ask for position AND sleep time (in ms)
    ; Save to autofacet_config.ini
}

; Configuration for QUICKTILBUDY
QUICKTILBUDY_POINT1 := { X: 123, Y: 456, SLEEP: 500 }
QUICKTILBUDY_POINT2 := { X: 789, Y: 012, SLEEP: 300 }
; ... to QUICKTILBUDY_POINT8

; Update SCRIPT_VERSION
SCRIPT_VERSION := "6.3.0"