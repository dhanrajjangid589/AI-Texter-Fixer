#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ========== CONFIG ==========
; Your keys (OpenAI primary, Gemini fallback)
global OPENAI_API_KEY := "<Enter your KEY>"
global OPENAI_MODEL   := "gpt-4o-mini"

global GEMINI_API_KEY := "Enter Your Key" ; leave as-is if you want fallback
global GEMINI_MODEL   := "gemini-1.5-flash"

; Behavior
global TEMP_DEFAULT   := 0.6
global TEMP_FORCE     := 0.9
global MAX_CHARS      := 8000         ; safe cap
global AUTO_SELECT_ALL := true        ; if no selection detected, press Ctrl+A automatically

TrayTip("AI Text Fixer V2", "Loaded. Ctrl+Alt+R rephrases, Ctrl+Alt+F fixes")

; ========== HOTKEYS ==========
^!r:: HandleAction("rephrase")   ; Ctrl+Alt+R
^!f:: HandleAction("fix")        ; Ctrl+Alt+F

; ========== CORE ==========
HandleAction(mode) {
    sel := CaptureTextAuto()
    if (!sel.ok)
        return

    instr := BuildInstruction(mode)
    resp := TryRewrite(instr, sel.text, TEMP_DEFAULT)

    if (!resp.ok) {
        MsgBox "ERR (" resp.status "): " resp.message
        if (sel.wasCut)
            PasteReplace(sel.text)  ; restore original if we cut
        return
    }

    ; If rephrase didn't change much, try stronger, then fix
    if (mode = "rephrase" && IsEffectivelySame(sel.text, resp.text)) {
        force := "Rewrite so it is NOT identical to the input. Fix mistakes and improve clarity. Make at least one wording change. Return only the rewritten text."
        resp2 := TryRewrite(force, sel.text, TEMP_FORCE)
        if (resp2.ok && !IsEffectivelySame(sel.text, resp2.text)) {
            PasteReplace(resp2.text)
            return
        }
        fix := BuildInstruction("fix")
        resp3 := TryRewrite(fix, sel.text, 0.3)
        if (resp3.ok && !IsEffectivelySame(sel.text, resp3.text)) {
            PasteReplace(resp3.text)
            return
        }
        if (sel.wasCut)
            PasteReplace(sel.text)
        TrayTip("AI Text Fixer V2", "No change. Try selecting a longer passage.")
        return
    }

    PasteReplace(resp.text)
}

BuildInstruction(mode) {
    base := ""
    if (mode = "fix") {
        base := "Correct spelling, grammar, and punctuation. Preserve meaning. Return only the corrected text."
    } else if (mode = "rephrase") {
        base := "Rewrite this text in a clearer, more natural way. Correct mistakes. Provide an alternative phrasing not identical to the input. Return only the rewritten text."
    } else {
        base := "Act as a concise copy editor. Return only the edited text."
    }
    return base
}

; ========== AUTO CAPTURE / REPLACE ==========
CaptureTextAuto() {
    sel := { ok:false, wasCut:false, usedSelectAll:false, text:"" }

    old := A_Clipboard
    A_Clipboard := ""

    ; Try copy current selection
    Send "^c"
    ClipWait(0.5)
    txt := A_Clipboard

    ; If nothing, optionally Ctrl+A to select all
    if ((txt = "" || StrLen(Trim(txt)) = 0) && AUTO_SELECT_ALL) {
        A_Clipboard := ""
        Send "^a"
        Sleep 40
        Send "^c"
        ClipWait(0.5)
        txt := A_Clipboard
        if (txt != "" && StrLen(Trim(txt)) > 0)
            sel.usedSelectAll := true
    }

    ; If still nothing, try cut (some Electron apps behave better)
    if (txt = "" || StrLen(Trim(txt)) = 0) {
        A_Clipboard := ""
        Sleep 30
        Send "^x"
        ClipWait(0.6)
        txt := A_Clipboard
        if (txt != "" && StrLen(Trim(txt)) > 0)
            sel.wasCut := true
    }

    A_Clipboard := old

    if (txt = "" || StrLen(Trim(txt)) = 0) {
        MsgBox "No text found. Click the input and try again."
        return sel
    }
    if (StrLen(txt) > MAX_CHARS) {
        MsgBox "Selection too large (> " . MAX_CHARS . ")."
        if (sel.wasCut)
            PasteReplace(txt)
        return sel
    }

    sel.ok := true
    sel.text := txt
    return sel
}

PasteReplace(text) {
    old := A_Clipboard
    A_Clipboard := text
    Sleep 30
    Send "^v"
    Sleep 70
    A_Clipboard := old
}

; ========== ROUTER (OpenAI → Gemini fallback) ==========
TryRewrite(instruction, userText, temperature) {
    if (OPENAI_API_KEY && SubStr(OPENAI_API_KEY,1,3)="sk-") {
        r := OpenAI_Rewrite(instruction, userText, temperature)
        if (r.ok)
            return r
        if (r.status=429 && InStr(StrLower(r.message), "insufficient_quota")) {
            if (GEMINI_API_KEY) {
                g := Gemini_Rewrite(instruction, userText, temperature)
                if (g.ok)
                    return g
                return g
            } else {
                r.message := r.message . "`n`nTip: Project has no API credit. Either add billing or set GEMINI_API_KEY for fallback."
                return r
            }
        }
        return r
    } else if (GEMINI_API_KEY) {
        return Gemini_Rewrite(instruction, userText, temperature)
    } else {
        return { ok:false, status:-1, message:"No provider configured. Add OpenAI billing or set GEMINI_API_KEY." }
    }
}

; ========== OPENAI HTTP ==========
OpenAI_Rewrite(instruction, userText, temperature := 0.7) {
    url := "https://api.openai.com/v1/chat/completions"
    payload := BuildChatJson(OPENAI_MODEL, instruction, userText, temperature)

    req := ComObject("WinHttp.WinHttpRequest.5.1")
    try {
        req.Open("POST", url, false)
        req.SetRequestHeader("Content-Type", "application/json")
        req.SetRequestHeader("Authorization", "Bearer " . OPENAI_API_KEY)
        req.Send(payload)
        status := req.Status
        body := req.ResponseText
    } catch {
        return { ok:false, status:-1, message:"HTTP error (OpenAI open/send)" }
    }

    if (status != 200)
        return { ok:false, status:status, message:body }

    text := ExtractAssistantContent(body)
    if (!text)
        return { ok:false, status:200, message:"Parse error. Raw: " . SubStr(body,1,500) }

    text := JsonUnescape(text)
    return { ok:true, status:200, text:text }
}

; ========== GEMINI HTTP (fallback) ==========
Gemini_Rewrite(instruction, userText, temperature := 0.7) {
    url := "https://generativelanguage.googleapis.com/v1beta/models/" . GEMINI_MODEL . ":generateContent?key=" . GEMINI_API_KEY

    msg := instruction . "`n`nText:`n" . userText
    msgEsc := JEscape(msg)

    q := Chr(34)
    payload := "{"
    payload .= q "contents" q ":[{"
    payload .= q "parts" q ":[{" q "text" q ":" q msgEsc q "}]}],"
    payload .= q "generationConfig" q ":{"
    payload .= q "temperature" q ":" temperature
    payload .= "}"
    payload .= "}"

    req := ComObject("WinHttp.WinHttpRequest.5.1")
    try {
        req.Open("POST", url, false)
        req.SetRequestHeader("Content-Type", "application/json")
        req.Send(payload)
        status := req.Status
        body := req.ResponseText
    } catch {
        return { ok:false, status:-1, message:"HTTP error (Gemini open/send)" }
    }

    if (status != 200)
        return { ok:false, status:status, message:body }

    text := GeminiExtract(body)
    if (!text)
        return { ok:false, status:200, message:"Gemini parse error. Raw: " . SubStr(body,1,500) }

    text := JsonUnescape(text)
    return { ok:true, status:200, text:text }
}

GeminiExtract(json) {
    m := 0
    if RegExMatch(json, 's)"candidates"\s*:\s*\[\s*\{.*?"content"\s*:\s*\{.*?"parts"\s*:\s*\[\s*\{.*?"text"\s*:\s*"([^"]*)"', &m)
        return m[1]
    return ""
}

; ========== JSON BUILDING / PARSING HELPERS ==========
BuildChatJson(model, instruction, userText, temperature) {
    q := Chr(34)
    msg := instruction . "`n`nText:`n" . userText
    msgEsc := JEscape(msg)

    json := "{"
    json .= q "model" q ":" q model q ","
    json .= q "messages" q ":["
        json .= "{"
        json .= q "role" q ":" q "system" q ","
        json .= q "content" q ":" q "You are a precise text editor. Return only the edited text." q
        json .= "},"
        json .= "{"
        json .= q "role" q ":" q "user" q ","
        json .= q "content" q ":" q msgEsc q
        json .= "}"
    json .= "],"
    json .= q "temperature" q ":" temperature
    json .= "}"
    return json
}

JEscape(s) {
    q := Chr(34)
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, q, "\" q)
    s := StrReplace(s, "`r", "")
    s := StrReplace(s, "`n", "\n")
    return s
}
JsonUnescape(s) {
    s := StrReplace(s, "\r\n", "`n")
    s := StrReplace(s, "\n", "`n")
    s := StrReplace(s, '\"', '"')
    s := StrReplace(s, "\\\\", "\")
    return s
}

ExtractAssistantContent(json) {
    m := 0
    if RegExMatch(json, 's)"choices"\s*:\s*\[\s*\{.*?"message"\s*:\s*\{.*?"content"\s*:\s*"([^"]*)"', &m)
        return m[1]
    if RegExMatch(json, 's)"content"\s*:\s*"([^"]*)"', &m)
        return m[1]
    return ""
}

; ========== “SAME TEXT” CHECK ==========
IsEffectivelySame(a, b) {
    na := NormalizeText(a), nb := NormalizeText(b)
    return (na = nb)
}
NormalizeText(s) {
    s := Trim(s)
    s := RegExReplace(s, "\s+", " ")
    s := RegExReplace(s, "[\.\!\?]+$", "")
    return StrLower(s)
}
