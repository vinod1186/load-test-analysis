; ============================================================
; FINAL STABLE OFFICE LOAD-TIME SCRIPT
; Word + PowerPoint + Excel
; ============================================================

Global $CSV = @ScriptDir & "\LoadResults.csv"
Global $TIMEOUT = 15000 ; 15 seconds

; Create CSV with header only once
If Not FileExists($CSV) Then
    FileWrite($CSV, "Application,Run,LoadTime_ms" & @CRLF)
EndIf

; ============================
; RUN TESTS (10 runs each)
; ============================
For $i = 1 To 10

    ; WORD
    MeasureLoad("Word", "[CLASS:OpusApp]", $i)

    ; POWERPOINT (2 detection attempts)
    If Not MeasureLoad("PowerPoint", "[CLASS:PPTFrameClass]", $i) Then
        MeasureLoad("PowerPoint", "[CLASS:mdiClass]", $i)
    EndIf

    ; EXCEL
    MeasureLoad("Excel", "[CLASS:XLMAIN]", $i)

Next

MsgBox(64, "Completed", "Load time test finished!" & @CRLF & _
       "CSV saved at: " & @CRLF & $CSV)

Exit


; ============================================================
; FUNCTION: Measure Load Time
; Returns TRUE if success, FALSE if fail
; ============================================================
Func MeasureLoad($appName, $className, $runNum)

    Local $start = TimerInit()

    Local $hWnd = WinGetHandle($className)

    If @error Then
        WriteCSV($appName, $runNum, -1)
        Return False
    EndIf

    WinActivate($hWnd)
    Local $active = WinWaitActive($hWnd, "", $TIMEOUT / 1000)

    If $active = 0 Then
        WriteCSV($appName, $runNum, -1)
        Return False
    EndIf

    Local $elapsed = TimerDiff($start)
    WriteCSV($appName, $runNum, Int($elapsed))

    Return True
EndFunc


; ============================================================
; FUNCTION: Write to CSV
; ============================================================
Func WriteCSV($app, $run, $value)
    FileWrite($CSV, $app & "," & $run & "," & $value & @CRLF)
EndFunc
