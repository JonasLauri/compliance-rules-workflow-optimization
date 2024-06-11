; STANDART COMMENTS FOR MONDAY TEMP. BR PROCESS
; COMMENTS WHEN PEP/PRIMARY PEP ACTIVE

<^Numpad1::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because PEP remains in same active position`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return

<^!Numpad1::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because primary PEP remains in same active position`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return


; COMMENTS WHEN PEP/PRIMARY PEP ALREADY FORMER

<^Numpad2::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because 18 months have not passed since PEP is considered former`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return

<^!Numpad2::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because 18 months have not passed since primary PEP is considered former`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return


; COMMENTS WHEN PEP/PRIMARY PEP FORMER

<^Numpad3::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because PEP has already been considered former since ________`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return


<^!Numpad3::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because primary PEP has already been considered former since ________`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return


; COMMENT WHEN >1 RELATIONS ARE IDENTIFIED 

<^Numpad4::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because primary PEP {Name} remains in same active position/has not yet considered former as less than 18 months have passed since end of last position/has already been considered former since _____ and other identified primary PEP {Name} remains in same active position/has not yet considered former as less than 18 months have passed since end of last position/has already been considered former since _____`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return

; COMMENT WHEN POSITION NOT CONSIDERED IN ALL MARKET AREAS

<^!Numpad6::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
This is an exception. However, according to provided guidelines from SOP: Customer PEP Screening and PEP definition guidelines this position is not considered as a PEP in all market areas. Due to this rule will be added to rule out matches from screening`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return


; COMMENT WHEN NOTHING HAS CHANGED

<^Numpad0::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
However, these changes have no influence on the final decision because changed fields in PEP profile are irrelevant and not taken into consideration during investigation as entity level is decided based on requirements described in Customer PEP Screening process documentation`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return



; TO BE CONTINUE .. ONE MORE CMT CAN BE ADDED > COMMENT WHEN THERE ARE NO CHANGES MADE IN THE RELEVANT FIELDS OF PEP PROFILE
; UPDATED 2022-11-08:
; 	HOTKEY <^Numpad0 ADDED
;	HOTKEY <^Numpad4 COMMENT ADJUSTED - ADDED "has not yet considered former as less than 18 months have passed since end of last position"
; UPDATED 2023-11-13:
; NEW COMMENT ADDED - WHEN POSITION NOT CONSIDERED IN ALL MARKET AREAS