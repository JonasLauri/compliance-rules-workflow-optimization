; STANDART COMMENTS FOR FRIDAY TEMP. BR PROCESS
; COMMENTS WHEN PEP/PRIMARY PEP ACTIVE

<^Numpad1::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
PEP remains in same active position, according to Online Compliance `n
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
primary PEP remains in same active position, according to Online Compliance `n
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
PEP has not yet considered former as less than 18 months have passed since end of last position, according to Online Compliance `n
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
18 months have not passed since primary PEP is considered former, according to Online Compliance `n
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
PEP has already been considered former since ________ , according to Online Compliance `n
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
primary PEP has already been considered former since ________ , according to Online Compliance `n
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
primary {Name} _____ and other identified primary {Name} _____
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return


; COMMENT WHEN YOB EXCEPTION APPLY

<^Numpad5::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
This is an exception. However, according to Online Compliance, PEP was born in _____ , so, additional exception for YOB _____ will be added & in order to mitigate the risk of not identifying the true match in Compliance Link`n
)
ClipWait, 2              ; wait max. 2 seconds for the clipboard to contain data. 
if (!ErrorLevel)         ; If NOT ErrorLevel, ClipWait found data on the clipboard
    Send, ^v             ; paste the text
Sleep, 300
clipboard := ClipSaved   ; restore original clipboard
ClipSaved =              ; Free the memory in case the clipboard was very large.
return


; COMMENT WHEN COUNTRY EXCEPTION APPLY

<^Numpad6::
ClipSaved := ClipboardAll ; save the entire clipboard to the variable ClipSaved
clipboard := ""           ; empty the clipboard (start off empty to allow ClipWait to detect when the text has arrived)
clipboard =               ; copy this text:
(
This is an exception. However, according to provided guidelines from SOP: Customer PEP Screening and PEP definition guidelines this position is not considered as a PEP in _____ . Due to this additional rule will be added to rule out matches from _____`n
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
However, according to provided guidelines from SOP: Customer PEP Screening and PEP definition guidelines this position is not considered as a PEP in all market areas. Due to this rule will be added to rule out matches from screening`n
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
changed fields in PEP profile are irrelevant and not taken into consideration during investigation as entity level is decided based on requirements described in Customer PEP Screening process documentation`n
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
;	HOTKEY <^Numpad0 ADDED
;	HOTKEYS <^Numpad5, <^Numpad6 COMMENT SUPPLEMENTED - ADDED "This is an exception."
;	HOTKEY <^Numpad2 COMMENT ADJUSTED - ADDED "PEP has not yet considered former as less than 18 months have passed since end of last position"
; UPDATED 2023-11-13:
; NEW COMMENT ADDED - WHEN POSITION NOT CONSIDERED IN ALL MARKET AREAS