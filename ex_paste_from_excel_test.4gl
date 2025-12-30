IMPORT reflect

IMPORT FGL paste_from_excel

MAIN
    DEFINE arr1 DYNAMIC ARRAY OF RECORD
        field11 INTEGER,
        field12 STRING,
        field13 DATE
    END RECORD
    DEFINE arr2 DYNAMIC ARRAY OF RECORD
        field21 INTEGER,
        field22 STRING
    END RECORD
    DEFINE i INTEGER
    DEFINE ok BOOLEAN
    DEFINE error_text STRING

    WHENEVER ANY ERROR STOP
    DEFER INTERRUPT
    DEFER QUIT
    OPTIONS FIELD ORDER FORM
    OPTIONS INPUT WRAP

    CLOSE WINDOW SCREEN

    OPEN WINDOW w WITH FORM "ex_paste_from_excel_test" ATTRIBUTES(TEXT = "Paste From Excel Test")

    DIALOG ATTRIBUTES(UNBUFFERED)
        INPUT ARRAY arr1 FROM scr1.* ATTRIBUTES(WITHOUT DEFAULTS = TRUE)

            ON ACTION pastefromexcel ATTRIBUTES(TEXT = "Paste Special (from Excel)", VALIDATE = NO)

                CALL paste_from_excel.paste(reflect.Value.valueOf(arr1)) RETURNING ok, error_text
                IF NOT ok THEN
                    CALL FGL_WINMESSAGE("Error", error_text, "stop")
                END IF

            ON ACTION selectall ATTRIBUTES(TEXT = "Select All To Copy", VALIDATE = NO)
                DISPLAY ARRAY arr1 TO scr1.*
                    BEFORE DISPLAY
                        CALL DIALOG.setSelectionMode("scr1", 1)
                        CALL DIALOG.setSelectionRange("scr1", 1, -1, TRUE)
                END DISPLAY
                LET int_flag = 0
        END INPUT
        INPUT ARRAY arr2 FROM scr2.* ATTRIBUTES(WITHOUT DEFAULTS = TRUE)

            ON ACTION pastefromexcel ATTRIBUTES(TEXT = "Paste Special (from Excel)", VALIDATE = NO)

                CALL paste_from_excel.paste(reflect.Value.valueOf(arr2)) RETURNING ok, error_text
                IF NOT ok THEN
                    CALL FGL_WINMESSAGE("Error", error_text, "stop")
                END IF

            ON ACTION selectall ATTRIBUTES(TEXT = "Select All To Copy", VALIDATE = NO)
                DISPLAY ARRAY arr2 TO scr2.*
                    BEFORE DISPLAY
                        CALL DIALOG.setSelectionMode("scr2", 1)
                        CALL DIALOG.setSelectionRange("scr2", 1, -1, TRUE)
                END DISPLAY
                LET int_flag = 0

            ON ACTION dialogtouched
                CALL ui.Interface.frontCall("standard","cbGet",[],error_text)
                DISPLAY ORD(error_text.subString(1,error_text.getLength()))
             
        END INPUT

        ON ACTION clear ATTRIBUTES(TEXT = "Clear", VALIDATE = NO)
            CALL arr1.clear()
            CALL arr2.clear()

        ON ACTION populate ATTRIBUTES(TEXT = "Populate", VALIDATE = NO)
            CALL arr1.clear()
            FOR i = 1 TO 10
                LET arr1[i].field11 = i
                LET arr1[i].field12 = ASCII (64 + i), ASCII (96 + i), ASCII (96 + i)
                LET arr1[i].field13 = TODAY + i
            END FOR
            CALL arr2.clear()
            FOR i = 1 TO 10
                LET arr2[i].field21 = i * 100
                LET arr2[i].field22 = ASCII (96 + i), "00"
            END FOR

        ON ACTION CLOSE
            EXIT DIALOG
    END DIALOG

END MAIN
