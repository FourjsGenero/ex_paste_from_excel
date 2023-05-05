IMPORT reflect

IMPORT FGL paste_from_excel

MAIN
    DEFINE arr DYNAMIC ARRAY OF RECORD
        field1 INTEGER,
        field2 STRING,
        field3 DATE
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

    INPUT ARRAY arr FROM scr.* ATTRIBUTES(WITHOUT DEFAULTS = TRUE, UNBUFFERED)
        ON ACTION clear ATTRIBUTES(TEXT = "Clear", VALIDATE = NO)
            CALL arr.clear()

        ON ACTION populate ATTRIBUTES(TEXT = "Populate", VALIDATE = NO)
            CALL arr.clear()
            FOR i = 1 TO 10
                LET arr[i].field1 = i
                LET arr[i].field2 = ASCII (64 + i), ASCII (96 + i), ASCII (96 + i)
                LET arr[i].field3 = TODAY + i
            END FOR

        ON ACTION pastefromexcel ATTRIBUTES(TEXT = "Paste Special (from Excel)", VALIDATE = NO)

            CALL paste_from_excel.paste(reflect.Value.valueOf(arr)) RETURNING ok, error_text
            IF NOT ok THEN
                CALL FGL_WINMESSAGE("Error", error_text, "stop")
            END IF
    END INPUT

END MAIN
