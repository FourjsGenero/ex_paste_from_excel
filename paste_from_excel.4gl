IMPORT reflect

FUNCTION paste(reflect_array reflect.Value)
DEFINE reflect_row, reflect_cell reflect.Value
DEFINE str, line, cell STRING
DEFINE tok_line, tok_cell base.StringTokenizer
DEFINE idx INTEGER
DEFINE first BOOLEAN = TRUE
DEFINE current_array_length INTEGER
DEFINE auto_append_row BOOLEAN = FALSE

    -- Note length of array in case need to restore to this length
    LET current_array_length = reflect_array.getLength()

    -- test if last row has values.
    -- if it is all blank then it is an auto append row and we can paste into ot
    -- if not then we will start the paste on the next row
    LET auto_append_row = TRUE
    LET reflect_row = reflect_array.getArrayElement(current_array_length)
    FOR idx = 1 TO reflect_row.getType().getFieldCount()
        LET reflect_cell = reflect_row.getField(idx)
        IF reflect_cell.toString().getLength() > 0 THEN
            LET auto_append_row = FALSE
            EXIT FOR
        END IF
    END FOR

    -- get the value form the clipboard
    -- not for Firefox users e # https://developer.mozilla.org/en-US/docs/Web/API/Clipboard_API
    CALL ui.Interface.frontCall("standard","cbget",[], str)

    -- Split line by use of ASCII(10) (LF).  
    -- Note: does not cater for use of ASCII(10) inside quoted string
    LET tok_line = base.StringTokenizer.createExt(str,ASCII(10), "\\", TRUE)
    
    WHILE tok_line.hasMoreTokens()
        -- get next line
        LET line = tok_line.nextToken()

        -- Add a row at end of array
        IF first AND auto_append_row THEN
            # don't add row
        ELSE
            # add row
            CALL reflect_array.appendArrayElement()
        END IF
        LET reflect_row = reflect_array.getArrayElement( reflect_array.getLength())

        -- First time through, check that number of fields in clipboard line equals number of elements in array row
        IF first THEN
            LET tok_cell = base.StringTokenizer.createExt(line,ASCII(9), "\\", TRUE)
            VAR tok_count = tok_cell.countTokens()
            VAR field_count = reflect_row.getType().getFieldCount()
            IF tok_count = field_count THEN
                #OK
            ELSE
                IF NOT auto_append_row THEN
                    CALL reflect_array.deleteArrayElement(reflect_array.getLength()) -- delete element that has been added
                END IF
                RETURN FALSE, "Number of elements in clipboard does not match number of elements in array"
            END IF
            LET first = FALSE
        END IF
        
        -- Split line into cells by use of ASCII(9) (TAB).  
        -- Note: does not cater for use of ASCII(9) inside quoted string
        LET tok_cell = base.StringTokenizer.createExt(line,ASCII(9), "\\", TRUE)
        LET idx = 0
        WHILE tok_cell.hasMoreTokens()
            -- Get the next value in clipboard line and assign it to next element in array row
            LET cell = tok_cell.nextToken()
            LET idx = idx + 1
            LET reflect_cell = reflect_row.getField(idx)
            CALL reflect_cell.set(reflect.Value.copyOf(cell))

            #TODO add test if could not assign value
            #TODO this test thinks that 01/01/2023 is different from 1/1/2023 so isnot perfect and probably needs to be done a different way
            #IF reflect.Value.copyOf(cell).toString() = reflect_cell.toString() THEN
            #    #OK
            #ELSE
            #    DISPLAY reflect.Value.copyOf(cell).toString()
            #   DISPLAY reflect_cell.toString()
            #    RETURN FALSE, "Unable to assign this datatype"
            #END IF
        END WHILE
    END WHILE
    RETURN TRUE, ""
END FUNCTION