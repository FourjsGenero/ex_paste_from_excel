# ex_paste_from_excel
Example showing how to use reflection to take clipboard with array data and append into array variable

Uses reflection to read Clipboard and paste into a 4gl array

Intended Usage is in an INPUT ARRAY similar to ...

```
    INPUT ARRAY arr ...
    
        ON ACTION pastefromexcel ATTRIBUTES(VALIDATE=NO, TEXT="Paste Special (from Excel)")
            DISPLAY arr.getLength()
            CALL paste_from_excel(reflect.Value.valueOf(arr)) RETURNING ok, error_text
            IF NOT ok THEN
                CALL FGL_WINMESSAGE("Error", error_text,"stop")
            END IF
```            

If the paste was succesful, ok will be TRUE
If the paste was unsuccesful, ok will be FALSE, and an error message will be in error_text

Assumes clipboard has data that has lines/rows seperated by ASCII(10) LF, and cells within the line/row seperate by ASCII(9) TAB

To test a succesful paste, create an Excel spreadsheet with the following values.  Select the 6 cells and Copy (Control-C), then in the program select Paste Special (from Excel)

|A|B  |C       |
|-|---|--------|
|1|AAA|1/1/2024|
|2|BBB|2/2/2024|

To test failure conditions, use the following, should get an error message "Number of elements in clipboard does not match number of elements in array"

|A|B  |
|-|---|
|1|AAA|
|2|BBB|


|A|B  |C       |D|
|-|---|--------|-|
|1|AAA|1/1/2024|1|
|2|BBB|2/2/2024|2|

Also note the result of the following test.  The number can do into the string field, whilst the text cna't go into the number field.  A future enhancement would be to test and cater for this.

|A  |B|C       |
|---|-|--------|
AAA|1|1/1/2024
BBB|2|2/2/2024

The data will be appended to the end of the array.  If the last row of the array is blank, it will start on that line.  If the last row the array contains data, it will start on the next line


