# :warning: DEPRECATED :warning:

# ex_paste_from_excel

Note: this example was written before the ability to paste a range of cells from Excel into a GBC application was reintroduced with GBC 5.00.08 (https://4js.com/online_documentation/fjs-gbc-manual-html/#gbc-topics/gbc_whatsnew_50008_2.html).  The coding techniques here are still interesting but you no longer need to explciitly add code to do this.

Example showing how to use reflection to take clipboard with array data and paste into a 4gl array variable

Select and Copy a range of cells from Excel

![Screen Shot 2023-05-05 at 4 49 30 PM](https://user-images.githubusercontent.com/13615993/236379853-101df8e9-b359-47f9-bbde-409f91694bfa.png)

Start the test program and select Paste Special (from Excel)

![Screen Shot 2023-05-05 at 4 49 37 PM](https://user-images.githubusercontent.com/13615993/236379852-aac3c4d4-1d65-4490-abc0-c13b3bfcf4d7.png)

Values in Excel should appear in program

![Screen Shot 2023-05-05 at 4 49 42 PM](https://user-images.githubusercontent.com/13615993/236379848-a49f704b-e4a6-43fe-bd15-a8d751431ec0.png)


## Test Notes

To test a succesful paste, make sure you have a range that matches the 4gl array i.e.

|A|B  |C       |
|-|---|--------|
|1|AAA|1/1/2024|
|2|BBB|2/2/2024|

To test failure conditions, have a different shaped range, should get an error message "Number of elements in clipboard does not match number of elements in array"

|A|B  |
|-|---|
|1|AAA|
|2|BBB|


|A|B  |C       |D|
|-|---|--------|-|
|1|AAA|1/1/2024|1|
|2|BBB|2/2/2024|2|

Also note the result of the following test.  The number can go into the string field, whilst the text can't go into the number field.  A future enhancement would be to test and cater for this.

|A  |B|C       |
|---|-|--------|
AAA|1|1/1/2024
BBB|2|2/2/2024

The data will be appended to the end of the array.  If the last row of the array is blank, it will start on that line.  If the last row the array contains data, it will start on the next line.

## Code Usage

Intended Usage is in an INPUT ARRAY similar to ...

```
    INPUT ARRAY arr ...
    
        ON ACTION pastefromexcel ATTRIBUTES(VALIDATE=NO, TEXT="Paste Special (from Excel)")
            CALL paste_from_excel.paste(reflect.Value.valueOf(arr)) RETURNING ok, error_text
            IF NOT ok THEN
                CALL FGL_WINMESSAGE("Error", error_text,"stop")
            END IF
```            

If the paste was succesful, ok will be TRUE
If the paste was unsuccesful, ok will be FALSE, and an error message will be in error_text

Assumes clipboard has data that has lines/rows seperated by ASCII(10) LF, and cells within the line/row seperate by ASCII(9) TAB


