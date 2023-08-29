
![Zrzut ekranu 2023-08-29 085635](https://github.com/yoman38/Added-Excel-Functions/assets/124726056/676226bf-957f-4384-acb0-fc6665ed4772)

Info can be found in excel “insert function”

**Excel Functions:**

1. **EQUIV_X (VBA)** - Searches for a value in a range of cells and returns its position.

2. **RECHERCHE_X (VBA)** - Searches for a value in a range of cells and returns a value at the same position in another range of cells.

3. **MAX_IF_COLOR (VBA)** - Returns the MAX value of cells with a specified background color (works with cells colored using MFC).

4. **MIN_IF_COLOR (VBA)** - Returns the MIN value of cells with a specified background color (works with cells colored using MFC).

5. **AVERAGE_IF_COLOR (VBA)** - Calculates the average of cells with a specified background color (ignores non-numeric values, works with cells colored using MFC).

6. **COUNT_COLORED (VBA)** - Counts the number of cells with a colored background (cells with white background or no background are not counted, works with cells colored using MFC).

7. **COUNT_IF_COLOR (VBA)** - Counts the number of cells with a specified background color (works with cells colored using MFC).

8. **COLOR_NO (VBA)** - Returns the color number of a cell (works with cells colored using MFC).

9. **COLOR_NO_HEX (VBA)** - Returns the color number from a hexadecimal color value (e.g., #00ff00), returns -1 if there's an error.

10. **COLOR_NO_HEX_CELL (VBA)** - Returns the hexadecimal color value from the background color of a cell.

11. **COLOR_NO_RGB_CELL (VBA)** - Returns the RGB values from the background color of a cell and allows setting a separator (optional).

12. **SUM_IF_COLORED (VBA)** - Calculates the sum of cells with a colored background, ignoring non-numeric values (cells with white background or no background are not included, works with cells colored using MFC).

13. **SUM_IF_COLOR (VBA)** - Calculates the sum of cells with a specified background color, ignoring non-numeric values (works with cells colored using MFC).

14. **COUNT_IF_REGEX (VBA)** - Counts the number of cells that match a regular expression.

15. **REGEX_EXTRACT (VBA)** - Extracts one or more parts of a string using regular expressions and allows setting a separator (optional).

16. **REGEX_MATCH (VBA)** - Checks if a string matches a regular expression.

17. **REGEX_REPLACE (VBA)** - Replaces one or more parts of a string using regular expressions.

18. **SUM_IF_REGEX (VBA)** - Calculates the sum of cells that match a regular expression.

19. **DUPLICATES (VBA)** - Returns TRUE if the cell range contains duplicates (ignoring empty cells) or FALSE if it doesn't have any duplicates.

20. **DUPLICATE_ADDRESSES (VBA)** - Returns the addresses of duplicates in a cell range (ignoring empty cells) and allows setting the address separator (optional) and the value to return if there are no duplicates (optional).

21. **DUPLICATE_LIST (VBA)** - Returns the list of duplicates in a cell range (ignoring empty cells) and allows setting the list separator (optional) and the value to return if there are no duplicates (optional).

22. **UNIQUE_LIST (VBA)** - Returns the list of values in a cell range, excluding duplicates (ignoring empty cells), and allows setting the list separator (optional) and sorting order (optional).

23. **JOIN (VBA)** - Joins the values in a cell range and allows setting the separator (optional).

24. **JOIN_SORTED (VBA)** - Joins the values in a cell range, separated by a separator (optional), and sorts them in ascending (default) or descending order.

25. **JOIN_NON_EMPTY (VBA)** - Joins the values in a cell range (ignoring empty cells) and allows setting the separator (optional).

26. **JOIN_NON_EMPTY_SORTED (VBA)** - Joins the non-empty values in a cell range, separated by a separator (optional), and sorts them in ascending (default) or descending order.

27. **JOIN_UNIQUE (VBA)** - Joins the unique values (without duplicates or empty cells) in a cell range and allows setting the separator (optional).

28. **JOIN_UNIQUE_SORTED (VBA)** - Joins the unique values (without duplicates or empty cells) in a cell range, separated by a separator (optional), and sorts them in ascending (default) or descending order.

29. **EXTRACT_WORD (VBA)** - Returns the nth word from a string (allows defining up to 3 separators).

30. **TEXT_COUNT (VBA)** - Counts the number of times a value appears in a text.

31. **CONVERT_TO_HOURS (VBA)** - Converts time values to the number of hours (e.g., "03:45:00" and "3h45" will be converted to 3.75).

32. **DATE_DIFF (VBA)** - Returns the difference in days between two dates.

33. **DATE_EU (VBA)** - Returns TRUE if Excel uses the EU date format (dd/mm) or FALSE if it uses the US date format (mm/dd).

34. **DATE_REVERSE_MD (VBA)** - Returns a date by reversing the day and month of a date (returns the same date if it's not possible).

35. **DAYS_IN_MONTH (VBA)** - Returns the number of days in a month based on a date.

36. **ISO_WEEK_NUMBER (VBA)** - Returns the ISO week number based

 on a date (from 1900 to 2200).

37. **ASCENSION_DATE (VBA)** - Returns the date of Ascension based on a year (or the year of a date) from 1900 to 2200.

38. **PENTECOST_MONDAY_DATE (VBA)** - Returns the date of Pentecost Monday based on a year (or the year of a date) from 1900 to 2200.

39. **EASTER_DATE (VBA)** - Returns the date of Easter based on a year (or the year of a date) from 1900 to 2200.

40. **IS_EASTER (VBA)** - Returns TRUE if the date (from 1900 to 2200) corresponds to Easter or FALSE if it doesn't.

41. **IS_EMAIL (VBA)** - Returns TRUE if the string is a valid email address or FALSE if it's not.

42. **IS_URL (VBA)** - Returns TRUE if the string is a valid URL or FALSE if it's not.

43. **HTML_TABLE (VBA)** - Joins the values in a cell range as a simple HTML table.

44. **ADVANCED_HTML_TABLE (VBA)** - Joins a cell range as an HTML table, preserving major formatting.

**VBA Functions (Without UserForm):**

1. **arrayAdd** - Increases the size of an array by 1 and adds a value to the last position.

2. **arrayCount** - Returns the number of times the searched value is present in the array.

3. **arrayDebug** - Displays the content (or part of the content) of an array in a MsgBox.

4. **arrayDuplicates** - Returns True if the array contains duplicates or False if it doesn't.

5. **arrayDuplicatesDelete** - Deletes all duplicates from an array.

6. **arrayDuplicatesList** - Counts the number of times each value is present in the array and adds a 2nd dimension to the array to save these values (1 = unique, 2 = double value, etc).

7. **arrayEmpty** - Returns True if the array is empty or False if it's not.

8. **arrayPos** - Returns the (first) position of the searched value in the array or returns -1 if the value is not found.

9. **arrayRandomize** - Randomly shuffles the values of an array.

10. **arraySortAsc** - Sorts the values of an array in ascending order.

11. **arraySortDesc** - Sorts the values of an array in descending order.

12. **arrayMax** - Returns the largest numeric value present in the array.

13. **arrayMin** - Returns the smallest numeric value present in the array.

14. **arrayNumDelete** - Deletes a value from an array (based on its position in the array).

15. **arrayValuesDelete** - Deletes all values in an array matching the searched value.

16. **inArray** - Returns True if the value is found in the array or False if it's not.

17. **array2dDebug** - Displays the content (or part of the content) of a 2-dimensional array in a MsgBox.

18. **regexExtract** - Extracts one or more parts of a string using regular expressions.

19. **regexMatch** - Tests if a string matches a regular expression.

20. **regexReplace** - Replaces one or more parts of a string using regular expressions.

21. **colLetter** - Returns a column number as letter(s) from a column number as digit(s).

22. **colNum** - Returns a column number as digit(s) from a column number as letter(s).

23. **lastCol** - Returns the column number of the last value in a row, with the option to specify the sheet (optional).

24. **lastRow** - Returns the row number of the last value in a column (in numeric or letter format), with the option to specify the sheet (optional).

25. **lastUsedCol** - Returns the number of the last used column in the sheet.

26. **lastUsedRow** - Returns the number of the last used row in the sheet.

27. **isInt** - Returns True if the value is an integer number or False if it's not.

28. **intRand** - Returns a random integer between two values.

29. **cellsSearch** - Searches for a value in a cell range and returns the list of addresses of all cells containing the searched value (as an array).

30. **isoWeekNum** - Returns the ISO week number based on a date (from 1900 to 2200).

31. **nbDaysMonth** - Returns the number of days in a month based on a date.

32. **euDate** - Returns True if Excel uses the EU date format (dd/mm) or False if it uses the US date format (mm/dd).

33. **easterDate** - Returns the date of Easter based on a year (or the year of a date) from 1900 to 2200.

34

. **isEmail** - Returns True if the string is a valid email address or False if it's not.

35. **isUrl** - Returns True if the string is a valid URL or False if it's not.

36. **mail** - Sends an email (without using Outlook) through an email solution compatible with all email addresses.

37. **htmlCodePage** - Retrieves the HTML code of a web page, returns -1 in case of an error.

38. **internet** - Returns True if connected to the internet or False if not connected (or blocked by security software).

39. **linkOpen** - Opens a web page link and returns True (returns False if the link can't be opened).

40. **colorToHexa** - Returns the hexadecimal color value from a Color value.

Why C# dev in WIP ? C# would allow the user to get native function, i.e hovering '?' would display the description instead of having to click on fx currently.

42. **hexaToColor** - Returns the Color value of a hexadecimal color (e.g., #00ff00), returns -1 in case of an error.

**VBA Functions (With UserForm):**

1. **colorBox** - Opens a dialog box allowing the user to choose a color from a palette of 160 colors.

2. **datePicker** - Opens a dialog box in the form of a calendar allowing the user to choose a date (from 1900 to 2100).

