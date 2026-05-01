# Config Sheet

In "Config" sheet from cell B21 to B31, the values should be string without
"\n" character. If user enter any new line character, double spaces, etc,
these should be stripped when doing fill_formula_wb.

In "Config" sheet cell B32, this is a date field. In many user's system,
the date format is set to US format and they cannot change to ISO
yyyy-mm-dd format as admin does not allow changing system time. The desired
format is ISO format. We can convert the user entered date to ISO date
string and write as string or we can set the format to ISO date format.

Can you come up with the plan to implement these features?
