# Command Line

It is desirable to have command line program called `mini` to work on he
excel file direcly. The command line program should be put in the file call
`mini.py` with the main caommand being `mini`.

When working with the command line, the excel workbook can be made
invisible so that it will not steal focus.

The main command we want to run on this excel workbook are:

1. fill_formula_wb
2. commercial
3. technical

`fill_formula_wb` is to be run before `commercial` to make sure things are
align before pdf is produced. Messages from `xw.apps.active.alert` messages
are to be output to the terminal and any decision requried can be
interactively display. The PDF produced need not be opened. Only when the
process is finished successfully, the user can open these manually.

`technical` command can be called upon the commercial excel produced in the
process.

Let us make a plan to implement this feature.
