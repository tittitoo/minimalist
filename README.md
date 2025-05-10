# Code intended to help in tendering / proposal

The project to help in tendering or proposal writing. The code interfaces with excel and uses excel as UI.

## Shortcuts

- `Ctrl+e`: Fill formula
- `Ctrl+j`: Hide rows
- `Ctrl+m`: Unhide rows
- `Ctrl+w`: Insert rows (You can select a number of cells and do the insert as well)
- `Ctrl+q`: Delete rows

## Deployment

### Depandencies

The following software packages are required:

- Ananconda (this will include python, pandas, xlwings, numpy, etc.)
- uv (if bid commandline is to be used)

### xlwings Setup

xlwings add-in is required for us because we are using excel as the UI to talk to python.
xlwings come preinstalled in ananconda. To install add-in, run `xlwings addin install`. Once
the addin is installed, you need to set Interpreter path and PYTHONPATH.

![xlwings add in](https://filedn.com/liTeg81ShEXugARC7cg981h/Resources/xlwings-add-in.png)
The interpreter path is the path to python installation. The PYTHONPATH is where python code
for the project lives. There is a slight different in Mac and Windows configuration.

Example configuration in Mac:

- _Interpreter path_: `/Users/infowizard/Repos/github.com/tittitoo/minimalist/.venv/bin/python`
- _PYTHONPATH_: `/Users/infowizard/Repos/github.com/tittitoo/minimalist`

In Windows:

- _Interpreter path_: `C:\ProgramData\anaconda3\python.exe`
- _PYTHONPATH_: `C:\Users\linzar\Jason Electronics Pte Ltd\Bid Proposal - Documents\@tools`
