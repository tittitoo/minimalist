# xlwings not able to find SharePoint local folder

Currently our solution for Sharepoint folder is to generate commercial or
technical proposal in local folder pointed to by "def get_workbook_directory(wb):
" when wb.fullname returns a url. We however knows the location of the
folder in the python code below and the folder structure.

```python
# Handle case for different users
username = getpass.getuser()
match username:
    case "oliver":
        RFQ = "~/OneDrive - Jason Electronics Pte Ltd/Shared Documents/@rfqs/"
        HO = "~/OneDrive - Jason Electronics Pte Ltd/Shared Documents/@handover/"
        CO = "~/OneDrive - Jason Electronics Pte Ltd/Shared Documents/@costing/"
        DOCS = "~/OneDrive - Jason Electronics Pte Ltd/Shared Documents/@docs/"
        BID_ALIAS = f"alias bid=\"uv run --quiet '{Path(r'~/OneDrive - Jason Electronics Pte Ltd/Shared Documents/@tools/bid.py').expanduser().resolve()}'\""

    case "carol_lim":
        RFQ = "~/Jason Electronics Pte Ltd/Bid Proposal - Documents/@rfqs/"
        DOCS = "~/Jason Electronics Pte Ltd/Bid Proposal - Documents/@docs/"
        BID_ALIAS = f"alias bid=\"uv run --quiet '{Path(r'~/Jason Electronics Pte Ltd/Bid Proposal - @tools/bid.py').expanduser().resolve()}'\""

    case _:
        RFQ = "~/Jason Electronics Pte Ltd/Bid Proposal - Documents/@rfqs/"
        HO = "~/Jason Electronics Pte Ltd/Bid Proposal - Documents/@handover/"
        CO = "~/Jason Electronics Pte Ltd/Bid Proposal - Documents/@costing/"
        DOCS = "~/Jason Electronics Pte Ltd/Bid Proposal - Documents/@docs/"
        BID_ALIAS = f"alias bid=\"uv run --quiet '{Path(r'~/Jason Electronics Pte Ltd/Bid Proposal - Documents/@tools/bid.py').expanduser().resolve()}'\""


```

The source folder is in @rfqs pointed to by RFQ variable. Inside @rfqs
folder, there are subfolders by year. For example,

- 2026
- 2025
- 2024

With this information, when Sharepoint folder returns a url, rather than
defaulting to a local folder, we need to find the folder location based on
`wb.name` in @rfqs and use that folder if it can be found. In the rare
cases that two locations were found we should use the shallower folder and
inform the user. Excel files are preceded with the job code and ends with
version number so that they can considered to be unique in most cases.

Only if the folder location cannot be found, it should default to the
Downloads folder.

Let us come up with a plan to implement this feature.
