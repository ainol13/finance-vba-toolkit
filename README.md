# Finance VBA Toolkit

A collection of Excel VBA tools designed to streamline everyday finance workflows, especially repetitive data preparation and reference lookups.

## Modules

- **[SplitDatabySheet.bas](https://github.com/ainol13/finance-vba-toolkit/blob/main/modules/SplitDatabySheet.bas)** — Splits the active worksheet’s data into separate worksheets based on a column you choose. It carries over the header and formatting, and puts rows with a blank grouping value—but other data present—into a review log so they’re not overlooked.

- **[XLookupVBA.bas](https://github.com/ainol13/finance-vba-toolkit/blob/main/modules/XLookupVBA.bas)** — Looks up keys in a configured master workbook and fills the matching values into the active worksheet. It performs a similar job to Excel’s `XLOOKUP` formula, but packages the process into a reusable, profile-based macro: select a lookup, run it across the data, and review any keys that weren’t found. Think of it as a quality-of-life upgrade for recurring lookups.

## Supporting form

- **[frmMasterSelector.frm](https://github.com/ainol13/finance-vba-toolkit/blob/main/forms/frmMasterSelector.frm)** — Provides the profile selector used by `XLookupVBA.bas`. It lets you choose a lookup profile, run the lookup, or open the associated master workbook.

## Setup

Import the `.bas` modules and the `.frm` form into an Excel VBA project. Before running the lookup tool, update the profiles in `GetProfile` in `XLookupVBA.bas` with the appropriate master workbook paths, worksheet names, key columns, and output columns. The paths currently in the code are examples; replace them with approved locations for your environment.
