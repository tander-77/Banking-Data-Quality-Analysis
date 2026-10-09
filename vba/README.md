# VBA Data Cleaning & Validation

This directory contains the original VBA source code developed to automate data quality corrections in a banking customer dataset.

## Overview

The `BankingDataCleaning.bas` module contains seven procedures designed to improve data consistency and completeness through rule-based transformations and targeted validation.

The procedures address:

- Standardization of month abbreviations
- Rounding of numerical duration values
- Median imputation for missing ages
- Standardization of missing categorical values
- Validation of selected data quality conditions

Each cleaning procedure tracks the number of values modified and displays a completion message.

## Execution

The macros were developed for Microsoft Excel using Visual Basic for Applications (VBA).

To inspect or run the procedures:

1. Open the appropriate Excel workbook.
2. Open the VBA editor using `Alt + F11`.
3. Import `BankingDataCleaning.bas` through **File → Import File**.
4. Select the worksheet containing the banking dataset.
5. Execute the appropriate cleaning procedures.
6. Run `ValidateCleanedData` to check the targeted data quality conditions.

**Note:** These procedures modify worksheet values directly. Run them on a backup copy of the original dataset.

## Limitations

The validation procedure checks selected missing-value and categorical inconsistencies. It does not independently validate every column or verify the accuracy of individual customer records against external sources.

The original source code is preserved to document the implementation completed during the project.
