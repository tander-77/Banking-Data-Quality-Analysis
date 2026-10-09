# Data Quality Assessment & Cleaning Methodology

## Project Overview

This project focused on improving the quality of a banking customer dataset intended to support analysis of customer interest in term deposits.

The original dataset contained missing information, inconsistent categorical values, and numerical formatting issues. To address these problems, I developed Microsoft Excel VBA procedures that applied targeted cleaning rules and tracked the number of corrections performed.

The cleaning process documented **1,271 corrections** across six data fields, followed by a separate validation procedure to check for remaining targeted data quality issues.

## Data Quality Assessment

The assessment focused on three primary dimensions of data quality:

- **Completeness:** Identifying missing values in customer attributes.
- **Consistency:** Standardizing categorical labels and representations of unavailable information.
- **Validity:** Correcting numerical values that did not conform to the expected field format.

## Cleaning Procedures & Results

| Data Field | Identified Issue | Cleaning Method | Corrections |
|---|---|---|---:|
| Month | Inconsistent full month names | Standardized to three-letter abbreviations | 223 |
| Duration | Decimal values in a whole-second field | Rounded to whole numbers | 716 |
| Age | Missing numeric values | Replaced with median age of 39 | 28 |
| Marital Status | Blank categorical values | Standardized to `unknown` | 18 |
| Contact | Inconsistent missing-value representation | Replaced `N/A` with `unknown` | 17 |
| Previous Outcome | Inconsistent missing-value representations | Standardized to `unknown` | 269 |
| **Total** | | | **1,271** |

The procedures were implemented using VBA loops, conditional logic, string operations, and numerical calculations.

Each cleaning procedure reported the number of records modified, providing a documented record of the transformations.

## Data Validation

After the cleaning procedures were completed, a separate VBA validation macro was executed to evaluate selected data quality conditions.

The validation checked for targeted missing and inconsistent values involving age, marital status, contact information, and previous campaign outcomes.

The procedure returned:

**"Validation complete! No targeted data quality issues remain."**

This result confirmed that the conditions evaluated by the validation procedure were no longer detected in the cleaned dataset.

The validation did not independently verify every data field or establish that all customer information was factually accurate.

## Handling Uncertain Values

An important consideration during the cleaning process was distinguishing confirmed inconsistencies from potentially legitimate observations.

Several unusual values were preserved rather than automatically modified:

- **Negative account balances:** Retained because negative balances may represent legitimate financial conditions.
- **`pdays` values of -1:** Preserved because their meaning could not be confirmed without supporting documentation.
- **Unusually high values in `previous`:** Flagged as requiring additional verification rather than being automatically changed.

Preserving uncertain values helped reduce the risk of introducing unsupported assumptions into the dataset.

## Business Relevance

Data quality directly influences the reliability of business reporting and predictive analysis.

Inconsistent categories can distort group comparisons, while missing numerical information can reduce usable observations or influence analytical results.

For a banking organization evaluating customer interest in term deposits, these issues could affect customer segmentation, marketing analysis, and predictive modeling.

The VBA cleaning procedures demonstrate how rule-based automation can improve dataset consistency and reduce the manual effort required to prepare data for analysis.

## Limitations & Future Improvements

Although the cleaning procedures addressed identified issues, several limitations remain:

- The validation procedure covered selected data quality conditions rather than every field.
- Median imputation introduced estimated values for missing ages.
- Certain unusual observations require external verification.
- The original macros operate on the active Excel worksheet.
- The procedures were developed for the structure and formatting of the project dataset.

Future improvements could include expanded validation coverage, automated data quality reporting, more robust error handling, and implementation of the cleaning workflow using Python or SQL.

## Conclusion

This project demonstrates the use of Excel VBA automation to identify and correct data quality issues through structured, repeatable procedures.

By documenting 1,271 corrections and performing targeted validation, the project highlights practical skills in data preparation, automation, validation, and analytical problem-solving.

The resulting dataset is more consistent and suitable for subsequent analysis, while the documented limitations identify areas requiring additional verification or future development.
