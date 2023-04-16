# Platform-Control-Python-scripts
A couple of Python scripts I made whilst working at Platform Control. These scripts automated/streamlined a couple of tasks.

## Python script | combine_amazon_reports
### Introduction
This script translates and combines multiple Amazon payment reports with different languages/layouts. The script looks in an Excel file for the correct translations. You should download the payment reports at sellercentral. The reports can be found in 'Reports\Payments\Date Range Reports'.

**Libraries used**
- regex
- openpyxl
- pandas
- os

### How it works
The Python script reads the Excel master file containing the translated payment and column header values. The Excel files are read by Python using the library 'openpyxl'. The script looks for the English names in the file and combines them with the translated values. Columns need to start at 1 because Excel has no row or colum '0'.

1. Importing libraries and declaring variables
2. Check how many English type of payments there are in the Excel file
3. Defining function `translate_payment_reports(countrycode)`. This function takes the country code as a string input e.g. 'ES'.
4. The function `translate_payment_reports(countrycode)` returns the following data `translation_list, translated_transaction_type_list, dictionary_english_translated`
5. These return values are then used in the main program loop which translates the different Excel reports.
6. In the main program the input Excel reports are read an appended to a Pandas dataframe.
7. If the main program is done the Pandas dataframe will be exported as an Excel file.
