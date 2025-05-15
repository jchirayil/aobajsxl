# Test Cases: JSON/XLSX Processing Library

**Version:** 1.0
**Date:** May 15, 2025

This document details the individual test cases for the JSON/XLSX processing library, as outlined in the Test Plan.

## 1. Functional Test Cases

1.	Test that the correct data is retrieved on reading the Excel file
2.	Test when-\
  a.	File contains-

      i.	Characters - alphabetic, numeric, alphanumeric, symbols, foreign characters, custom characters (EUDC)\
      ii.	(Pivot) Tables, Forms\
      iii.	Illustrations, Controls, Charts, Sparklines, Filters, Links and Comments\
      iv.	Background\
      v.	Page breaks\
      vi.	Formulas\
      vii.	Defined names\
      viii.	Macros

    b.	Different page layout (margins, orientation, page size, etc.) is set
  	
    c.	Rows/Columns are-

      i.	Merged\
      ii.	Hidden\
      iii.	Frozen\
      iv.	Resized (to very small or large scale)

    d.	Sheet is bound to different data sources\
    e.	Cell is bound to – Stocks, Geography, Currencies data types\
    f.	Sorting is applied\
    g.	Filtering is applied\
    h.	Validation is applied\
    i.	Grouping (with subtotal) is applied\
    j.	Different styles/formatting are applied\

3.	Test that specified data is written in the Excel file correctly
4.	Test when-\
  a.	Different JSON data is used

      i.	Plain\
      ii.	Nested\
      iii.	Arrays\
      iv.	Empty/null\
      v.	Duplicate\
      vi.	Lengthy/High precision
  	
    b.	JSON format is incorrect\
    c.	Different datatypes are used

      i.	Test that datatypes are preserved in exported Excel 

5.	Test that warning message appears for entities which are not read/written\
6.	Test that no/handled errors occur when data is read/written from/to \
    a.	new/existing sheets\
    b.	invalid/corrupted Excel files\
    c.	not accessible files
  	
7.	Test that appropriate message appears on reading/writing to password protected/unauthorized files
8.	Test that data is read/written from/to Excel files sync/async operations

## 2. Non-Functional Test Cases

1.	Test that large data is read/written from/to Excel files\
  a.	Test that time and memory consumption is acceptable in this case.
2.	Test that data can be read/written many times through sync/async operations
3.	Test that generated files can be opened in different Excel versions/apps
4.	Test that logs are maintained for operations done with detailed errors if any.
5.	Test that no/handled errors occur on performing above operations


**Note:** This is an initial set of test cases. More detailed test cases will be added as development progresses and more specific functionalities are implemented. The "Priority" indicates the importance of the test case (High = critical, Medium = important, Low = less critical).
