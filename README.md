# excel-vba
Visual Basic for Applications scripts to enhance Excel

Provides VBA modules that implement commonly needed (at least by me)
functionality for manipulating Excel spreadsheets.


## Usage

Import the module or modules you want to use into the Visual Basic editor
within Excel.

**TODO:** include step-by-step instructions here


## Module: haystack-search.bas

Provides functionality for searching a spreadsheet to find matching values


### Main functions

#### FindMatchingRow

Returns the row in worksheet `haystack` with a value in `haystack_column` that matches `needle`.

    @param   haystack         the worksheet object that should be searched
    @param   haystack_column  the column (number or letter) that should be
                              searched
    @param   needle           the value to find in `haystack_column`
    @param   start_row        the row that the search should start with.
                              Defaults to 2 (assumes a headings row).
    @param   blanks_allowed   a boolean indicating whether blank values in
                              `haystack_column` should stop the search for
                              `needle`. Default is False (a blank value will
                              stop the search.
    @return  Long  the number of the row in `haystack` that contains the
                   matching value to `needle`. Returns `0` if not found.


#### FindMatchingRows

Returns a collection of rows in worksheet `haystack` with a value in `haystack_column` that matches `needle`.

    @param   haystack         the worksheet object that should be searched
    @param   haystack_column  the column (number or letter) that should be
                              searched
    @param   needle           the value to find in `haystack_column`
    @param   start_row        the row that the search should start with.
                              Defaults to 2 (assumes a headings row).
    @param   blanks_allowed   a boolean indicating whether blank values in
                              `haystack_column` should stop the search for
                              `needle`. Default is False (a blank value will
                              stop the search.
    @return  Collection  a collection of row numbers in `haystack` that
                         contains the matching value to `needle`. Returns an
                         empty collection if no matches are found.


### Supporting functions

#### ConvertCharColumnToNumber

When provided a column letter it returns the column's number

    @param   column_letter  the letter designation for a column
    @return  Long  the numeric representation for the column


#### DetermineColumnNumber

Returns the column number referenced by `column`

    @param   column  a letter or number representation of a column
    @return  Long  the numeric representation for `Column`




## Module: special-sauce.bas

Provides basic functionality that really should be part of VBA but isn't :(


### Main functions

#### Implode

Returns a string containing all the values in `entries` separated by `delimiter`.

    @param   entries    a Collection of values that will be combined into a
                        single string
    @param   delimiter  a string indicating what character(s) should be used
                        to separate the values in `entries`
