# Common Excel Addin

Common Excel Addin aims to provide commonly used formatting features in a single tab with customisable settings. 

## Quickstart
- Download [Common Excel Addin v0.1.0](https://github.com/therepos/msexcel/blob/main/temp/bas/commonaddin-r010.xlam)  
- Install the downloaded .xlam file. See [how to install add-ins]. 

## Features

![Features](/img/img-commonaddin-tabmain.png)

### Custom Functions

|Functions|Description|
|:--|:--|
|[**XCOMPARE**](https://github.com/therepos/msexcel/blob/main/temp/bas/XCOMPARE.bas)|Returns the word difference between two ranges.|
|[**XEXTRACTAFTER**](https://github.com/therepos/msexcel/blob/main/temp/bas/XEXTRACTAFTER.bas)|Returns the part of a selected range after a specified word.|
|[**XEXTRACTBEFORE**](https://github.com/therepos/msexcel/blob/main/temp/bas/XEXTRACTBEFORE.bas)|Returns the part of a selected range before a specified word.|
|[**XFIND**](https://github.com/therepos/msexcel/blob/main/temp/bas/XFIND.bas)|Returns the word search results on a selected range based on a specified word list.|
|[**XHASNUMBER**](https://github.com/therepos/msexcel/blob/main/temp/bas/XHASNUMBER.bas)|Returns True if there is a number in the selected range.|
|[**XLOOKUP**](https://github.com/therepos/msexcel/blob/main/temp/bas/XLOOKUP.bas)|Returns the matched lookup value from a search list.|
|[**XREPLACEWORDS**](https://github.com/therepos/msexcel/blob/main/temp/bas/XREPLACEWORDS.bas)|Replaces words in a selected range based on specified replacement word list.|
|[**XSPELLNUMBER**](https://github.com/therepos/msexcel/blob/main/temp/bas/XSPELLNUMBER.bas)|Spells monetary values in dollar and cents.|
|[**XSUBSTITUTEPREFIX**](https://github.com/therepos/msexcel/blob/main/temp/bas/XSUBSTITUTEPREFIX.bas)|Replaces the prefix of a selected range based on a specified replacement.|
|[**XSUBSTITUTESUFFIX**](https://github.com/therepos/msexcel/blob/main/temp/bas/XSUBSTITUTESUFFIX.bas)|Replaces the suffix of a selected range based on a specified replacement.|
|[**XTRANSLATE**](https://github.com/therepos/msexcel/blob/main/temp/bas/XTRANSLATE.bas)|Returns the Google Translation result on a selected range.|
|[**XCELLFORMULA**](https://github.com/therepos/msexcel/blob/main/temp/bas/XCELLFORMULA.bas)|Returns formula of the selected cell.|
|[**XCLEANTEXT**](https://github.com/therepos/msexcel/blob/main/temp/bas/XCLEANTEXT.bas)|Removes excess non-alphanumeric characters|
|[**XGETPAGENUMBER**](https://github.com/therepos/msexcel/blob/main/temp/bas/XGETPAGENUMBER.bas)|Returns page number.|
|[**XIFDATE**](https://github.com/therepos/msexcel/blob/main/temp/bas/XIFDATE.bas)|Returns True if it is date format.|
|[**XREMOVEBETWEEN**](https://github.com/therepos/msexcel/blob/main/temp/bas/XREMOVEBETWEEN.bas)|Removes text between two specified delimiters.|
|[**XREMOVESYMBOLS**](https://github.com/therepos/msexcel/blob/main/temp/bas/XREMOVESYMBOLS.bas)|Removes leading and trailing symbols from text.|
|[**XSHEETNAME**](https://github.com/therepos/msexcel/blob/main/temp/bas/XSHEETNAME.bas)|Returns worksheet name.|
|[**XSUBSTITUTEMULTIPLE**](https://github.com/therepos/msexcel/blob/main/temp/bas/XSUBSTITUTEMULTIPLE.bas)|Substitutes multiple words.|

### Examples
![ExampleA](/img/img-commonaddin-r010.gif)


[how to install add-ins]: https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460#:~:text=COM%20add%2Din-,Click%20the%20File%20tab%2C%20click%20Options%2C%20and%20then%20click%20the,install%2C%20and%20then%20click%20OK.
