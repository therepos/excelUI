## Calculation
- To count numeric cells, use `=COUNT()`. To count text cells, use `=COUNTA()`.

## Deletion
- To delete thousands or more rows of formulated cells in filtered mode, always clear the unwanted contents first.

## Webscrape
- [Selenium][S1]
- [IE][S2]

[S1]: https://stackoverflow.com/questions/57216623/using-google-chrome-in-selenium-vba-installation-steps
[S2]: https://www.guru99.com/data-scraping-vba.html

## XML ERROR in Style

Use Notepad++ regex replace the following with nothing<sup>[1][S3]</sup>: 
```
ss:StyleID=".*?(?=")"
ss:StyleID="s64"
```
Then delete 
```
<Styles>….</Styles>
```
[S3]: https://stackoverflow.com/questions/19788870/xml-error-in-style-reason-missing-tag

## Left Lookup
```
=INDEX(Target_Range, MATCH(Target_Item, Lookup_Range, 0))
```
Return results to the left of lookup value<sup>[1][S4]</sup>.

[S4]: https://www.excel-easy.com/examples/left-lookup.html#:~:text=The%20VLOOKUP%20function%20only%20looks,value%20in%20a%20given%20range

## Replace Carriage Return
```
SELECT REPLACE(REPLACE(@str, CHAR(13), ''), CHAR(10), '')
```
A newline in SQL or script string can be any of CR, LF or CR+LF. To get them all, you need something like this. See [article][S5].
```
=MID(A2,FIND("""",A2)+1,FIND("""",A2,FIND("""",A2)+1)-FIND("""",A2)-1)
```
Extract comments within two symbols<sup>[1][S6]</sup>. 

[S5]: https://stackoverflow.com/questions/951518/replace-a-newline-in-tsql
[S6]: https://www.extendoffice.com/documents/excel/4861-excel-extract-text-between-single-quotes-double-quotes.html

## Search Whole Word Match
```
'.' + column + '.' LIKE '%[^a-z]pit[^a-z]%'
```
Full text indexes is the answer<sup>[1][S7]</sup>. FYI unless you are using _CS collation, there is no need for a-zA-Z.

[S7]: https://stackoverflow.com/questions/5444300/search-for-whole-word-match-with-sql-server-like-pattern


[Shell](https://www.automateexcel.com/vba/shell/)
