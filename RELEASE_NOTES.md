#### 0.0.1-beta - February 07 2016
* Fix typo in `Workbook.fromXlsxBytes`.
* Swap parameter order for `sheetAt` and `sheetNamed` to allow piping.
* Rename `rows` and `cells`, appending `...WithContent` to clarify that empty
  rows or cells will be skipped. This can be a common mistake when trying to
  iterate through a document and getting an unexpected count.

#### 0.0.1-alpha - February 06 2016
* Initial release
