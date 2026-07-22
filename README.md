# XLSX Template in Browser

On-the-fly generation of .xlsx (Excel) files directly within the browser using
templates constructed in Excel.

This library utilizes the native browser DOMParser and XMLSerializer for XML
manipulation, resulting in fewer dependencies and a smaller bundle size
compared to other similar projects.

The basic principle is this: You create a template in Excel. This can be
formatted as you wish, contain formulae etc. In this file, you put placeholders
using a specific syntax (see below). In code, you build a map of placeholders
to values and then load the template, substitute the placeholders for the
relevant values, and generate a new .xlsx file that you can then serve to the
user.

## Placeholders

Placeholders are inserted in cells in a spreadsheet. It does not matter how
those cells are formatted, so e.g. it is OK to insert a placeholder (which is
text content) into a cell formatted as a number or currecy or date, if you
expect the placeholder to resolve to a number or currency or date.

### Scalars

Simple placholders take the format `${name}`. Here, `name` is the name of a
key in the placeholders map. The value of this placholder here should be a
scalar, i.e. not an array or object. The placeholder may appear on its own in a
cell, or as part of a text string. For example:

    | Extracted on: | ${extractDate} |

might result in (depending on date formatting in the second cell):

    | Extracted on: | Jun-01-2013 |

Here, `extractDate` may be a date and the second cell may be formatted as a
number.

Inside scalars there possibility to use array indexers. 
For example: 

Given data

    var template = { extractDates: ["Jun-01-2113", "Jun-01-2013" ]}

which will be applied to following template

    | Extracted on: | ${extractDates[0]} |

will results in the 

    | Extracted on: | Jun-01-2113 |

### Columns

You can use arrays as placeholder values to indicate that the placeholder cell
is to be replicated across columns. In this case, the placeholder cannot appear
inside a text string - it must be the only thing in its cell. For example,
if the placehodler value `dates` is an array of dates:

    | ${dates} |

might result in:

    | Jun-01-2013 | Jun-02-2013 | Jun-03-2013 |

### Tables

Finally, you can build tables made up of multiple rows. In this case, each
placeholder should be prefixed by `table:` and contain both the name of the
placeholder variable (a list of objects) and a key (in each object in the list).
For example:

    | Name                 | Age                 |
    | ${table:people.name} | ${table:people.age} |

If the replacement value under `people` is an array of objects, and each of
those objects have keys `name` and `age`, you may end up with something like:

    | Name        | Age |
    | John Smith  | 20  |
    | Bob Johnson | 22  |

If a particular value is an array, then it will be repeated across columns as
above.

## Generating reports

To use the library, you need the code like this:

```js
    import { generateXlsx, downloadXlsx } from 'xlsx-template-browser'

    // Set your values to fill the template
    const values = {
      extractDate: new Date(),
      dates: [ new Date("2013-06-01"), new Date("2013-06-02"), new Date("2013-06-03") ],
      people: [
        {name: "John Smith", age: 20},
        {name: "Bob Johnson", age: 22}
      ]
    }

    // You can provide URL to the Excel template or the array buffer
    const template = '[URL TO YOUR TEMPLATE]'

    // You can generate the report and keep it in memory
    const generateReport = async () => {
      return await generateXlsx(template, values)
    }

    // Alternatively, you can download the generated report immediately
    const downloadReport = async () => {
      return await downloadXlsx(template, values, 'My fabulous report.xlsx')
    }
```

## Comments and validations

You can add custom comments and data validations to a spreadsheet. To do so,
add the `$comments` and `$validations` arrays to the data object you pass to
the template. Each array should contain objects with the following structure:

```js
import { downloadXlsx } from 'xlsx-template-browser'

// Values used to fill the template
const values = {
  extractDate: new Date(),
  $comments: [
    {
      sheet: 'Sheet1',
      ref: 'A1',
      text: 'My fabulous comment in cell A1'
    },
    {
      sheet: 'Sheet1',
      ref: 'A2',
      text: 'Another comment nearby'
    }
  ],
  $validations: [
    {
      sheet: 'Sheet1',
      type: 'custom',
      sqref: '$A:$A',
      formula: '=ISNUMBER(A1)'
    },
    {
      sheet: 'Sheet1',
      type: 'list',
      sqref: '$B:$B',
      formula: '=INDIRECT("NamedTable[FieldOfNamedTable]")'
    }
  ]
}

// You can provide either a URL to the Excel template or an ArrayBuffer.
const template = '[URL TO YOUR TEMPLATE]'

// Download the generated report.
const downloadReport = async () => {
  return await downloadXlsx(template, values, 'My fabulous report.xlsx')
}
```

> **Note:** Comments can only be added to worksheets that already contain
  at least one comment in the template. This is an Excel file format limitation.

## Caveats

* The spreadsheet must be saved in `.xlsx` format. `.xls`, `.xlsb` or `.xlsm`
  won't work.
* Column (array) and table (array-of-objects) insertions cause rows and cells to
  be inserted or removed. When this happens, only a limited number of
  adjustments are made:
    * Merged cells and named cells/ranges to the right of cells where insertions
      or deletions are made are moved right or left, appropriately. This may
      not work well if cells are merged across rows, unless all rows have the
      same number of insertions.
    * Merged cells, named tables or named cells/ranges below rows where further
      rows are inserted are moved down.
    * Formulae are not adjusted.
    * Tables, merged cells, conditional formatting, and data validation may
      break if arrays are inserted from the left, as existing cells will be shifted
* As a corollary to this, it is not always easy to build formulae that refer
  to cells in a table (e.g. summing all rows) where the exact number of rows
  or columns is not known in advance. There are two strategies for dealing
  with this:
    * Put the table as the last (or only) thing on a particular sheet, and
      use a formula that includes a large number of rows or columns in the
      hope that the actual table will be smaller than this number.
    * Use named tables. When a placeholder in a named table causes columns or
      rows to be added, the table definition (i.e. the cells included in the
      table) will be updated accordingly. You can then use things like
      `TableName[ColumnName]` in your formula to refer to all values in a given
      column in the table as a logical range.
* Placeholders only work in simple cells and tables.

## Changelog History

### Version 0.3.1

* `$comments` and `$validations` are added

### Version 0.3.1

* Bugfix in types
* Proper extensions of rows and columns for tables, pivot tables, data validations,
  and conitional formatting un case of multiple extensions
* Proper handling of undefined ${table:} values

### Version 0.3.0

* Performance is optimized
* The libary supports merged cells, tables, pivot tables, changes in comments,
  URL rendering

## Alternatives

There are few altertantives for this library:

* [xlsx-template](https://www.npmjs.com/package/xlsx-template) - this library
  was the source of inspiration. The only issue that it works in Node only.
  Also it allows you to insert images and has more dependencies. In contrast to
  **xlsx-template**, **xlsx-template-browser** also support property accessors
  with bracket notation or mixed notation.
* [node-xlsx](https://www.npmjs.com/package/node-xlsx) - this library allows to
  manipulate the data in worksheet, but the style of cells will be lost. The
  fork [js-xlsx](https://github.com/protobi/js-xlsx/) might address the issue,
  but it hasn't been updated for 4 years and could be considered overkill.
  Also **node-xlsx** is under Apache2 license.
