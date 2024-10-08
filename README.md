# xml2xlsx

[![image](https://travis-ci.org/pkaczynski/xml2xlsx.svg?branch=master)](https://travis-ci.org/pkaczynski/xml2xlsx)

[![image](https://coveralls.io/repos/github/pkaczynski/xml2xlsx/badge.svg?branch=master)](https://coveralls.io/github/pkaczynski/xml2xlsx?branch=master)

[![image](https://img.shields.io/pypi/v/xml2xlsx.svg)](https://pypi.python.org/pypi/xml2xlsx)

[![image](https://img.shields.io/pypi/pyversions/xml2xlsx.svg)](https://pypi.python.org/pypi/xml2xlsx/)

[![image](https://img.shields.io/pypi/dd/xml2xlsx.svg)](https://pypi.python.org/pypi/xml2xlsx/)

Creating `xlsx` files from `xml` template using
[openpyxl](https://bitbucket.org/openpyxl/openpyxl).

## Target

This project is intended to create `xlsx` files from `xml` api to
`openpyxl`, supposedly generated by other tamplate engines (i.e. django,
jinja).

This is a merely an xml parser translating mostly linearly to worksheet,
rows and finally cells of the Excel workbook.

### Example

An xml file like this one

```xml
<workbook>
    <worksheet title="test">
        <row><cell>This</cell><cell>is</cell><cell>a TEST</cell></row>
        <row><cell>Nice, isn't it?</cell></row>
    </worksheet>
</workbook>
```

can be parsed to create a neat Excel workbook with two rows of data in
one worksheet. Parsing can be done using command line (provided that you
have your system paths set correctly:

```bash
xml2xlsx < input.xml > output.xml
```

or as a library call

```python
from xml2xlsx import xml2xlsx
template = '<sheet title="test"></sheet>'
f = open('test.xlsx', 'wb')
f.write(xml2xlsx(template))
f.close()
```
This is mainly intended (and was developed for this purpose) to parse
files generated by other templating engines, like django template
system. One can generate an excel workbook from template like this:

``` xml
{% for e in list %}
    <row><cell>{{ e.name }}</cell></row>
{% endfor %}
```

## Features

Basic features of the library include creating multiple, named sheets
within one workbook and creating rows of cells in these sheets. However,
there are more possibiliteis to create complex excel based reports.

### Cell type

Each cell can be specified to use one of the types:

-   string (default)
-   number
-   date

Type is defined in `type` cell attribute. The cell value is converted
appropriately to the type specified. If you insert a number in the cell
value and do not specify `type="number"` attribute, you will find Excel
complaining about storing nubers as text.

Since there are more date formats than countries, you have to be aware
of current locale. The simplest way to be i18n compatible is to specify
date format in `date-fmt` attribute and pass compatible (possibily non
localized) date in the cell value, as in the following example

``` xml
...
<row><cell type="date" date-fmt="%Y-%m-%d">2016-10-01</cell></row>
<row><cell type="date" date-fmt="%d.%m.%Y">01.10.2016</cell></row>
...
```

Generated excel file will have two rows with the same date (1st of
October 2016) with date formatted according to Excel defaults (and
current locale).

::: warning
::: title
Warning
:::

Excel tries to be very smart and converts date-like text to date format.
Please use `type="date"` and `date-fmt` attribute always if you pass
dates to cells.
:::

### Columns

Columns can be tackled only in a limited way, i.e. only column widths
can be changed. Column properties are defined in `columns` tag as one or
more child of the `sheet` tag. It is possible to specify a range of
columns using `start` and `end` atrributes. For example:

``` xml
...
<sheet title="test">
    <columns start="A" end="D" width="123"/>
    <row><cell>Test</cell></row>
</sheet>
...
```

### Formulas

`xml2xls` can effectively create cells with formulas in them. The only
limitation (as with `openpyxl`) is using English names of the functions.

For example:

``` xml
...
<row><cell>=SUM(A1:A5)</cell></row>
...
```

### Cell referencing

The parser can store positions of the cell in a dictionary-like
structure. It then can be referenced to create complex formulas. Each
value of the cell is preprocessed using string format with stored
values. This means that these values can be referenced using `{` and `}`
brackets.

#### Current row and column

There are two basic values that can always be used, i.e. `row` and `col`
which return current row number and column name.

``` xml
<workbook>
    <sheet>
        <row><cell>{col}{row}</cell></row>
    </sheet>
</workbook>
...
```

would create a workbook with a text \"A1\" included in the `A1` cell of
the worksheet. Using template languages, you can create more complicated
constructs, like (using django template system):

``` xml
...
{% for e in list %}
<row>
    <cell type="date" date-fmt="%Y-%m-%d">{{ e|date:"Y-m-d" }}</cell>
    <cell>=TEXT(A{row}, "ddd")</cell>
</row>
{% endfor %}
...
```

would create a list of rows with a date in the first column and weekday
names for these dates in the second column (provided `list` context
variable contains a list of dates).

#### Specified cell

It is also possible to store cell possible to store names of specified
cells in a pseudo-variable (as in a dictionary). One has to use `ref-id`
attribute of the `cell` tag and then reuse the value of this attribute
in the remainder of the xml input. This is very useful in formulas. A
simple example would be referencing another cell in a formula like this:

``` xml
...
<row><cell ref-id="mycell">This is just a test</cell></row>
...
<row><cell>={mycell}</cell></row>
...
```

which would create an excel formula referencing a cell with \"this is
just a test\" text, whatever this cell address was.

::: warning
::: title
Warning
:::

Using the same identifier in `ref-id` attribute for two different cells
**overwrites** the cell reference, i.e. the last cell in the xml
template would be referenced.
:::

A more complex example using django template engine to create summaries
can look like this:

``` xml
...
{% for e in list %}
    <row>
        <cell ref-id="{% if forloop.first %}start{% elsif forloop.last %}end{% endif %}">
            {{ e }}
        </cell>
    </row>
{% endfor %}
<row>
    <cell>Summary</cell>
    <cell>=SUM({start}:{end})</cell>
</row>
...
```

#### List of cells

Referencing a single cell can be harsh when dealing with complex
reports. Especially when creating summaries of irregularly
sheet-distributed data. `xml2xlsx` can append a cell to a variable-like
list, as in `ref-id` attribute, to reuse it as a comma concatenated
value. Instead of `ref-id`, one has to use `ref-append` attribute.

This is a simple example to demonstrate the feature:

```xml
...
<sheet>
    <row>
        <cell ref-append="mylist">1</cell>
        <cell ref-append="mylist">2</cell>
    </row>
    <row><cell ref-append="mylist">3</cell></row>
    <row><cell>=SUM({mylist})</cell></row>
</sheet>
```

This will generate an Excel sheet with `A3` cell containing formula to
sum `A1`, `B1` and `A2` cells (`=SUM(A1, B1, A2)`).

#### Referencing limitations

It is perfectly possible to reference a cell in another sheet with both
`ref-id` and `ref-append`. However, there is a limitation to that. Since
`xml2xslx` is a linear parser, you are only allowed to reference already
parsed elements. This means, you have to create sheets in a proper order
(sheets referencing other sheets must be created **after** referenced
cells are parsed).

The following example **will not work**:

``` xml
...
<sheet title="one">
    <row><cell>{mycell}</cell></row>
</sheet>
<sheet title="two">
    <row><cell ref-id="mycell">XYZ</cell></row>
</sheet>
...
```

However, it is possible to make this exmaple work **and** retain the
same worksheet ordering using `index` attribute:

``` xml
...
<sheet title="two">
    <row><cell ref-id="mycell">XYZ</cell></row>
</sheet>
<sheet title="one" index="0">
    <row><cell>{mycell}</cell></row>
</sheet>
...
```

### Cell formatting

The cell format can be specified using various attributes of the cell
tag. Only font formatting can be specifed for now.

#### Font format

A font format is specified in in `font` attribute. It is a semicolon
separated dict like list of font formats as specified in
[font](http://openpyxl.readthedocs.io/en/default/api/openpyxl.styles.fonts.html#openpyxl.styles.fonts.Font)
class of [openpyxl](https://bitbucket.org/openpyxl/openpyxl) library.

An example to create a cell with bold 10px font:

``` 
...
<cell font="bold: True; size: 10px;">Cell formatted</cell>
...
```

### Planned features

Here is the (probably incomplete) wishlist for the project

-   Global font and cell styles
-   Row widths and column heights
-   Horizontal and vertical cell merging
-   XML validation with XSD to quickly raise an error if parsing wrong
    xml

## XML Schema Reference

Parsed xml should be enclosed in a `workbook` tag. Each `workbook` tag
can have multiple `sheet`. The hierarchy continues to `row` and `cell`
tags.

Here is a complete list of available attributes of these tags.

### `workbook`

No attributes for now.

### `sheet`

Attribute

:   `title`

Usage

:   Specifies the worksheet title

Attribute

:   `index`

Usage

:   Specifies the worksheet index. This is relative to already created
    indexes. An index of 0 creates sheet at the beginning of the sheets
    collection.

### `row`

No attributes for now

### `columns`

Attribute

:   `start`

Usage

:   Specifies the starting column for the column range (in a letter
    format).

Attribute

:   `end`

Usage

:   Specifies the ending column for the column range (in a letter
    format).

Default

:   Same as `start` attribute

Attribute

:   `width`

Usage

:   Specifies the width for all columns in the range. It is in px
    format.

### `cell`

Attribute

:   `type`

Usage

:   Specifies the resulting type of the excel cell.

Type

:   One of `unicode`, `date`, `number`

Default

:   `unicode`

Attribute

:   `date-fmt`

Usage

:   Specifies the format of the date parsed as in [strftime and
    strptime](https://docs.python.org/2/library/datetime.html#strftime-and-strptime-behavior)
    functions of `datetime` standard python library.

Remarks

:   Parsed only if `type="date"`.

Attribute

:   `font`

Usage

:   Sepcifies font formatting for a single cell.

Type

:   List of semicolon separated dict-like values in form of
    `key: value; key: value;`

Remarks

:   Key and values are arguments of `Font` clas in `openpyxl`.

### Release History

#### 0.2

-   Added documentation
-   Added cell referencing with inter-sheet possibility
-   Changed `sheet` title attribute from `name` to `title`
-   Added possibility to set index for a sheet