# Calco

I started this project after spending time once again on code generating `CSV`
files that we open with a spreadsheet software, add formatting, calculations, 
etc. The final spreadsheet document is a tool for the users with the current 
data snapshot.

The code generating the `CSV` file can run several times but we have to manually
re-create the spreadsheet document.

*Calco* tries to separate the data from the spreadsheet presentation and the all
the calculations.

*Calco* implements a DSL (domain specific language) that abstracts the 
calculations and the basic needs for styling and formatting.

It generates a spreadsheet document in different formats depending on the 
selected engine.

The output depends on the engine. The office ([*LibreOffice*](https://www.libreoffice.org)
or [*OpenOffice*](http://www.openoffice.org/)) engine uses an input document as 
template for more sophisticated layouts. The `DefaultEngine` writes simple text 
useful to check the spreadsheet definition.

A spreadsheet contains one or more sheets.

A sheet contain cells (viewed as rows and columns).

A cell can contain:

* literal values (like numbers, dates, times or strings)
* references to other cells
* formulas or functions combining literal values, references, conditionals, 
arithmetics expressions and calls to built-in functions

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'calco'
```

And then execute:

```bash
$ bundle
```

Or install it as:

```bash
$ gem install calco
```

## Usage

### Running the specification:

```bash
$ rspec spec
```
    
Try running specifications with format output...

```bash
$ rspec -f d
```

### Running the examples:

```bash
$ ruby examples/example.rb
```

Replace the `example.rb` with any file of the example directory.

#### Example files (in order of *complexity*)

* [examples/using_date_functions.rb](examples/using_date_functions.rb)
* [examples/example.rb](examples/example.rb)
* [examples/multiplication_tables.rb](examples/multiplication_tables.rb)
* [examples/compute_cells.rb](examples/compute_cells.rb)
* [examples/write_csv.rb](examples/write_csv.rb)
* [examples/write_ods.rb](examples/write_ods.rb)
* [examples/register_function.rb](examples/register_function.rb)

## Quick start

Let's review the [using_date_functions.rb](examples/using_date_functions.rb)
example.

The code must first require some files, like `date` as the example is using 
dates.

```ruby
require 'date'
require 'calco'
```

Now, create the document definition.

```ruby
doc = spreadsheet do

  definitions do

    set some_date: Date.today

    function some_year: year(some_date)
    function age: year(today) - year(some_date)
    
  end
  
  sheet do

    column value_of(:some_date)

    column :some_year
    column :age

  end

end
```

The above code uses the spreadsheet method that returns a document created using 
the given definition (the code block).

The definition contains two parts: the `definitions` and the `sheet`.

The *definitions* part contains all the *variable* declarations (see 
`set some_date`) and the functions (see `some_year` and `age`).

The *sheet* part describes the content of the sheet columns. Here the first 
columns is going to contain the value of the `some_date` variable and the next 
two columns will contain the functions, so that the final spreadsheet is going 
to compute the values using the functions/formulas.

The last part writes the result to the console (using the `$stdout` variable).

```ruby
doc.save($stdout) do |spreadsheet|

  sheet = spreadsheet.current
  
  sheet[:some_date] = Date.new(1934, 10, 3)
  sheet.write_row 3
  
  sheet[:some_date] = Date.new(2004, 6, 19)
  sheet.write_row 5
  
end
```

Saving a document involves calling the `save` method with a block. The block 
receives a spreadsheet object. A spreadsheet object has a current sheet. We can 
assign values to the existing variables (see `sheet[:some_date]` statements) and 
ask the spreadsheet to write a row (passing the index of the row), see 
`sheet.write_row 3`.

The object passed to the block is the same as the one called to save, next code 
is doing the same thing.

```ruby
doc.save($stdout) do

  sheet = doc.current
  
  sheet[:some_date] = Date.new(1934, 10, 3)
  sheet.write_row 3
  
  sheet[:some_date] = Date.new(2004, 6, 19)
  sheet.write_row 5
  
end
```

The final output is:

    A3: 1934-10-03
    B3: YEAR(A3)
    C3: YEAR(TODAY())-YEAR(A3)

    A5: 2004-06-19
    B5: YEAR(A5)
    C5: YEAR(TODAY())-YEAR(A5)

## Definitions

As explained in the previous example, building a spreadsheet object requires to
setup the definitions. What can a definition block contain?

First see the definitions specs: [spec/definitions_spec.rb](spec/definitions_spec.rb)

### Values, variables, references and functions

See specifications

* [spec/variables_spec.rb](spec/variables_spec.rb)
* [spec/absolute_references_spec.rb](spec/absolute_references_spec.rb)
* [spec/conditions_spec.rb](spec/conditions_spec.rb)
* [spec/functions_spec.rb](spec/functions_spec.rb)
* [spec/errors_spec.rb](spec/errors_spec.rb)
* [spec/smart_types_spec.rb](spec/smart_types_spec.rb)
* [spec/builtin_functions_spec.rb](spec/builtin_functions_spec.rb)

### Aggregations...

See specifications

* [spec/range_spec.rb](spec/range_spec.rb)

### Sheet manipulation

See specifications

* [spec/sheet_spec.rb](spec/sheet_spec.rb)
* [spec/sheet_selections_spec.rb](spec/sheet_selections_spec.rb)
* [spec/spreadsheet_spec.rb](spec/spreadsheet_spec.rb)
* [spec/content_change_spec.rb](spec/content_change_spec.rb)

### Styles

See specifications

* [spec/styles_spec.rb](spec/styles_spec.rb)

### Engines

See specifications

* [spec/default_engine_spec.rb](spec/default_engine_spec.rb)
* [spec/csv_engine_spec.rb](spec/csv_engine_spec.rb)
* [spec/calculator_engine_spec.rb](spec/calculator_engine_spec.rb)

## LibreOffice engine

The office engine uses a template file when it writes the output file. It 
searches for each sheet a template sheet in the template file after its name.

If the engine finds a template sheet, it removes the content and inserts the
generated rows. If the template sheet contains the header, the engine does
not remove it (using the `has_titles` directive).

If the engine does not find a template sheet, it appends a new sheet.

Here is an example:

```ruby
engine = Calco::OfficeEngine.new('names.ods')

doc = spreadsheet(engine) do

  definitions do

    set name: ''
    
  end
  
  sheet('Main') do

    has_titles true

    column value_of(:name)

  end
  
end
```

The code creates an office engine and sets a template files named `names.ods`.

The spreadsheet definition defines a sheet named _Main_ and says that the
template sheet contains the header row (see `has_titles true`).

## Tips

### Titles row...

The words *title* or *header* to refer to the first row of a sheet that
represent the columns names or labels...

Because spreadsheets do not use row 0, row 0 (using the Sheet#row method) 
returns headers or empty if no header is set (see 
[spec/header_row_spec.rb](spec/header_row_spec.rb)).

A sheet is marked as having a header row (a first row with titles) by using
the `:title` option or the `has_titles` method.

The use of the header row depends on the engine (the office engine does not
write the column titles).

The effect is that, if the sheet is marked as having a titles row, the data 
output starts at index 2, remember 0 is the header row. Otherwise the data 
starts at index 1. It is important because the formulas contain references like
`A<n>` and `n` must start at 1 if the first row does not contain headers.

### Saving as CSV files

The CSV engine writes files containing formulas (functions) instead of computing
the values. Also, for values like dates and times, it uses function instead of 
plain strings. If you open the CSV file with *LibreOffice* (or *OpenOffice*) it 
recognizes the functions so that you do no loose the benefit of having 
formulas/functions.

As an example the output of [examples/write_csv.rb](examples/write_csv.rb) is

```
Start,End,Duration
"=TIMEVALUE(""12:10:00"")","=TIMEVALUE(""15:30:00"")",=B2-A2
"=TIMEVALUE(""11:00:00"")","=TIMEVALUE(""16:30:00"")",=B3-A3
"=TIMEVALUE(""10:01:00"")","=TIMEVALUE(""12:05:00"")",=B4-A4
"","",=SUM(C1:C4)
```

If you open the file with *LibreOffice* you get something like

```
Start     End       Duration
12:10:00  15:30:00  03:20:00
11:00:00  16:30:00  05:30:00
10:01:00  12:05:00  02:04:00
                    10:54:00
```

The time values are real time values, not strings. The formulas are computed.

## Todo

1. cross-sheet references
2. specs for office engine
   * currencies
   * percentages
   * time
   * date, now
   * absolute $A$5
   * styles (both formulas and values)
3. CSV engine (using the calculator)

## Done

1. specs: func not found, add func, func arity (also check error)
2. build-in function, improve with function registration with types and args,
create `date_functions.rb`, etc.
3. replace "Left" function implementation with build-in mechanism
4. add example showing the function registration
5. split into files...
6. improve ugly code in 'element.rb)'
7. refactor document 'save'
8. 'DefaultEngine' should implement 'save' method
9. add a 'calculator' engine (as an example)
1. simple calculator engine should use internal context for computation
1. specs for simple calculator engine
  * string values
  * with titles and function
  * column title/names => used in output
  * absolute cells reference
  * skip cells
12. spreadsheets have no row '0', 0 always refers to header row
13. specs for errors
    * err: unknown var
    * err: unknown function
    * err: assign the same var twice
    * err: declare the function and var twice
    * err: ArgumentError => `column :price, 'pp'` ('title:' missing)
    * err: column id not found
14. specs for sheet
    * accept expressions like: `function actual_price: 1 + (tax_rate / 100)`
    * last sheet is current
    * setting current sheet => `doc.sheet("a").current`
15. ids for columns
    * err: id for column not found
16. spec for CSV engine
    * `=DOLLAR(A1,2)`
    * `=SUM(A1:A3)`
    * DOLLAR for formula cell
    * DOLLAR and include conditional style `=DOLLAR((A3/100)+STYLE(IF(CURRENT()>3,"Red","Green")))`
    * % type specification
17. specs for Spreadsheet
    * Spreadsheet#row(n) returns value of current sheet
18. removed engine explicit dependency for Element(s), engine is injected by
    sheets using Sheet#compile
19. fixed elements generate methods that did not return values using engine
20. spec for default engine
    * skip columns, sheet#row returns ''
21. API to change engine
22. changed office example to use conditional styles
23. added `empty_row` (in all engines)
24. explained examples in top of files
25. wrote a gem description, reused examples
26. use formula when applying dynamic styles to values
27. mutliple sheets

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
