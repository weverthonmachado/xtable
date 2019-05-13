# xtable: Stata module to export table output to Excel


`xtable` exports output from Stata's built-in command `table` to an Excel spreadsheet. It works as drop-in replacement: you can just replace `table` with `xtable` and call it the same way (see [Usage](#usage) for minor restrictions and additional options). `table` is a very powerful and flexible command, but it's not easy to get its nice tables out of Stata for further processing. You have to resort to copy/paste or, at best, [`logout`](http://fmwww.bc.edu/RePEc/bocode/l/logout.html). The `putexcel` command  [introduced](https://blog.stata.com/2013/09/25/export-tables-to-excel/) in Stata 13 made exporting stuff to Excel a lot easier, but it relies on stored results and `table` produces none. 

`xtable` leverages `table`'s `replace` option to create a matrix that reproduces as best as possible what's shown on screen and then exports it using `putexcel`. Because it depends on `putexcel`, `xtable` requires Stata 13 or newer.

*For exporting Stata's output to Word, see [`asdoc`](https://www.statalist.org/forums/forum/general-stata-discussion/general/1435798-asdoc-an-easy-way-of-creating-publication-quality-tables-from-stata-commands).  [`tabout`](http://tabout.net.au/docs/home.php) has a lot of cool features and is very useful, but it can be a little cumbersome for 3-way or higher dimension tables.* 

## Installation 

Install it by typing:
```stata
net install xtable, from ("https://raw.githubusercontent.com/weverthonmachado/xtable/master")
```

## Usage

## Basic syntax

You can use the exact same syntax from `table`, because `xtable` will just pass the arguments to `table` and then export the results. So, instead of running:

```stata
sysuse auto
table foreign rep78, c(mean mpg sd mpg)
```

you can just append an "x" and run:

```stata
sysuse auto
xtable foreign rep78, c(mean mpg sd mpg)
```

Your data will be preserved, so the only difference you will see is a link to the Excel spreadsheet containing the exported table:

![](output.png)

And the spreadsheet will look like this:

![](excel.png)

The only real restriction is that you can not use it with the `by` prefix (i.e. `by varname: xtable`). But bear in mind that you can specify row, column, supercolumn and up to four superrow variables, so you can get up to 7-way tabulations. 

Also, the `concise` option, that suppresses rows with all missing entries, will not affect the exported table. If you use it, you will still get a concise table on Stata's results window, but the Excel spreadsheet will contain all rows. 

## Exporting options

By default, `xtable` will export the tabulation to a file named "xtable.xlsx" in the current working directory, overwriting it if it already exists. You can control the exporting process by using the following options, which will be passed to `putexcel`:

- `filename(string)`: name of the Excel file to be used. Default is "xtable.xlsx". Both .xlsx and .xls extensions are accepted. If you do not specify an extension, .xlsx will be used;
- `sheet(string)`: name of the Excel worksheet to be used. If the sheet exists in the specified file, the default is to modify it. To replace, use `sheet(example, replace)`;
- `replace`: overwrite Excel file. Default is to modify if filename() is specified or overwrite "xtable.xlsx". Note that `table` also has a `replace` option which is not honored by `xtable`, althoguh it actually uses it internally;
- `modify`: modify Excel file.

Example:
```stata
webuse byssin
xtable workplace smokes race [fw=pop], c(mean prob) format(%9.3f) sc filename(myfile) sheet(prevalence) replace
```

Finally, the `noput` option will keep `xtable` from writing to any file. Instead, it will just store the matrix in r(xtable), so you can include it in a `putexcel` call (or use it in another way)

```stata
webuse byssin
xtable workplace smokes race [fw=pop], c(mean prob) format(%9.3f) noput

putexcel A1 = ("A nice and informative title") A3 = (r(xtable), names) using myfile.xlsx, replace
```

This might be particularly useful if you use Stata 14 or newer, which added [formatting options](https://blog.stata.com/2017/01/10/creating-excel-tables-with-putexcel-part-1-introduction-and-formatting/) to `putexcel`.

## Author

**Weverthon Machado**  
PhD Candidate in Sociology  
[Universidade do Estado do Rio de Janeiro](http://www.iesp.uerj.br/)  

[weverthonmachado.github.io](https://weverthonmachado.github.io)
