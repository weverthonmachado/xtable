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

The only real restriction is that you can not use it with the `by` prefix (i.e. `by varname: xtable`). But bear in mind that with `table` and `xtable` you can specify row, column, supercolumn and up to four superrow variables, so you can get up to 7-way tabulations. 

Also, the `concise` option, that suppress rows with all missing entries, will not affect the exported table. If you use it, you will still get a concise table on Stata's results window, but the Excel spreadsheet will contain all rows. 


## Work in progress

There are a few cosmetic that I will improve soon, such as handling supercolumns labels. 

Other things in the roadmap include an option to not call `putexcel` directly and just keep the created matrix on memory instead, so that the user can fully explore `putexcel`'s features, use it in a for loop, etc. E.g.:

```stata
****** Will not work with current version of xtable ******
sysuse auto
putexcel set myfile.xlsx
foreach v in rep78 headroom foreign {
    xtable `v', cont(mean mpg)
    putexcel A1=matrix(xtable, names), sheet(`v') modify
}
```



## Author

**Weverthon Machado**  
PhD Candidate in Sociology  
[Universidade do Estado do Rio de Janeiro](http://www.iesp.uerj.br/)  

[weverthonmachado.github.io](https://weverthonmachado.github.io)
