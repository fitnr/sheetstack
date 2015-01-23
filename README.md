# sheetstack

Sheetstack is a command line utility that merges multiple XLS/X sheets into a single CSV.

It's a simple extension of [J](https://www.npmjs.com/package/j) useful for processing files with the exact same layout split into several worksheets. 

Sheetstack adds a grouping column to the resulting CSV. By default this is the name of the sheet, but it could be anything.

## Install

````
npm install sheetstack
````

## Usage

Let's say we have an xls file with two sheets, "dogs" and "cats":

````csv
name,best friend
Pluto,Mickey
Santa's Little Helper,Bart
Scooby Doo,Shaggy
````

````csv
name,best friend
Cat in the Hat,the fish
Garfield,Jon
Hello Kitty,you
````

The simplest use will combine all the files and output the result to stdout.

````
$ sheetstack file.xls

sheet,name,best friend
dogs,Pluto,Mickey
dogs,Santa's Little Helper,Bart
dogs,Scooby Doo,Shaggy
cats,Cat in the Hat,the fish
cats,Garfield,Jon
cats,Hello Kitty,you
````

### Sheets
The `--sheets` option controls which sheets are included, and in what order.

````
$ sheetstack --sheets dog,cat file.xls

sheet,name,best friend
dogs,Pluto,Mickey
dogs,Santa's Little Helper,Bart
dogs,Scooby Doo,Shaggy
````

### Groups
The `--groups` option specifies custom values for the grouping column, `--group-name` sets the value for the top of the column.

````
$ sheetstack --groups canine,feline --group-name species

species,name,best friend
canine,Pluto,Mickey
canine,Santa's Little Helper,Bart
canine,Scooby Doo,Shaggy
feline,Cat in the Hat,the fish
feline,Garfield,Jon
feline,Hello Kitty,you
````

### Removing leading lines

By default, sheetstack removes the first line from all sheets except for the first one. The can be changed with the `--rm-lines` setting. If `--rm-lines` is set to 0, no lines will be removed. Higher values will remove more lines, but no lines will be removed from the first sheet.

### Output format

You can also specify row-separator and field-separator options, which are passed through to J:

````
$ sheetstack --row-sep '\r\n' --field-sep ;
````

