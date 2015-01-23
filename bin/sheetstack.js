#!/usr/bin/env node

var j = require('j'),
    fs = require('fs'),
    stream = require('stream'),
    program = require('commander'),
    XLSX = require('xlsx'),
    concat = require('concat-stream');

var version = '0.1.0';

function list(val) { return val.split(','); }

program
    .version(version)
    .usage('[OPTIONS] <file>\n  Combine XLS or XLSX sheets into a single CSV')
    .option('-s, --sheets <items>', 'list of sheets to read (by default, all sheets will be read', list)
    .option('-g, --groups <items>', 'value of the field to be added at the start of each line (by default, the name of the sheet)', list)
    .option('-r, --rm-lines <n>', 'number of lines to remove from the start of every sheet (except for the first). default: 1', 1)
    .option('-n, --group-name <value>', 'name of grouping column. default: sheet', 'sheet')
    .option('-F, --field-sep <sep>', 'CSV field separator', ',')
    .option('-R, --row-sep <sep>', 'CSV row separator', "\n")
    .option('-o, --output <file>', 'output to specified file')
    .option('-q, --quiet', 'quiet mode');

// get cli args
program.parse(process.argv);

// set options
var
    opts = {
        FS: program.fieldSep,
        RS: program.rowSep
    },
    filename = program.args[0],
    writer = (program.output) ? fs.createWriteStream(program.output, {flags: 'w'}) : concat(function(data) {console.log(data);});

// Allow for piping

if (filename === "-") {

    process.stdin.pipe(concat(function(data){
        w = j.read(data);
        sheetstack(w);
    }));

} else {

    converted = j.readFile(filename);
    sheetstack(converted);

}

function sheetstack(converted) {
    var XL = (converted[0].utils.make_csv) ? converted[0] : XLSX,
        workbook = converted[1],
        sheets = (program.sheets) ? program.sheets : workbook.SheetNames,
        groups = (program.groups) ? program.groups : sheets,
        re = new RegExp('[^' + program.fieldSep + ']');

    function grouper(g) {
        return function(x) {
            return g + program.fieldSep + x;
        };
    }

    function goodLine(x) {
        return x.search(re) > -1 && x.length > -1;
    }

    var add_group_name = grouper(program.groupName);

    var sheet, group, csv;

    for (var i = 0, s = sheets.length; i < s; i++) {
        sheet = sheets[i];
        add_group = grouper(groups[i]);

        csv = XL.utils.make_csv(workbook.Sheets[sheet], opts).split('\n');

        if (i === 0) {
            firstline = add_group_name(csv[0]);
            mutated = csv.splice(1).filter(goodLine).map(add_group);
            csv = [firstline].concat(mutated);
        }
        else if (program.rmLines > 0) {
            csv = csv.splice(program.rmLines).filter(goodLine).map(add_group);
        }
        writer.write(csv.join('\n') + '\n');
    }

    writer.end();
}
