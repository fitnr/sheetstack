#!/usr/bin/env node

var j = require('j'),
    fs = require('fs'),
    stream = require('stream'),
    program = require('commander'),
    XLSX = require('xlsx'),
    concat = require('concat-stream');

var version = '0.1.0';

program
    .version(version)
    .usage('[OPTIONS] <file>\n  Combine XLS or XLSX sheets into a single CSV')
    .option('-g, --groups <comma-separated list>', 'value of the field to be added at the start of each line (by default, the name of the sheet)')
    .option('-r, --rm-lines <number>', 'number of lines to remove from the start of every sheet (except for the first). default: 1', 1)
    .option('-n, --group-name <name>', 'name of grouping column. default: sheet', 'sheet')
    .option('-F, --field-sep <char>', 'CSV field separator', ',')
    .option('-R, --row-sep <sep>', 'CSV row separator', "\n")
    .option('-o, --output <file>', 'output to specified file')
    .option('-q, --quiet', 'quiet mode');

// get cli args
program.parse(process.argv);

// set options
var
    sheet = '',
    group = '',
    csv = '',
    opts = {
        FS: program.fieldSep,
        RS: program.rowSep
    },
    filename = program.args[0],
    reader = (filename === "-") ? j.read : j.readFile,
    converted = reader(filename),
    XL = (converted[0].utils.make_csv) ? converted[0] : XLSX,
    workbook = converted[1],
    sheets = (program.sheets) ? program.sheets : workbook.SheetNames,
    groups = (program.groups) ? program.groups : sheets,
    re = new RegExp('[^' + program.fieldSep + ']'),
    grouper = function (g) { return function(x) { return g + program.fieldSep + x; }; },
    add_group_name = grouper(program.groupName),
    goodLine = function (x) { return x.search(re) > -1 && x.length > -1; },
    writer = concat(function(data) { console.log(data); });

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