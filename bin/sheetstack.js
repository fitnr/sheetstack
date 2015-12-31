#!/usr/bin/env node

var j = require('j'),
    fs = require('fs'),
    stream = require('stream'),
    program = require('commander'),
    concat = require('concat-stream'),
    sheetstack = require('../index');

var version = '0.1.3';

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
var filename = program.args[0],
    writer = (program.output) ? fs.createWriteStream(program.output, {flags: 'w'}) : concat(function(data) {
        console.log(data);
    }),
    callback = function(err, csv) {
        if (err) throw err;

        writer.write(csv);
        writer.end();
    };

// Allow for piping
if (filename === "-") {
    process.stdin.pipe(concat(function(data){
        w = j.read(data);
        sheetstack(w, program, callback);
    }));

} else {
    converted = j.readFile(filename);
    sheetstack(converted, program, callback);
}
