var XLSX = require('xlsx');

function grouper(g, sep) {
    return function(x) {
        return g + sep + x;
    };
}

// ignore lines that are empty or only empty fields
function goodLine(re) {
    return function(x) {
        return x.length > -1 && x.search(re) > -1;
    };
}

function sheetstack(data, config, callback) {
    /*
    :data contents of xls/x file 
    :config object sheets, groups, fieldSep, rowSep, groupName, rmLines
    :callback function. Should take err, string arguments
    */
    try {
        var XL = (data[0].utils.sheet_to_csv) ? data[0] : XLSX,
            workbook = data[1],
            sheets = (config.sheets) ? config.sheets : workbook.SheetNames,
            groups = (config.groups) ? config.groups : sheets,
            fieldSep = config.fieldSep || ',',
            rmLines = config.rmLines || 0,
            rowSep = config.rowSep || "\n",
            re = new RegExp('[^' + fieldSep + ']'),
            opts = {
                FS: fieldSep,
                RS: rowSep
            },
            add_group_name = grouper(config.groupName || 'sheet', fieldSep);

        var filter = goodLine(re);
        var csv = '';

        // for each sheet
        for (var i = 0, s = sheets.length; i < s; i++) {
            // function for prefixing group
            var add_group = grouper(groups[i], fieldSep);
            // pull rows from sheet
            var rows = XL.utils.sheet_to_csv(workbook.Sheets[sheets[i]], opts).split('\n');
            var body;

            // on the first sheet, add the group name
            if (i === 0) {
                var firstline = add_group_name(rows[0]);
                var mutated = rows.splice(1).filter(filter).map(add_group);
                body = [firstline].concat(mutated);
                
            } else {
                body = rows.splice(rmLines).filter(filter).map(add_group);
            }

            csv += body.join('\n') + '\n';
        }

        callback(null, csv);

    } catch(e) {
        callback(e);
    }
}

module.exports = sheetstack;
