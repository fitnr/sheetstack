var sheetstack = require('../index');
var j = require('j');

var callback = function(err, data) {
    if (err) console.error(err);
    console.assert(data.indexOf('dogs') > -1);
    console.assert(data.indexOf('cats') > -1);
    console.assert(data.indexOf('ONE FAMILY HOMES') > -1);
    console.assert(data.indexOf('sdsd') > -1);
};

console.assert(typeof(sheetstack) == 'function');

var contents = j.readFile(__dirname + '/test.xlsx');

sheetstack(contents, {}, callback);
