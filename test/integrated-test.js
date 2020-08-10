var assert = require('assert');
var XLSX = require('xlsx');
var XLSX_CALC = require("../");

describe('XLSX with XLSX_CALC', function() {

    function assert_values(sheet_expected, sheet_calculated) {
        for (var prop in sheet_expected) {
            if(prop.match(/[A-Z]+[0-9]+/)) {
                assert.equal(sheet_expected[prop].v, sheet_calculated[prop].v, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"\nexpected ' + sheet_expected[prop].v + " got " + sheet_calculated[prop].v);
                assert.equal(sheet_expected[prop].w, sheet_calculated[prop].w, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
                assert.equal(sheet_expected[prop].t, sheet_calculated[prop].t, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
            }
        }
    }

    function erase_values_that_contains_formula(sheets) {
        for (var sheet in sheets) {
            for (var prop in sheet) {
                if(prop.match(/[A-Z]+[0-9]+/) && sheet[prop].f) {
                    sheet[prop].v = null;
                }
            }
        }
    }

    it('recalc the workbook Sheet1', async function() {
        var workbook = XLSX.readFile('test/testcase.xlsx');
        erase_values_that_contains_formula(workbook.Sheets);
        var original_sheet = XLSX.readFile('test/testcase.xlsx').Sheets.Sheet1;
        await XLSX_CALC(workbook);
        assert_values(original_sheet, workbook.Sheets.Sheet1);
    });
    
    it('recalc the workbook Sheet OffSet', async function() {
        var workbook = XLSX.readFile('test/testcase.xlsx');
        erase_values_that_contains_formula(workbook.Sheets);
        var original_sheet = XLSX.readFile('test/testcase.xlsx').Sheets.OffSet;
        await XLSX_CALC(workbook);
        assert_values(original_sheet, workbook.Sheets.OffSet);
    });

});
