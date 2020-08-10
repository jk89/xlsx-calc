(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else {
		var a = factory();
		for(var i in a) (typeof exports === 'object' ? exports : root)[i] = a[i];
	}
})(this, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/index.js");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/Calculator.js":
/*!***************************!*\
  !*** ./src/Calculator.js ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const find_all_cells_with_formulas = __webpack_require__(/*! ./find_all_cells_with_formulas.js */ "./src/find_all_cells_with_formulas.js");

class Calculator {
    
    constructor(workbook, exec_formula) {
        this.workbook = workbook;
        this.expressions = [];
        this.exec_formula = exec_formula;
        this.variables = {};
        let formulas = find_all_cells_with_formulas(workbook, exec_formula);
        for (let i = formulas.length - 1; i >= 0; i--) {
            let exp = exec_formula.build_expression(formulas[i]);
            this.expressions.push(exp);
        }
    }
    
    setVar(var_name, value) {
        let variable = this.variables[var_name];
        if (variable) {
            variable.setValue(value);
        } else {
            this.expressions.forEach(exp => {
                this.setVarOfExpression(exp, var_name, value);
            });
        }
    }
    
    setVarOfExpression(exp, var_name, value) {
        for (let i = 0; i < exp.args.length; i++) {
            let arg = exp.args[i];
            if (arg === var_name) {
                exp.args[i] = this.variables[var_name] || (this.variables[var_name] = new RawValue(value));
            } else if (typeof arg === 'object' && (arg.name === 'Expression' || arg.name === 'UserFn')) {
                this.setVarOfExpression(arg, var_name, value);
            }
        }
    }
    
    execute() {
        let promises = this.expressions.map(exp => {
            return exp.update_cell_value();
        });
        return Promise.all(promises);
    }
}

module.exports = Calculator;

/***/ }),

/***/ "./src/Exp.js":
/*!********************!*\
  !*** ./src/Exp.js ***!
  \********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const RefValue = __webpack_require__(/*! ./RefValue.js */ "./src/RefValue.js");
const Range = __webpack_require__(/*! ./Range.js */ "./src/Range.js");
const resolve_promises = __webpack_require__(/*! ./resolve_promises.js */ "./src/resolve_promises.js");
let exp_id = 0;
let exec_id = 0;

module.exports = function Exp(formula) {
    var self = this;
    self.id = ++exp_id;
    self.args = [];
    self.name = 'Expression';
    self.update_cell_value = update_cell_value;
    self.formula = formula;
    
    function handleException(e, formula, resolve, reject) {
        if (e.message == '#N/A') {
            formula.cell.v = 42;
            formula.cell.t = 'e';
            formula.cell.w = e.message;
            resolve();
        }
        else {
            //console.error('Error', current_execution, e);
            reject(e);
            //throw e;
        }
    }
    
    function update_cell_value() {
        return new Promise((resolve, reject) => {
            //let current_execution = exec_id++;
            try {
                //console.log('Exec', current_execution, formula.name, formula.cell.f);
                var val_or_promise = self.calc();
                if (typeof val_or_promise === 'object' && typeof val_or_promise['then'] === 'function') {
                    val_or_promise.then(res => {
                        formula.cell.v = res;
                        if (typeof(formula.cell.v) === 'string') {
                            formula.cell.t = 's';
                        }
                        else if (typeof(formula.cell.v) === 'number') {
                            formula.cell.t = 'n';
                        }
                        resolve(formula.cell.v);
                    }).catch(e => {
                        //console.log('Exp', self.id, 'error:', e);
                        //reject(e);
                        handleException(e, formula, resolve, reject);
                    });
                }
                else {
                    formula.cell.v = val_or_promise;
                    if (typeof(formula.cell.v) === 'string') {
                        formula.cell.t = 's';
                    }
                    else if (typeof(formula.cell.v) === 'number') {
                        formula.cell.t = 'n';
                    }
                    resolve(formula.cell.v);
                }
            }
            catch (e) {
                handleException(e, formula, resolve, reject);
            }
            finally {
                formula.status = 'done';
            }
        });
    }
    
    function exec(op, args, fn) {
        for (var i = 0; i < args.length; i++) {
            if (args[i] === op) {
                try {
                    var r = fn(args[i - 1].calc(), args[i + 1].calc());
                    args.splice(i - 1, 3, new RawValue(r));
                    i--;
                }
                catch (e) {
                    console.error(e);
                    throw Error(formula.name + ': evaluating ' + formula.cell.f + '\n' + e.message);
                    //throw e;
                }
            }
        }
    }

    function exec_minus(args) {
        for (var i = args.length; i--;) {
            if (args[i] === '-') {
                var r = -args[i + 1].calc();
                if (typeof args[i - 1] !== 'string' && i > 0) {
                    args.splice(i, 1, '+');
                    args.splice(i + 1, 1, new RawValue(r));
                }
                else {
                    args.splice(i, 2, new RawValue(r));
                }
            }
        }
    }
    
    self.calc = function() {
        return new Promise((resolve, reject) => {
            resolve_promises(self.args.concat()).then(args => {
                try {
                    exec_minus(args);
                    exec('^', args, function(a, b) {
                        return Math.pow(+a, +b);
                    });
                    exec('*', args, function(a, b) {
                        return (+a) * (+b);
                    });
                    exec('/', args, function(a, b) {
                        return (+a) / (+b);
                    });
                    exec('+', args, function(a, b) {
                        return (+a) + (+b);
                    });
                    exec('&', args, function(a, b) {
                        return '' + a + b;
                    });
                    exec('<', args, function(a, b) {
                        return a < b;
                    });
                    exec('>', args, function(a, b) {
                        return a > b;
                    });
                    exec('>=', args, function(a, b) {
                        return a >= b;
                    });
                    exec('<=', args, function(a, b) {
                        return a <= b;
                    });
                    exec('<>', args, function(a, b) {
                        return a != b;
                    });
                    exec('=', args, function(a, b) {
                        return a == b;
                    });
                    if (args.length == 1) {
                        if (typeof args[0] === 'object' && typeof args[0]['then'] === 'function') {
                            args[0].then(resolve).catch(reject);
                            return;
                        }
                        if (typeof(args[0].calc) != 'function') {
                            return resolve(args[0]);
                        }
                        else {
                            return resolve(args[0].calc());
                        }
                    }
                    else {
                        console.log('something is not right');
                    }
                } catch(e) {
                    reject(e);
                }
            }).catch(reject);
        });
    };

    var last_arg;
    self.push = function(buffer) {
        if (buffer) {
            var v;
            if (!isNaN(buffer)) {
                v = new RawValue(+buffer);
            }
            else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
                v = new Range(buffer.trim().replace(/\$/g, ''), formula);
            }
            else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
                v = new Range(buffer.trim().replace(/\$/g, ''), formula);
            }
            else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+:[A-Z]+$/)) {
                v = new Range(buffer.trim().replace(/\$/g, ''), formula);
            }
            else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+:[A-Z]+$/)) {
                v = new Range(buffer.trim().replace(/\$/g, ''), formula);
            }
            else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+$/)) {
                v = new RefValue(buffer.trim().replace(/\$/g, ''), formula);
            }
            else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+$/)) {
                v = new RefValue(buffer.trim().replace(/\$/g, ''), formula);
            }
            else if (typeof buffer === 'string' && !isNaN(buffer.trim().replace(/%$/, ''))) {
                v = new RawValue(+(buffer.trim().replace(/%$/, '')) / 100.0);
            }
            else {
                v = buffer;
            }
            if (((v === '=') && (last_arg == '>' || last_arg == '<')) || (last_arg == '<' && v === '>')) {
                self.args[self.args.length - 1] += v;
            }
            else {
                self.args.push(v);
            }
            last_arg = v;
            //console.log(self.id, '-->', v);
        }
    };
};

/***/ }),

/***/ "./src/Range.js":
/*!**********************!*\
  !*** ./src/Range.js ***!
  \**********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const col_str_2_int = __webpack_require__(/*! ./col_str_2_int.js */ "./src/col_str_2_int.js");
const int_2_col_str = __webpack_require__(/*! ./int_2_col_str.js */ "./src/int_2_col_str.js");
const getSanitizedSheetName = __webpack_require__(/*! ./getSanitizedSheetName.js */ "./src/getSanitizedSheetName.js");

module.exports = function Range(str_expression, formula) {
    
    function promiseInSeq(sheet, matrix, min_row, max_row, min_col, max_col, sheet_name, resolve, reject, _row) {
        //console.log('min_row =', min_row, 'max_row =', max_row);
        //console.log('min_col =', min_col, 'max_col =', max_col);
        for (let i = min_row; i <= max_row; i++) {
            //console.log('ok i =', i);
            let row;
            if (_row) {
                row = _row;
            }
            else {
                row = [];
                matrix.push(row);
            }
            for (let j = min_col; j <= max_col; j++) {
                let cell_name = int_2_col_str(j) + i;
                let cell_full_name = sheet_name + '!' + cell_name;
                //console.log('range <<', cell_name, 'i =', i, 'j =', j);
                if (formula.formula_ref[cell_full_name]) {
                    if (formula.formula_ref[cell_full_name].status === 'working') {
                        //console.log('Circular ref in range');
                        reject('Circular ref');
                        return;
                    } else if (formula.formula_ref[cell_full_name].status === 'new') {
                        formula.exec_formula(formula.formula_ref[cell_full_name]).then(r=>{
                            row.push(sheet[cell_name].v);
                            //console.log('recursao asincrona');
                            j++;
                            if (j > max_col) {
                                j = 0;
                                i++;
                            }
                            promiseInSeq(sheet, matrix, i, max_row, j, max_col, sheet_name, resolve, reject, row);
                        }).catch(reject);
                        return;
                    } else if (formula.formula_ref[cell_full_name].status === 'done') {
                        row.push(sheet[cell_name].v);
                    }
                }
                else if (sheet[cell_name]) {
                    row.push(sheet[cell_name].v);
                }
                else {
                    row.push(null);
                }
            }
        }
        resolve();
    }
    
    this.calc = function() {
        return new Promise((resolve, reject) => {
            try {
                var range_expression, sheet_name, sheet;
                if (str_expression.indexOf('!') != -1) {
                    var aux = str_expression.split('!');
                    sheet_name = getSanitizedSheetName(aux[0]);
                    range_expression = aux[1];
                }
                else {
                    sheet_name = formula.sheet_name;
                    range_expression = str_expression;
                }
                sheet = formula.wb.Sheets[sheet_name];
                var arr = range_expression.split(':');
                var min_row = parseInt(arr[0].replace(/^[A-Z]+/, ''), 10) || 0;
                var str_max_row = arr[1].replace(/^[A-Z]+/, '');
                var max_row;
                if (str_max_row === '' && sheet['!ref']) {
                    str_max_row = sheet['!ref'].split(':')[1].replace(/^[A-Z]+/, '');
                }
                // the max is 1048576, but TLE
                max_row = parseInt(str_max_row == '' ? '500000' : str_max_row, 10);
                var min_col = col_str_2_int(arr[0]);
                var max_col = col_str_2_int(arr[1]);
                var matrix = [];
                new Promise((res, rej) => {
                    promiseInSeq(sheet, matrix, min_row, max_row, min_col, max_col, sheet_name, res, rej);
                }).then(() => {
                    resolve(matrix);
                }).catch(reject);
            } catch(e) {
                reject(e);
            }
        });
    };
};


/***/ }),

/***/ "./src/RawValue.js":
/*!*************************!*\
  !*** ./src/RawValue.js ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function RawValue(value) {
    this.setValue = function(v) {
        value = v;
    };
    this.calc = function() {
        return value;
    };
};


/***/ }),

/***/ "./src/RefValue.js":
/*!*************************!*\
  !*** ./src/RefValue.js ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const getSanitizedSheetName = __webpack_require__(/*! ./getSanitizedSheetName.js */ "./src/getSanitizedSheetName.js");

module.exports = function RefValue(str_expression, formula) {
    var self = this;
    this.name = 'RefValue';
    this.str_expression = str_expression;
    this.formula = formula;

    self.parseRef = function() {
        var sheet, sheet_name, cell_name, cell_full_name;
        if (str_expression.indexOf('!') != -1) {
            var aux = str_expression.split('!');
            sheet_name = getSanitizedSheetName(aux[0]);
            sheet = formula.wb.Sheets[sheet_name];
            cell_name = aux[1];
        }
        else {
            sheet = formula.sheet;
            sheet_name = formula.sheet_name;
            cell_name = str_expression;
        }
        if (!sheet) {
            throw Error("Sheet " + sheet_name + " not found.");
        }
        cell_full_name = sheet_name + '!' + cell_name;
        return {
            sheet: sheet,
            sheet_name: sheet_name,
            cell_name: cell_name,
            cell_full_name: cell_full_name
        };
    };

    this.calc = function() {
        return new Promise((resolve, reject) => {
            var resolved_ref = self.parseRef();
            var sheet = resolved_ref.sheet;
            var cell_name = resolved_ref.cell_name;
            var cell_full_name = resolved_ref.cell_full_name;
            var ref_cell = sheet[cell_name];
            if (!ref_cell) {
                return resolve(null);
            }
            var formula_ref = formula.formula_ref[cell_full_name];
            if (formula_ref) {
                if (formula_ref.status === 'new') {
                    formula.exec_formula(formula_ref).then(res => {
                        resolve(res);
                    }).catch(reject);
                    //return sheet[cell_name].v;
                }
                else if (formula_ref.status === 'working') {
                    //console.log('Circular ref in Ref');
                    reject('Circular ref');
                    //throw new Error('Circular ref');
                }
                else if (formula_ref.status === 'done') {
                    if (sheet[cell_name].t === 'e') {
                        reject(sheet[cell_name].w);
                        //console.log('ref is an error');
                        //throw new Error(sheet[cell_name].w);
                    }
                    return resolve(sheet[cell_name].v);
                }
            }
            else {
                return resolve(sheet[cell_name].v);
            }
        });
    };
};


/***/ }),

/***/ "./src/UserFnExecutor.js":
/*!*******************************!*\
  !*** ./src/UserFnExecutor.js ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const resolve_promises = __webpack_require__(/*! ./resolve_promises.js */ "./src/resolve_promises.js");

module.exports = function UserFnExecutor(user_function) {
    var self = this;
    self.name = 'UserFn';
    self.args = [];
    
    self.calc = function() {
        return new Promise((resolve, reject) => {
            resolve_promises(self.args).then(args => {
                try {
                    resolve(user_function.apply(self, args.map(f=>f.calc())));
                } catch(e) {
                    reject(e);
                }
            }).catch(reject);
        });
    };
    
    self.push = function(buffer) {
        self.args.push(buffer);
    };
    
};

/***/ }),

/***/ "./src/UserRawFnExecutor.js":
/*!**********************************!*\
  !*** ./src/UserRawFnExecutor.js ***!
  \**********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function UserRawFnExecutor(user_function) {
    var self = this;
    self.name = 'UserRawFn';
    self.args = [];
    self.calc = function() {
        return user_function.apply(self, self.args);
    };
    self.push = function(buffer) {
        self.args.push(buffer);
    };
};


/***/ }),

/***/ "./src/col_str_2_int.js":
/*!******************************!*\
  !*** ./src/col_str_2_int.js ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function col_str_2_int(col_str) {
    var r = 0;
    var colstr = col_str.replace(/[0-9]+$/, '');
    for (var i = colstr.length; i--;) {
        r += Math.pow(26, colstr.length - i - 1) * (colstr.charCodeAt(i) - 64);
    }
    return r - 1;
};

/***/ }),

/***/ "./src/exec_formula.js":
/*!*****************************!*\
  !*** ./src/exec_formula.js ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const Exp = __webpack_require__(/*! ./Exp.js */ "./src/Exp.js");
const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const UserFnExecutor = __webpack_require__(/*! ./UserFnExecutor.js */ "./src/UserFnExecutor.js");
const UserRawFnExecutor = __webpack_require__(/*! ./UserRawFnExecutor.js */ "./src/UserRawFnExecutor.js");

var xlsx_Fx = {};
var xlsx_raw_Fx = {};

import_functions(__webpack_require__(/*! ./formulas.js */ "./src/formulas.js"));
import_raw_functions(__webpack_require__(/*! ./formulas-raw.js */ "./src/formulas-raw.js"));

const common_operations = {
    '*': 'multiply',
    '+': 'plus',
    '-': 'minus',
    '/': 'divide',
    '^': 'power',
    '&': 'concat',
    '<': 'lt',
    '>': 'gt',
    '=': 'eq'
};

function import_raw_functions(functions, opts) {
    for (var key in functions) {
        xlsx_raw_Fx[key] = functions[key];
    }
}

function import_functions(formulajs, opts) {
    opts = opts || {};
    var prefix = opts.prefix || '';
    for (var key in formulajs) {
        var obj = formulajs[key];
        if (typeof(obj) === 'function') {
            xlsx_Fx[prefix + key] = obj;
        }
        else if (typeof(obj) === 'object') {
            import_functions(obj, my_assign(opts, { prefix: key + '.' }));
        }
    }
}

function my_assign(dest, source) {
    var obj = JSON.parse(JSON.stringify(dest));
    for (var k in source) {
        obj[k] = source[k];
    }
    return obj;
}

function build_expression(formula) {
    formula.status = 'working';
    var root_exp;
    var str_formula = formula.cell.f;
    if (str_formula[0] == '=') {
        str_formula = str_formula.substr(1);
    }
    var exp_obj = root_exp = new Exp(formula);
    var buffer = '',
        is_string = false,
        was_string = false;
    var fn_stack = [{
        exp: exp_obj
    }];
    for (var i = 0; i < str_formula.length; i++) {
        if (str_formula[i] == '"') {
            if (is_string) {
                exp_obj.push(new RawValue(buffer));
                is_string = false;
                was_string = true;
            }
            else {
                is_string = true;
            }
            buffer = '';
        }
        else if (is_string) {
            buffer += str_formula[i];
        }
        else if (str_formula[i] == '(') {
            var o, trim_buffer = buffer.trim(),
                special = xlsx_Fx[trim_buffer];
            var special_raw = xlsx_raw_Fx[trim_buffer];
            if (special_raw) {
                special = new UserRawFnExecutor(special_raw, formula);
            }
            else if (special) {
                special = new UserFnExecutor(special, formula);
            }
            else if (trim_buffer) {
                //Error: "Worksheet 1"!D145: Function INDEX not found
                throw new Error('"' + formula.sheet_name + '"!' + formula.name + ': Function ' + buffer + ' not found');
            }
            o = new Exp(formula);
            fn_stack.push({
                exp: o,
                special: special
            });
            exp_obj = o;
            buffer = '';
        }
        else if (common_operations[str_formula[i]]) {
            if (!was_string) {
                exp_obj.push(buffer);
            }
            was_string = false;
            exp_obj.push(str_formula[i]);
            buffer = '';
        }
        else if (str_formula[i] === ',' && fn_stack[fn_stack.length - 1].special) {
            was_string = false;
            fn_stack[fn_stack.length - 1].exp.push(buffer);
            fn_stack[fn_stack.length - 1].special.push(fn_stack[fn_stack.length - 1].exp);
            fn_stack[fn_stack.length - 1].exp = exp_obj = new Exp(formula);
            buffer = '';
        }
        else if (str_formula[i] == ')') {
            var v, stack = fn_stack.pop();
            exp_obj = stack.exp;
            exp_obj.push(buffer);
            v = exp_obj;
            buffer = '';
            exp_obj = fn_stack[fn_stack.length - 1].exp;
            if (stack.special) {
                stack.special.push(v);
                exp_obj.push(stack.special);
            }
            else {
                exp_obj.push(v);
            }
        }
        else {
            buffer += str_formula[i];
        }
    }
    root_exp.push(buffer);
    return root_exp;
}

function exec_formula(formula) {
    let root_exp = build_expression(formula);
    return root_exp.update_cell_value();
}

exec_formula.set_fx = function set_fx(name, fn) {
    xlsx_Fx[name] = fn;
};

exec_formula.exec_fx = function exec_fx(name, args) {
    return xlsx_Fx[name].apply(this, args);
};

exec_formula.import_functions = import_functions;
exec_formula.import_raw_functions = import_raw_functions;
exec_formula.build_expression = build_expression;
module.exports = exec_formula;


/***/ }),

/***/ "./src/find_all_cells_with_formulas.js":
/*!*********************************************!*\
  !*** ./src/find_all_cells_with_formulas.js ***!
  \*********************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function find_all_cells_with_formulas(wb, exec_formula) {
    var formula_ref = {};
    var cells = [];
    for (var sheet_name in wb.Sheets) {
        var sheet = wb.Sheets[sheet_name];
        for (var cell_name in sheet) {
            if (sheet[cell_name].f) {
                var formula = formula_ref[sheet_name + '!' + cell_name] = {
                    formula_ref: formula_ref,
                    wb: wb,
                    sheet: sheet,
                    sheet_name: sheet_name,
                    cell: sheet[cell_name],
                    name: cell_name,
                    status: 'new',
                    exec_formula: exec_formula
                };
                cells.push(formula);
            }
        }
    }
    return cells;
};


/***/ }),

/***/ "./src/formulas-raw.js":
/*!*****************************!*\
  !*** ./src/formulas-raw.js ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const int_2_col_str = __webpack_require__(/*! ./int_2_col_str.js */ "./src/int_2_col_str.js");
const col_str_2_int = __webpack_require__(/*! ./col_str_2_int.js */ "./src/col_str_2_int.js");
const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const Range = __webpack_require__(/*! ./Range.js */ "./src/Range.js");
const RefValue = __webpack_require__(/*! ./RefValue.js */ "./src/RefValue.js");

function raw_offset(cell_ref, rows, columns, height, width) {
    height = (height || new RawValue(1)).calc();
    width = (width || new RawValue(1)).calc();
    if (cell_ref.args.length === 1 && cell_ref.args[0].name === 'RefValue') {
        var ref_value = cell_ref.args[0];
        var parsed_ref = ref_value.parseRef();
        var col = col_str_2_int(parsed_ref.cell_name) + columns.calc();
        var col_str = int_2_col_str(col);
        var row = +parsed_ref.cell_name.replace(/^[A-Z]+/g, '') + rows.calc();
        var cell_name = col_str + row;
        if (height === 1 && width === 1) {
            return new RefValue(cell_name, ref_value.formula).calc();
        }
        else {
            var end_range_col = int_2_col_str(col + width - 1);
            var end_range_row = row + height - 1;
            var end_range = end_range_col + end_range_row;
            var str_expression = parsed_ref.sheet_name + '!' + cell_name + ':' + end_range;
            return new Range(str_expression, ref_value.formula).calc();
        }
    }
}

function resolveOnErrorValue(onerrorvalue, resolve, reject) {
    let v_or_promise = onerrorvalue.calc();
    if (typeof v_or_promise === 'object' && typeof v_or_promise['then'] === 'function') {
        //console.log('resolvendo onerrorvalue');
        v_or_promise.then(r => {
            //console.log('On error value =', r);
            resolve(r);
        }).catch(e=> {
            //console.log('Erro no on error value');
            reject(e);
        });
    } else {
        //console.log('valor no caso de erro =', v_or_promise);
        resolve(v_or_promise);
    }
}

function iferror(cell_ref, onerrorvalue) {
    return new Promise((resolve, reject) => {
        try {
            cell_ref.calc().then(value=>{
                //console.log('tudo ok com o cell_ref...', typeof value);
                if (typeof value === 'undefined') {
                    resolveOnErrorValue(onerrorvalue, resolve, reject);
                } else if (typeof value === 'number' && (isNaN(value) || value === Infinity || value === -Infinity)) {
                    resolveOnErrorValue(onerrorvalue, resolve, reject);
                } else {
                    resolve(value);
                }
            }).catch(e => {
                //console.log('2 error level');
                resolveOnErrorValue(onerrorvalue, resolve, reject);
            });
        } catch(e) {
            return onerrorvalue.calc();
        }
    });
}

module.exports = {
    'OFFSET': raw_offset,
    'IFERROR': iferror
};


/***/ }),

/***/ "./src/formulas.js":
/*!*************************!*\
  !*** ./src/formulas.js ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


// +---------------------+
// | FORMULAS REGISTERED |
// +---------------------+
let formulas = {
    'FLOOR': Math.floor,
    '_xlfn.FLOOR.MATH': Math.floor,
    'ABS': Math.abs,
    'SQRT': Math.sqrt,
    'VLOOKUP': vlookup,
    'MAX': max,
    'SUM': sum,
    'MIN': min,
    'CONCATENATE': concatenate,
    'IF': _if,
    'PMT': pmt,
    'COUNTA': counta,
    'IRR': irr,
    'NORM.INV': normsInv,
    '_xlfn.NORM.INV': normsInv,
    'STDEV': stDeviation,
    'AVERAGE': avg,
    'EXP': EXP,
    'LN': Math.log,
    '_xlfn.VAR.P': var_p,
    'VAR.P': var_p,
    '_xlfn.COVARIANCE.P': covariance_p,
    'COVARIANCE.P': covariance_p,
    'TRIM': trim,
    'LEN': len,
    'ISBLANK': is_blank
};

function len(a) {
    return ('' + a).length;
}

function trim(a) {
    return ('' + a).trim();
}

function is_blank(a) {
    return !a;
}

function covariance_p(a, b) {
    a = getArrayOfNumbers(a);
    b = getArrayOfNumbers(b);
    if (a.length != b.length) {
        return 'N/D';
    }
    var inv_n = 1.0 / a.length;
    var avg_a = sum.apply(this, a) / a.length;
    var avg_b = sum.apply(this, b) / b.length;
    var s = 0.0;
    for (var i = 0; i < a.length; i++) {
        s += (a[i] - avg_a) * (b[i] - avg_b);
    }
    return s * inv_n;
}

function getArrayOfNumbers(range) {
    var arr = [];
    for (var i = 0; i < range.length; i++) {
        var arg = range[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                if (typeof(matrix[j]) == 'number') {
                    arr.push(matrix[j]);
                }
                else if (Array.isArray(matrix[j])) {
                    for (var k = matrix[j].length; k--;) {
                        if (typeof(matrix[j][k]) == 'number') {
                            arr.push(matrix[j][k]);
                        }
                    }
                }
                // else {
                //   wtf is that?
                // }
            }
        }
        else {
            if (typeof(arg) == 'number') {
                arr.push(arg);
            }
        }
    }
    return arr;
}

function var_p() {
    var average = avg.apply(this, arguments);
    var s = 0.0;
    var c = 0;
    for (var i = 0; i < arguments.length; i++) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
                        s += Math.pow(matrix[j][k] - average, 2);
                        c++;
                    }
                }
            }
        }
        else {
            s += Math.pow(arg - average, 2);
            c++;
        }
    }
    return s / c;
}

function EXP(n) {
    return Math.pow(Math.E, n);
}

function avg() {
    return sum.apply(this, arguments) / counta.apply(this, arguments);
}

function stDeviation() {
    var array = getArrayOfNumbers(arguments);

    function _mean(array) {
        return array.reduce(function(a, b) {
            return a + b;
        }) / array.length;
    }
    var mean = _mean(array),
        dev = array.map(function(itm) {
            return (itm - mean) * (itm - mean);
        });
    return Math.sqrt(dev.reduce(function(a, b) {
        return a + b;
    }) / (array.length - 1));
}

/// Original C++ implementation found at http://www.wilmott.com/messageview.cfm?catid=10&threadid=38771
/// C# implementation found at http://weblogs.asp.net/esanchez/archive/2010/07/29/a-quick-and-dirty-implementation-of-excel-norminv-function-in-c.aspx
/*
 *     Compute the quantile function for the normal distribution.
 *
 *     For small to moderate probabilities, algorithm referenced
 *     below is used to obtain an initial approximation which is
 *     polished with a final Newton step.
 *
 *     For very large arguments, an algorithm of Wichura is used.
 *
 *  REFERENCE
 *
 *     Beasley, J. D. and S. G. Springer (1977).
 *     Algorithm AS 111: The percentage points of the normal distribution,
 *     Applied Statistics, 26, 118-121.
 *
 *      Wichura, M.J. (1988).
 *      Algorithm AS 241: The Percentage Points of the Normal Distribution.
 *      Applied Statistics, 37, 477-484.
 */
function normsInv(p, mu, sigma) {
    if (p < 0 || p > 1) {
        throw "The probality p must be bigger than 0 and smaller than 1";
    }
    if (sigma < 0) {
        throw "The standard deviation sigma must be positive";
    }

    if (p == 0) {
        return -Infinity;
    }
    if (p == 1) {
        return Infinity;
    }
    if (sigma == 0) {
        return mu;
    }

    var q, r, val;

    q = p - 0.5;

    /*-- use AS 241 --- */
    /* double ppnd16_(double *p, long *ifault)*/
    /*      ALGORITHM AS241  APPL. STATIST. (1988) VOL. 37, NO. 3
            Produces the normal deviate Z corresponding to a given lower
            tail area of P; Z is accurate to about 1 part in 10**16.
    */
    if (Math.abs(q) <= .425) { /* 0.075 <= p <= 0.925 */
        r = .180625 - q * q;
        val =
            q * (((((((r * 2509.0809287301226727 +
                            33430.575583588128105) * r + 67265.770927008700853) * r +
                        45921.953931549871457) * r + 13731.693765509461125) * r +
                    1971.5909503065514427) * r + 133.14166789178437745) * r +
                3.387132872796366608) / (((((((r * 5226.495278852854561 +
                        28729.085735721942674) * r + 39307.89580009271061) * r +
                    21213.794301586595867) * r + 5394.1960214247511077) * r +
                687.1870074920579083) * r + 42.313330701600911252) * r + 1);
    }
    else { /* closer than 0.075 from {0,1} boundary */

        /* r = min(p, 1-p) < 0.075 */
        if (q > 0)
            r = 1 - p;
        else
            r = p;

        r = Math.sqrt(-Math.log(r));
        /* r = sqrt(-log(r))  <==>  min(p, 1-p) = exp( - r^2 ) */

        if (r <= 5) { /* <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11 */
            r += -1.6;
            val = (((((((r * 7.7454501427834140764e-4 +
                                .0227238449892691845833) * r + .24178072517745061177) *
                            r + 1.27045825245236838258) * r +
                        3.64784832476320460504) * r + 5.7694972214606914055) *
                    r + 4.6303378461565452959) * r +
                1.42343711074968357734) / (((((((r *
                                1.05075007164441684324e-9 + 5.475938084995344946e-4) *
                            r + .0151986665636164571966) * r +
                        .14810397642748007459) * r + .68976733498510000455) *
                    r + 1.6763848301838038494) * r +
                2.05319162663775882187) * r + 1);
        }
        else { /* very close to  0 or 1 */
            r += -5;
            val = (((((((r * 2.01033439929228813265e-7 +
                                2.71155556874348757815e-5) * r +
                            .0012426609473880784386) * r + .026532189526576123093) *
                        r + .29656057182850489123) * r +
                    1.7848265399172913358) * r + 5.4637849111641143699) *
                r + 6.6579046435011037772) / (((((((r *
                            2.04426310338993978564e-15 + 1.4215117583164458887e-7) *
                        r + 1.8463183175100546818e-5) * r +
                    7.868691311456132591e-4) * r + .0148753612908506148525) * r + .13692988092273580531) * r +
                .59983220655588793769) * r + 1);
        }

        if (q < 0.0) {
            val = -val;
        }
    }

    return mu + sigma * val;
}

function irr(range, guess) {
    var min = -2.0;
    var max = 1.0;
    var n = 0;
    do {
        var guest = (min + max) / 2;
        var NPV = 0;
        for (var i = 0; i < range.length; i++) {
            var arg = range[i];
            NPV += arg[0] / Math.pow((1 + guest), i);
        }
        if (NPV > 0) {
            if (min === max) {
                max += Math.abs(guest);
            }
            min = guest;
        }
        else {
            max = guest;
        }
        n++;
    } while (Math.abs(NPV) > 0.000001 && n < 100000);
    //console.log(n);
    return guest;
}

function counta() {
    var r = 0;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
                        r++;
                    }
                }
            }
        }
        else {
            if (arg !== null && arg !== undefined) {
                r++;
            }
        }
    }
    return r;
}

function pmt(rate_per_period, number_of_payments, present_value, future_value, type) {
    type = type || 0;
    future_value = future_value || 0;
    if (rate_per_period != 0.0) {
        // Interest rate exists
        var q = Math.pow(1 + rate_per_period, number_of_payments);
        return -(rate_per_period * (future_value + (q * present_value))) / ((-1 + q) * (1 + rate_per_period * (type)));

    }
    else if (number_of_payments != 0.0) {
        // No interest rate, but number of payments exists
        return -(future_value + present_value) / number_of_payments;
    }
    return 0;
}

function _if(condition, _then, _else) {
    if (condition) {
        return _then || 0;
    }
    else {
        return _else || 0;
    }
}

function concatenate() {
    var r = '';
    for (var i = 0; i < arguments.length; i++) {
        r += arguments[i];
    }
    return r;
}

function sum() {
    //console.log('sum...');
    var r = 0;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (!isNaN(matrix[j][k])) {
                        r += +matrix[j][k];
                    }
                }
            }
        }
        else {
            r += +arg;
        }
    }
    //console.log('end sum.');
    return r;
}

function max() {
    var max = null;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var arr = arg;
            for (var j = arr.length; j--;) {
                max = max == null || max < arr[j] ? arr[j] : max;
            }
        }
        else if (!isNaN(arg)) {
            max = max == null || max < arg ? arg : max;
        }
        else {
            console.log('WTF??', arg);
        }
    }
    return max;
}

function min() {
    var result = null;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var arr = arg;
            for (var j = arr.length; j--;) {
                result = result == null || result > arr[j] ? arr[j] : result;
            }
        }
        else if (!isNaN(arg)) {
            result = result == null || result > arg ? arg : result;
        }
        else {
            console.log('WTF??', arg);
        }
    }
    return result;
}

function vlookup(key, matrix, return_index) {
    for (var i = 0; i < matrix.length; i++) {
        if (matrix[i][0] == key) {
            return matrix[i][return_index - 1];
        }
    }
    throw Error('#N/A');
}

module.exports = formulas;


/***/ }),

/***/ "./src/getSanitizedSheetName.js":
/*!**************************************!*\
  !*** ./src/getSanitizedSheetName.js ***!
  \**************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function getSanitizedSheetName(sheet_name) {
    var quotedMatch = sheet_name.match(/^'(.*)'$/);
    if (quotedMatch) {
        return quotedMatch[1];
    }
    else {
        return sheet_name;
    }
};


/***/ }),

/***/ "./src/index.js":
/*!**********************!*\
  !*** ./src/index.js ***!
  \**********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const int_2_col_str = __webpack_require__(/*! ./int_2_col_str.js */ "./src/int_2_col_str.js");
const col_str_2_int = __webpack_require__(/*! ./col_str_2_int.js */ "./src/col_str_2_int.js");
const exec_formula = __webpack_require__(/*! ./exec_formula.js */ "./src/exec_formula.js");
const find_all_cells_with_formulas = __webpack_require__(/*! ./find_all_cells_with_formulas.js */ "./src/find_all_cells_with_formulas.js");
const Calculator = __webpack_require__(/*! ./Calculator.js */ "./src/Calculator.js");


function exec_next(formulas, i, resolve, reject) {
    if (i === formulas.length) {
        resolve();
        return;
    }
    //console.log('executing', i+1, 'of', formulas.length,'...');
    exec_formula(formulas[i]).then(x => {
        exec_next(formulas, i + 1, resolve, reject);
    }).catch(err => {
        //console.error(err);
        reject(err);
    });
}

var XLSX_CALC = function(workbook) {
    return new Promise((resolve, reject) => {
        var formulas = find_all_cells_with_formulas(workbook, exec_formula);
        exec_next(formulas, 0, resolve, reject);
        //for (var i = formulas.length - 1; i >= 0; i--) {
        //    exec_formula(formulas[i]);
        //}
        //resolve();
    });
};

XLSX_CALC.calculator = function calculator(workbook) {
    return new Calculator(workbook, exec_formula);
};

XLSX_CALC.set_fx = exec_formula.set_fx;
XLSX_CALC.exec_fx = exec_formula.exec_fx;
XLSX_CALC.col_str_2_int = col_str_2_int;
XLSX_CALC.int_2_col_str = int_2_col_str;
XLSX_CALC.import_functions = exec_formula.import_functions;
XLSX_CALC.import_raw_functions = exec_formula.import_raw_functions;



module.exports = XLSX_CALC;

/***/ }),

/***/ "./src/int_2_col_str.js":
/*!******************************!*\
  !*** ./src/int_2_col_str.js ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function int_2_col_str(n) {
    var dividend = n + 1;
    var columnName = '';
    var modulo;
    var guard = 10;
    while (dividend > 0 && guard--) {
        modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(modulo + 65) + columnName;
        dividend = (dividend - modulo - 1) / 26;
    }
    return columnName;
};

/***/ }),

/***/ "./src/resolve_promises.js":
/*!*********************************!*\
  !*** ./src/resolve_promises.js ***!
  \*********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");

module.exports = function resolve_promises(args) {
    //console.log('resolvendo promises...');
    return new Promise((resolve, reject) => {
        let promises = [],
            new_args = [];
        for (let i = 0; i < args.length; i++) {
            let arg = args[i];
            if (typeof arg === 'object' && typeof arg['calc'] === 'function') {
                let val_or_promise = arg.calc();
                if (typeof val_or_promise === 'object' && typeof val_or_promise['then'] === 'function') {
                    promises.push(val_or_promise);
                    val_or_promise.then(r => {
                        new_args[i] = new RawValue(r);
                    }).catch(reject);
                }
                else {
                    new_args[i] = new RawValue(val_or_promise);
                }
            }
            else {
                new_args[i] = arg;
            }
        }
        Promise.all(promises).then(() => {
            //console.log('new_args =', new_args.map(a => a.calc()));
            resolve(new_args);
        }).catch(reject);
    });
};

/***/ })

/******/ });
});
//# sourceMappingURL=xlsx-calc.js.map