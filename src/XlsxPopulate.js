"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var externals = require("./externals");
var Workbook = require("./Workbook");
var FormulaError = require("./FormulaError");
var dateConverter = require("./dateConverter");

/**
 * xlsx-poulate namespace.
 * @namespace
 */

var XlsxPopulate = function () {
    function XlsxPopulate() {
        _classCallCheck(this, XlsxPopulate);
    }

    _createClass(XlsxPopulate, null, [{
        key: "dateToNumber",

        /**
         * Convert a date to a number for Excel.
         * @param {Date} date - The date.
         * @returns {number} The number.
         */
        value: function dateToNumber(date) {
            return dateConverter.dateToNumber(date);
        }

        /**
         * Create a new blank workbook.
         * @returns {Promise.<Workbook>} The workbook.
         */

    }, {
        key: "fromBlankAsync",
        value: function fromBlankAsync() {
            return Workbook.fromBlankAsync();
        }

        /**
         * Loads a workbook from a data object. (Supports any supported [JSZip data types]{@link https://stuk.github.io/jszip/documentation/api_jszip/load_async.html}.)
         * @param {string|Array.<number>|ArrayBuffer|Uint8Array|Buffer|Blob|Promise.<*>} data - The data to load.
         * @param {{}} [opts] - Options
         * @param {string} [opts.password] - The password to decrypt the workbook.
         * @returns {Promise.<Workbook>} The workbook.
         */

    }, {
        key: "fromDataAsync",
        value: function fromDataAsync(data, opts) {
            return Workbook.fromDataAsync(data, opts);
        }

        /**
         * Loads a workbook from file.
         * @param {string} path - The path to the workbook.
         * @param {{}} [opts] - Options
         * @param {string} [opts.password] - The password to decrypt the workbook.
         * @returns {Promise.<Workbook>} The workbook.
         */

    }, {
        key: "fromFileAsync",
        value: function fromFileAsync(path, opts) {
            return Workbook.fromFileAsync(path, opts);
        }

        /**
         * Convert an Excel number to a date.
         * @param {number} number - The number.
         * @returns {Date} The date.
         */

    }, {
        key: "numberToDate",
        value: function numberToDate(number) {
            return dateConverter.numberToDate(number);
        }

        /**
         * The Promise library.
         * @type {Promise}
         */

    }, {
        key: "Promise",
        get: function get() {
            return externals.Promise;
        },
        set: function set(Promise) {
            externals.Promise = Promise;
        }
    }]);

    return XlsxPopulate;
}();

/**
 * The XLSX mime type.
 * @type {string}
 */


XlsxPopulate.MIME_TYPE = Workbook.MIME_TYPE;

/**
 * Formula error class.
 * @type {FormulaError}
 */
XlsxPopulate.FormulaError = FormulaError;

module.exports = XlsxPopulate;