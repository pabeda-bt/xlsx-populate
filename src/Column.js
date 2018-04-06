"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var ArgHandler = require("./ArgHandler");
var addressConverter = require('./addressConverter');

// Default column width.
var defaultColumnWidth = 9.140625;

/**
 * A column.
 */

var Column = function () {
    // /**
    //  * Creates a new Column.
    //  * @param {Sheet} sheet - The parent sheet.
    //  * @param {{}} node - The column node.
    //  * @constructor
    //  * @ignore
    //  * @private
    //  */
    function Column(sheet, node) {
        _classCallCheck(this, Column);

        this._sheet = sheet;
        this._node = node;
    }

    /* PUBLIC */

    /**
     * Get the address of the column.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.anchored] - Anchor the address.
     * @returns {string} The address
     */


    _createClass(Column, [{
        key: "address",
        value: function address(opts) {
            return addressConverter.toAddress({
                type: 'column',
                columnName: this.columnName(),
                sheetName: opts && opts.includeSheetName && this.sheet().name(),
                columnAnchored: opts && opts.anchored
            });
        }

        /**
         * Get a cell within the column.
         * @param {number} rowNumber - The row number.
         * @returns {Cell} The cell in the column with the given row number.
         */

    }, {
        key: "cell",
        value: function cell(rowNumber) {
            return this.sheet().cell(rowNumber, this.columnNumber());
        }

        /**
         * Get the name of the column.
         * @returns {string} The column name.
         */

    }, {
        key: "columnName",
        value: function columnName() {
            return addressConverter.columnNumberToName(this.columnNumber());
        }

        /**
         * Get the number of the column.
         * @returns {number} The column number.
         */

    }, {
        key: "columnNumber",
        value: function columnNumber() {
            return this._node.attributes.min;
        }

        /**
         * Gets a value indicating whether the column is hidden.
         * @returns {boolean} A flag indicating whether the column is hidden.
         */ /**
            * Sets whether the column is hidden.
            * @param {boolean} hidden - A flag indicating whether to hide the column.
            * @returns {Column} The column.
            */

    }, {
        key: "hidden",
        value: function hidden() {
            var _this = this;

            return new ArgHandler("Column.hidden").case(function () {
                return _this._node.attributes.hidden === 1;
            }).case('boolean', function (hidden) {
                if (hidden) _this._node.attributes.hidden = 1;else delete _this._node.attributes.hidden;
                return _this;
            }).handle(arguments);
        }

        /**
         * Get the parent sheet.
         * @returns {Sheet} The parent sheet.
         */

    }, {
        key: "sheet",
        value: function sheet() {
            return this._sheet;
        }

        /**
         * Gets an individual style.
         * @param {string} name - The name of the style.
         * @returns {*} The style.
         */ /**
            * Gets multiple styles.
            * @param {Array.<string>} names - The names of the style.
            * @returns {object.<string, *>} Object whose keys are the style names and values are the styles.
            */ /**
               * Sets an individual style.
               * @param {string} name - The name of the style.
               * @param {*} value - The value to set.
               * @returns {Cell} The cell.
               */ /**
                  * Sets multiple styles.
                  * @param {object.<string, *>} styles - Object whose keys are the style names and values are the styles to set.
                  * @returns {Cell} The cell.
                  */ /**
                     * Sets to a specific style
                     * @param {Style} style - Style object given from stylesheet.createStyle
                     * @returns {Cell} The cell.
                     */

    }, {
        key: "style",
        value: function style() {
            var _this2 = this;

            return new ArgHandler("Column.style").case('string', function (name) {
                // Get single value
                _this2._createStyleIfNeeded();
                return _this2._style.style(name);
            }).case('array', function (names) {
                // Get list of values
                var values = {};
                names.forEach(function (name) {
                    values[name] = _this2.style(name);
                });

                return values;
            }).case(['string', '*'], function (name, value) {
                // If a row node is already defined that intersects with this column and that row has a style set, we
                // need to make sure that a cell node exists at the intersection so we can style it appropriately.
                // Fetching the cell will force a new cell node to be created with a style matching the column. So we
                // will fetch and style the cell at each row that intersects this column if it is already present or it
                // has a style defined.
                _this2.sheet().forEachExistingRow(function (row) {
                    if (row.hasStyle() || row.hasCell(_this2.columnNumber())) {
                        row.cell(_this2.columnNumber()).style(name, value);
                    }
                });

                // Set a single value for all cells to a single value
                _this2._createStyleIfNeeded();
                _this2._style.style(name, value);

                return _this2;
            }).case('object', function (nameValues) {
                // Object of key value pairs to set
                for (var name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    var value = nameValues[name];
                    _this2.style(name, value);
                }

                return _this2;
            }).case('Style', function (style) {
                // See Large Comment Above
                _this2.sheet().forEachExistingRow(function (row) {
                    if (row.hasStyle() || row.hasCell(_this2.columnNumber())) {
                        row.cell(_this2.columnNumber()).style(style);
                    }
                });

                _this2._style = style;
                _this2._node.attributes.style = style.id();

                return _this2;
            }).handle(arguments);
        }

        /**
         * Gets the width.
         * @returns {undefined|number} The width (or undefined).
         */ /**
            * Sets the width.
            * @param {number} width - The width of the column.
            * @returns {Column} The column.
            */

    }, {
        key: "width",
        value: function width(_width) {
            var _this3 = this;

            return new ArgHandler("Column.width").case(function () {
                return _this3._node.attributes.customWidth ? _this3._node.attributes.width : undefined;
            }).case('number', function (width) {
                _this3._node.attributes.width = width;
                _this3._node.attributes.customWidth = 1;
                return _this3;
            }).case('nil', function () {
                delete _this3._node.attributes.width;
                delete _this3._node.attributes.customWidth;
                return _this3;
            }).handle(arguments);
        }

        /**
         * Get the parent workbook.
         * @returns {Workbook} The parent workbook.
         */

    }, {
        key: "workbook",
        value: function workbook() {
            return this.sheet().workbook();
        }

        /* INTERNAL */

        /**
         * Convert the column to an XML object.
         * @returns {{}} The XML form.
         * @ignore
         */

    }, {
        key: "toXml",
        value: function toXml() {
            return this._node;
        }

        /* PRIVATE */

        /**
         * Create a style for this column if it doesn't already exist.
         * @returns {undefined}
         * @private
         */

    }, {
        key: "_createStyleIfNeeded",
        value: function _createStyleIfNeeded() {
            if (!this._style) {
                var styleId = this._node.attributes.style;
                this._style = this.workbook().styleSheet().createStyle(styleId);
                this._node.attributes.style = this._style.id();

                if (!this.width()) this.width(defaultColumnWidth);
            }
        }
    }]);

    return Column;
}();

module.exports = Column;