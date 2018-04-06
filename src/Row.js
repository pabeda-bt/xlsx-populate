"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");
var Cell = require("./Cell");
var regexify = require("./regexify");
var ArgHandler = require("./ArgHandler");
var addressConverter = require('./addressConverter');

/**
 * A row.
 */

var Row = function () {
    // /**
    //  * Creates a new instance of Row.
    //  * @param {Sheet} sheet - The parent sheet.
    //  * @param {{}} node - The row node.
    //  */
    function Row(sheet, node) {
        _classCallCheck(this, Row);

        this._sheet = sheet;
        this._init(node);
    }

    /* PUBLIC */

    /**
     * Get the address of the row.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.anchored] - Anchor the address.
     * @returns {string} The address
     */


    _createClass(Row, [{
        key: "address",
        value: function address(opts) {
            return addressConverter.toAddress({
                type: 'row',
                rowNumber: this.rowNumber(),
                sheetName: opts && opts.includeSheetName && this.sheet().name(),
                rowAnchored: opts && opts.anchored
            });
        }

        /**
         * Get a cell in the row.
         * @param {string|number} columnNameOrNumber - The name or number of the column.
         * @returns {Cell} The cell.
         */

    }, {
        key: "cell",
        value: function cell(columnNameOrNumber) {
            var columnNumber = columnNameOrNumber;
            if (typeof columnNameOrNumber === 'string') {
                columnNumber = addressConverter.columnNameToNumber(columnNameOrNumber);
            }

            // Return an existing cell.
            if (this._cells[columnNumber]) return this._cells[columnNumber];

            // No cell exists for this.
            // Check if there is an existing row/column style for the new cell.
            var styleId = void 0;
            var rowStyleId = this._node.attributes.s;
            var columnStyleId = this.sheet().existingColumnStyleId(columnNumber);

            // Row style takes priority. If a cell has both row and column styles it should have created a cell entry with a cell-specific style.
            if (!_.isNil(rowStyleId)) styleId = rowStyleId;else if (!_.isNil(columnStyleId)) styleId = columnStyleId;

            // Create the new cell.
            var cell = new Cell(this, columnNumber, styleId);
            this._cells[columnNumber] = cell;
            return cell;
        }

        /**
         * Gets the row height.
         * @returns {undefined|number} The height (or undefined).
         */ /**
            * Sets the row height.
            * @param {number} height - The height of the row.
            * @returns {Row} The row.
            */

    }, {
        key: "height",
        value: function height() {
            var _this = this;

            return new ArgHandler('Row.height').case(function () {
                return _this._node.attributes.customHeight ? _this._node.attributes.ht : undefined;
            }).case('number', function (height) {
                _this._node.attributes.ht = height;
                _this._node.attributes.customHeight = 1;
                return _this;
            }).case('nil', function () {
                delete _this._node.attributes.ht;
                delete _this._node.attributes.customHeight;
                return _this;
            }).handle(arguments);
        }

        /**
         * Gets a value indicating whether the row is hidden.
         * @returns {boolean} A flag indicating whether the row is hidden.
         */ /**
            * Sets whether the row is hidden.
            * @param {boolean} hidden - A flag indicating whether to hide the row.
            * @returns {Row} The row.
            */

    }, {
        key: "hidden",
        value: function hidden() {
            var _this2 = this;

            return new ArgHandler("Row.hidden").case(function () {
                return _this2._node.attributes.hidden === 1;
            }).case('boolean', function (hidden) {
                if (hidden) _this2._node.attributes.hidden = 1;else delete _this2._node.attributes.hidden;
                return _this2;
            }).handle(arguments);
        }

        /**
         * Gets the row number.
         * @returns {number} The row number.
         */

    }, {
        key: "rowNumber",
        value: function rowNumber() {
            return this._node.attributes.r;
        }

        /**
         * Gets the parent sheet of the row.
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
            var _this3 = this;

            return new ArgHandler("Row.style").case('string', function (name) {
                // Get single value
                _this3._createStyleIfNeeded();
                return _this3._style.style(name);
            }).case('array', function (names) {
                // Get list of values
                var values = {};
                names.forEach(function (name) {
                    values[name] = _this3.style(name);
                });

                return values;
            }).case(['string', '*'], function (name, value) {
                _this3._createCellStylesIfNeeded();

                // Style each existing cell within this row. (Cells don't inherit ow/column styles.)
                _.forEach(_this3._cells, function (cell) {
                    if (cell) cell.style(name, value);
                });

                // Set the style on the row.
                _this3._createStyleIfNeeded();
                _this3._style.style(name, value);

                return _this3;
            }).case('object', function (nameValues) {
                // Object of key value pairs to set
                for (var name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    var value = nameValues[name];
                    _this3.style(name, value);
                }

                return _this3;
            }).case('Style', function (style) {
                _this3._createCellStylesIfNeeded();

                // Style each existing cell within this row. (Cells don't inherit ow/column styles.)
                _.forEach(_this3._cells, function (cell) {
                    if (cell) cell.style(style);
                });

                _this3._style = style;
                _this3._node.attributes.s = style.id();
                _this3._node.attributes.customFormat = 1;

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
         * Clear cells that are using a given shared formula ID.
         * @param {number} sharedFormulaId - The shared formula ID.
         * @returns {undefined}
         * @ignore
         */

    }, {
        key: "clearCellsUsingSharedFormula",
        value: function clearCellsUsingSharedFormula(sharedFormulaId) {
            this._cells.forEach(function (cell) {
                if (!cell) return;
                if (cell.sharesFormula(sharedFormulaId)) cell.clear();
            });
        }

        /**
         * Find a pattern in the row and optionally replace it.
         * @param {string|RegExp} pattern - The search pattern.
         * @param {string} [replacement] - The replacement text.
         * @returns {Array.<Cell>} The matched cells.
         * @ignore
         */

    }, {
        key: "find",
        value: function find(pattern, replacement) {
            pattern = regexify(pattern);

            var matches = [];
            this._cells.forEach(function (cell) {
                if (!cell) return;
                if (cell.find(pattern, replacement)) matches.push(cell);
            });

            return matches;
        }

        /**
         * Check if the row has a cell at the given column number.
         * @param {number} columnNumber - The column number.
         * @returns {boolean} True if a cell exists, false otherwise.
         * @ignore
         */

    }, {
        key: "hasCell",
        value: function hasCell(columnNumber) {
            return !!this._cells[columnNumber];
        }

        /**
         * Check if the column has a style defined.
         * @returns {boolean} True if a style exists, false otherwise.
         * @ignore
         */

    }, {
        key: "hasStyle",
        value: function hasStyle() {
            return !_.isNil(this._node.attributes.s);
        }

        /**
         * Returns the nax used column number.
         * @returns {number} The max used column number.
         * @ignore
         */

    }, {
        key: "minUsedColumnNumber",
        value: function minUsedColumnNumber() {
            return _.findIndex(this._cells);
        }

        /**
         * Returns the nax used column number.
         * @returns {number} The max used column number.
         * @ignore
         */

    }, {
        key: "maxUsedColumnNumber",
        value: function maxUsedColumnNumber() {
            return this._cells.length - 1;
        }

        /**
         * Convert the row to an object.
         * @returns {{}} The object form.
         * @ignore
         */

    }, {
        key: "toXml",
        value: function toXml() {
            return this._node;
        }

        /* PRIVATE */

        /**
         * If a column node is already defined that intersects with this row and that column has a style set, we
         * need to make sure that a cell node exists at the intersection so we can style it appropriately.
         * Fetching the cell will force a new cell node to be created with a style matching the column.
         * @returns {undefined}
         * @private
         */

    }, {
        key: "_createCellStylesIfNeeded",
        value: function _createCellStylesIfNeeded() {
            var _this4 = this;

            this.sheet().forEachExistingColumnNumber(function (columnNumber) {
                if (!_.isNil(_this4.sheet().existingColumnStyleId(columnNumber))) _this4.cell(columnNumber);
            });
        }

        /**
         * Create a style for this row if it doesn't already exist.
         * @returns {undefined}
         * @private
         */

    }, {
        key: "_createStyleIfNeeded",
        value: function _createStyleIfNeeded() {
            if (!this._style) {
                var styleId = this._node.attributes.s;
                this._style = this.workbook().styleSheet().createStyle(styleId);
                this._node.attributes.s = this._style.id();
                this._node.attributes.customFormat = 1;
            }
        }

        /**
         * Initialize the row node.
         * @param {{}} node - The row node.
         * @returns {undefined}
         * @private
         */

    }, {
        key: "_init",
        value: function _init(node) {
            var _this5 = this;

            this._node = node;
            this._cells = [];
            this._node.children.forEach(function (cellNode) {
                var cell = new Cell(_this5, cellNode);
                _this5._cells[cell.columnNumber()] = cell;
            });
            this._node.children = this._cells;
        }
    }]);

    return Row;
}();

module.exports = Row;

/*
<row r="6" spans="1:9" x14ac:dyDescent="0.25">
    <c r="A6" s="1" t="s">
        <v>2</v>
    </c>
    <c r="B6" s="1"/>
    <c r="C6" s="1"/>
</row>
*/