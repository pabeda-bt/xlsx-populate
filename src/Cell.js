"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");
var ArgHandler = require("./ArgHandler");
var addressConverter = require("./addressConverter");
var dateConverter = require("./dateConverter");
var regexify = require("./regexify");
var xmlq = require("./xmlq");
var FormulaError = require("./FormulaError");
var Style = require("./Style");

/**
 * A cell
 */

var Cell = function () {
    // /**
    //  * Creates a new instance of cell.
    //  * @param {Row} row - The parent row.
    //  * @param {{}} node - The cell node.
    //  */
    function Cell(row, node, styleId) {
        _classCallCheck(this, Cell);

        this._row = row;
        this._init(node, styleId);
    }

    /* PUBLIC */

    /**
     * Gets a value indicating whether the cell is the active cell in the sheet.
     * @returns {boolean} True if active, false otherwise.
     */ /**
        * Make the cell the active cell in the sheet.
        * @param {boolean} active - Must be set to `true`. Deactivating directly is not supported. To deactivate, you should activate a different cell instead.
        * @returns {Cell} The cell.
        */


    _createClass(Cell, [{
        key: "active",
        value: function active() {
            var _this = this;

            return new ArgHandler('Cell.active').case(function () {
                return _this.sheet().activeCell() === _this;
            }).case('boolean', function (active) {
                if (!active) throw new Error("Deactivating cell directly not supported. Activate a different cell instead.");
                _this.sheet().activeCell(_this);
                return _this;
            }).handle(arguments);
        }

        /**
         * Get the address of the column.
         * @param {{}} [opts] - Options
         * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
         * @param {boolean} [opts.rowAnchored] - Anchor the row.
         * @param {boolean} [opts.columnAnchored] - Anchor the column.
         * @param {boolean} [opts.anchored] - Anchor both the row and the column.
         * @returns {string} The address
         */

    }, {
        key: "address",
        value: function address(opts) {
            return addressConverter.toAddress({
                type: 'cell',
                rowNumber: this.rowNumber(),
                columnNumber: this.columnNumber(),
                sheetName: opts && opts.includeSheetName && this.sheet().name(),
                rowAnchored: opts && (opts.rowAnchored || opts.anchored),
                columnAnchored: opts && (opts.columnAnchored || opts.anchored)
            });
        }

        /**
         * Gets the parent column of the cell.
         * @returns {Column} The parent column.
         */

    }, {
        key: "column",
        value: function column() {
            return this.sheet().column(this.columnNumber());
        }

        /**
         * Clears the contents from the cell.
         * @returns {Cell} The cell.
         */

    }, {
        key: "clear",
        value: function clear() {
            var hostSharedFormulaId = this._formulaRef && this._sharedFormulaId;

            delete this._value;
            delete this._formulaType;
            delete this._formula;
            delete this._sharedFormulaId;
            delete this._formulaRef;

            // TODO in future version: Move shared formula to some other cell. This would require parsing the formula...
            if (!_.isNil(hostSharedFormulaId)) this.sheet().clearCellsUsingSharedFormula(hostSharedFormulaId);

            return this;
        }

        /**
         * Gets the column name of the cell.
         * @returns {number} The column name.
         */

    }, {
        key: "columnName",
        value: function columnName() {
            return addressConverter.columnNumberToName(this.columnNumber());
        }

        /**
         * Gets the column number of the cell (1-based).
         * @returns {number} The column number.
         */

    }, {
        key: "columnNumber",
        value: function columnNumber() {
            return this._columnNumber;
        }

        /**
         * Find the given pattern in the cell and optionally replace it.
         * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
         * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in the cell will be replaced.
         * @returns {boolean} A flag indicating if the pattern was found.
         */

    }, {
        key: "find",
        value: function find(pattern, replacement) {
            pattern = regexify(pattern);

            var value = this.value();
            if (typeof value !== 'string') return false;

            if (_.isNil(replacement)) {
                return pattern.test(value);
            } else {
                var replaced = value.replace(pattern, replacement);
                if (replaced === value) return false;
                this.value(replaced);
                return true;
            }
        }

        /**
         * Gets the formula in the cell. Note that if a formula was set as part of a range, the getter will return 'SHARED'. This is a limitation that may be addressed in a future release.
         * @returns {string} The formula in the cell.
         */ /**
            * Sets the formula in the cell.
            * @param {string} formula - The formula to set.
            * @returns {Cell} The cell.
            */

    }, {
        key: "formula",
        value: function formula() {
            var _this2 = this;

            return new ArgHandler('Cell.formula').case(function () {
                // TODO in future: Return translated formula.
                if (_this2._formulaType === "shared" && !_this2._formulaRef) return "SHARED";
                return _this2._formula;
            }).case('nil', function () {
                _this2.clear();
                return _this2;
            }).case('string', function (formula) {
                _this2.clear();
                _this2._formulaType = "normal";
                _this2._formula = formula;
                return _this2;
            }).handle(arguments);
        }

        /**
         * Gets the hyperlink attached to the cell.
         * @returns {string|undefined} The hyperlink or undefined if not set.
         */ /**
            * Set or clear the hyperlink on the cell.
            * @param {string|undefined} hyperlink - The hyperlink to set or undefined to clear.
            * @returns {Cell} The cell.
            */

    }, {
        key: "hyperlink",
        value: function hyperlink() {
            var _this3 = this;

            return new ArgHandler('Cell.hyperlink').case(function () {
                return _this3.sheet().hyperlink(_this3.address());
            }).case('*', function (hyperlink) {
                _this3.sheet().hyperlink(_this3.address(), hyperlink);
                return _this3;
            }).handle(arguments);
        }

        /**
         * Gets the data validation object attached to the cell.
         * @returns {object|undefined} The data validation or undefined if not set.
         */ /**
            * Set or clear the data validation object of the cell.
            * @param {object|undefined} dataValidation - Object or null to clear.
            * @returns {Cell} The cell.
            */

    }, {
        key: "dataValidation",
        value: function dataValidation() {
            var _this4 = this;

            return new ArgHandler('Cell.dataValidation').case(function () {
                return _this4.sheet().dataValidation(_this4.address());
            }).case('boolean', function (obj) {
                return _this4.sheet().dataValidation(_this4.address(), obj);
            }).case('*', function (obj) {
                _this4.sheet().dataValidation(_this4.address(), obj);
                return _this4;
            }).handle(arguments);
        }

        /**
         * Callback used by tap.
         * @callback Cell~tapCallback
         * @param {Cell} cell - The cell
         * @returns {undefined}
         */
        /**
         * Invoke a callback on the cell and return the cell. Useful for method chaining.
         * @param {Cell~tapCallback} callback - The callback function.
         * @returns {Cell} The cell.
         */

    }, {
        key: "tap",
        value: function tap(callback) {
            callback(this);
            return this;
        }

        /**
         * Callback used by thru.
         * @callback Cell~thruCallback
         * @param {Cell} cell - The cell
         * @returns {*} The value to return from thru.
         */
        /**
         * Invoke a callback on the cell and return the value provided by the callback. Useful for method chaining.
         * @param {Cell~thruCallback} callback - The callback function.
         * @returns {*} The return value of the callback.
         */

    }, {
        key: "thru",
        value: function thru(callback) {
            return callback(this);
        }

        /**
         * Create a range from this cell and another.
         * @param {Cell|string} cell - The other cell or cell address to range to.
         * @returns {Range} The range.
         */

    }, {
        key: "rangeTo",
        value: function rangeTo(cell) {
            return this.sheet().range(this, cell);
        }

        /**
         * Returns a cell with a relative position given the offsets provided.
         * @param {number} rowOffset - The row offset (0 for the current row).
         * @param {number} columnOffset - The column offset (0 for the current column).
         * @returns {Cell} The relative cell.
         */

    }, {
        key: "relativeCell",
        value: function relativeCell(rowOffset, columnOffset) {
            var row = rowOffset + this.rowNumber();
            var column = columnOffset + this.columnNumber();
            return this.sheet().cell(row, column);
        }

        /**
         * Gets the parent row of the cell.
         * @returns {Row} The parent row.
         */

    }, {
        key: "row",
        value: function row() {
            return this._row;
        }

        /**
         * Gets the row number of the cell (1-based).
         * @returns {number} The row number.
         */

    }, {
        key: "rowNumber",
        value: function rowNumber() {
            return this.row().rowNumber();
        }

        /**
         * Gets the parent sheet.
         * @returns {Sheet} The parent sheet.
         */

    }, {
        key: "sheet",
        value: function sheet() {
            return this.row().sheet();
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
                  * Sets the styles in the range starting with the cell.
                  * @param {string} name - The name of the style.
                  * @param {Array.<Array.<*>>} - 2D array of values to set.
                  * @returns {Range} The range that was set.
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
            var _this5 = this;

            if (!this._style && !(arguments[0] instanceof Style)) {
                this._style = this.workbook().styleSheet().createStyle(this._styleId);
            }

            return new ArgHandler("Cell.style").case('string', function (name) {
                // Get single value
                return _this5._style.style(name);
            }).case('array', function (names) {
                // Get list of values
                var values = {};
                names.forEach(function (name) {
                    values[name] = _this5.style(name);
                });

                return values;
            }).case(["string", "array"], function (name, values) {
                var numRows = values.length;
                var numCols = values[0].length;
                var range = _this5.rangeTo(_this5.relativeCell(numRows - 1, numCols - 1));
                return range.style(name, values);
            }).case(['string', '*'], function (name, value) {
                // Set a single value for all cells to a single value
                _this5._style.style(name, value);
                return _this5;
            }).case('object', function (nameValues) {
                // Object of key value pairs to set
                for (var name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    var value = nameValues[name];
                    _this5.style(name, value);
                }

                return _this5;
            }).case('Style', function (style) {
                _this5._style = style;
                _this5._styleId = style.id();

                return _this5;
            }).handle(arguments);
        }

        /**
         * Gets the value of the cell.
         * @returns {string|boolean|number|Date|undefined} The value of the cell.
         */ /**
            * Sets the value of the cell.
            * @param {string|boolean|number|null|undefined} value - The value to set.
            * @returns {Cell} The cell.
            */ /**
               * Sets the values in the range starting with the cell.
               * @param {Array.<Array.<string|boolean|number|null|undefined>>} - 2D array of values to set.
               * @returns {Range} The range that was set.
               */

    }, {
        key: "value",
        value: function value() {
            var _this6 = this;

            return new ArgHandler('Cell.value').case(function () {
                return _this6._value;
            }).case("array", function (values) {
                var numRows = values.length;
                var numCols = values[0].length;
                var range = _this6.rangeTo(_this6.relativeCell(numRows - 1, numCols - 1));
                return range.value(values);
            }).case('*', function (value) {
                _this6.clear();
                _this6._value = value;
                return _this6;
            }).handle(arguments);
        }

        /**
         * Gets the parent workbook.
         * @returns {Workbook} The parent workbook.
         */

    }, {
        key: "workbook",
        value: function workbook() {
            return this.row().workbook();
        }

        /* INTERNAL */

        /**
         * Gets the formula if a shared formula ref cell.
         * @returns {string|undefined} The formula.
         * @ignore
         */

    }, {
        key: "getSharedRefFormula",
        value: function getSharedRefFormula() {
            return this._formulaType === "shared" ? this._formulaRef && this._formula : undefined;
        }

        /**
         * Check if this cell uses a given shared a formula ID.
         * @param {number} id - The shared formula ID.
         * @returns {boolean} A flag indicating if shared.
         * @ignore
         */

    }, {
        key: "sharesFormula",
        value: function sharesFormula(id) {
            return this._formulaType === "shared" && this._sharedFormulaId === id;
        }

        /**
         * Set a shared formula on the cell.
         * @param {number} id - The shared formula index.
         * @param {string} [formula] - The formula (if the reference cell).
         * @param {string} [sharedRef] - The address of the shared range (if the reference cell).
         * @returns {undefined}
         * @ignore
         */

    }, {
        key: "setSharedFormula",
        value: function setSharedFormula(id, formula, sharedRef) {
            this.clear();

            this._formulaType = "shared";
            this._sharedFormulaId = id;
            this._formula = formula;
            this._formulaRef = sharedRef;
        }

        /**
         * Convert the cell to an XML object.
         * @returns {{}} The XML form.
         * @ignore
         */

    }, {
        key: "toXml",
        value: function toXml() {
            // Create a node.
            var node = {
                name: 'c',
                attributes: this._remainingAttributes || {}, // Start with any remaining attributes we don't current handle.
                children: []
            };

            // Set the address.
            node.attributes.r = this.address();

            if (!_.isNil(this._formulaType)) {
                // Add the formula.
                var fNode = {
                    name: 'f',
                    attributes: this._remainingFormulaAttributes || {}
                };

                if (this._formulaType !== "normal") fNode.attributes.t = this._formulaType;
                if (!_.isNil(this._formulaRef)) fNode.attributes.ref = this._formulaRef;
                if (!_.isNil(this._sharedFormulaId)) fNode.attributes.si = this._sharedFormulaId;
                if (!_.isNil(this._formula)) fNode.children = [this._formula];

                node.children.push(fNode);
            } else if (!_.isNil(this._value)) {
                // Add the value. Don't emit value if a formula is set as Excel will show this stale value.
                var type = void 0,
                    text = void 0;
                if (typeof this._value === "string" || _.isArray(this._value)) {
                    // TODO: Rich text is array for now
                    type = "s";
                    text = this.workbook().sharedStrings().getIndexForString(this._value);
                } else if (typeof this._value === "boolean") {
                    type = "b";
                    text = this._value ? 1 : 0;
                } else if (typeof this._value === "number") {
                    text = this._value;
                } else if (this._value instanceof Date) {
                    text = dateConverter.dateToNumber(this._value);
                }

                if (type) node.attributes.t = type;
                var vNode = { name: 'v', children: [text] };
                node.children.push(vNode);
            }

            // If the style is set, set the style ID.
            if (!_.isNil(this._style)) {
                node.attributes.s = this._style.id();
            } else if (!_.isNil(this._styleId)) {
                node.attributes.s = this._styleId;
            }

            // Add any remaining children that we don't currently handle.
            if (this._remainingChildren) {
                node.children = node.children.concat(this._remainingChildren);
            }

            return node;
        }

        /* PRIVATE */

        /**
         * Initialize the cell node.
         * @param {{}|number} nodeOrColumnNumber - The existing node or the column number of a new cell.
         * @param {number} [styleId] - The style ID for the new cell.
         * @returns {undefined}
         * @private
         */

    }, {
        key: "_init",
        value: function _init(nodeOrColumnNumber, styleId) {
            if (_.isObject(nodeOrColumnNumber)) {
                // Parse the existing node.
                this._parseNode(nodeOrColumnNumber);
            } else {
                // This is a new cell.
                this._columnNumber = nodeOrColumnNumber;
                if (!_.isNil(styleId)) this._styleId = styleId;
            }
        }

        /**
         * Parse the existing node.
         * @param {{}} node - The existing node.
         * @returns {undefined}
         * @private
         */

    }, {
        key: "_parseNode",
        value: function _parseNode(node) {
            // Parse the column numbr out of the address.
            var ref = addressConverter.fromAddress(node.attributes.r);
            this._columnNumber = ref.columnNumber;

            // Store the style ID if present.
            if (!_.isNil(node.attributes.s)) this._styleId = node.attributes.s;

            // Parse the formula if present..
            var fNode = xmlq.findChild(node, 'f');
            if (fNode) {
                this._formulaType = fNode.attributes.t || "normal";
                this._formulaRef = fNode.attributes.ref;
                this._formula = fNode.children[0];

                this._sharedFormulaId = fNode.attributes.si;
                if (!_.isNil(this._sharedFormulaId)) {
                    // Update the sheet's max shared formula ID so we can set future IDs an index beyond this.
                    this.sheet().updateMaxSharedFormulaId(this._sharedFormulaId);
                }

                // Delete the known attributes.
                delete fNode.attributes.t;
                delete fNode.attributes.ref;
                delete fNode.attributes.si;

                // If any unknown attributes are still present, store them for later output.
                if (!_.isEmpty(fNode.attributes)) this._remainingFormulaAttributes = fNode.attributes;
            }

            // Parse the value.
            var type = node.attributes.t;
            if (type === "s") {
                // String value.
                var sharedIndex = xmlq.findChild(node, 'v').children[0];
                this._value = this.workbook().sharedStrings().getStringByIndex(sharedIndex);
            } else if (type === "inlineStr") {
                // Inline string value: can be simple text or rich text.
                var isNode = xmlq.findChild(node, 'is');
                if (isNode.children[0].name === "t") {
                    var tNode = isNode.children[0];
                    this._value = tNode.children[0];
                } else {
                    this._value = isNode.children;
                }
            } else if (type === "b") {
                // Boolean value.
                this._value = xmlq.findChild(node, 'v').children[0] === 1;
            } else if (type === "e") {
                // Error value.
                var error = xmlq.findChild(node, 'v').children[0];
                this._value = FormulaError.getError(error);
            } else {
                // Number value.
                var vNode = xmlq.findChild(node, 'v');
                this._value = vNode && Number(vNode.children[0]);
            }

            // Delete known attributes.
            delete node.attributes.r;
            delete node.attributes.s;
            delete node.attributes.t;

            // If any unknown attributes are still present, store them for later output.
            if (!_.isEmpty(node.attributes)) this._remainingAttributes = node.attributes;

            // Delete known children.
            xmlq.removeChild(node, 'f');
            xmlq.removeChild(node, 'v');
            xmlq.removeChild(node, 'is');

            // If any unknown children are still present, store them for later output.
            if (!_.isEmpty(node.children)) this._remainingChildren = node.children;
        }
    }]);

    return Cell;
}();

module.exports = Cell;

/*
<c r="A6" s="1" t="s">
    <v>2</v>
</c>
*/