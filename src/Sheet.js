"use strict";

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");
var Cell = require("./Cell");
var Row = require("./Row");
var Column = require("./Column");
var Range = require("./Range");
var Relationships = require("./Relationships");
var xmlq = require("./xmlq");
var regexify = require("./regexify");
var addressConverter = require("./addressConverter");
var ArgHandler = require("./ArgHandler");
var colorIndexes = require("./colorIndexes");

// Order of the nodes as defined by the spec.
var nodeOrder = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];

/**
 * A worksheet.
 */

var Sheet = function () {
    // /**
    //  * Creates a new instance of Sheet.
    //  * @param {Workbook} workbook - The parent workbook.
    //  * @param {{}} idNode - The sheet ID node (from the parent workbook).
    //  * @param {{}} node - The sheet node.
    //  * @param {{}} [relationshipsNode] - The optional sheet relationships node.
    //  */
    function Sheet(workbook, idNode, node, relationshipsNode) {
        _classCallCheck(this, Sheet);

        this._init(workbook, idNode, node, relationshipsNode);
    }

    /* PUBLIC */

    /**
     * Gets a value indicating whether the sheet is the active sheet in the workbook.
     * @returns {boolean} True if active, false otherwise.
     */ /**
        * Make the sheet the active sheet in the workkbok.
        * @param {boolean} active - Must be set to `true`. Deactivating directly is not supported. To deactivate, you should activate a different sheet instead.
        * @returns {Sheet} The sheet.
        */


    _createClass(Sheet, [{
        key: "active",
        value: function active() {
            var _this = this;

            return new ArgHandler('Sheet.active').case(function () {
                return _this.workbook().activeSheet() === _this;
            }).case('boolean', function (active) {
                if (!active) throw new Error("Deactivating sheet directly not supported. Activate a different sheet instead.");
                _this.workbook().activeSheet(_this);
                return _this;
            }).handle(arguments);
        }

        /**
         * Get the active cell in the sheet.
         * @returns {Cell} The active cell.
         */ /**
            * Set the active cell in the workbook.
            * @param {string|Cell} cell - The cell or address of cell to activate.
            * @returns {Sheet} The sheet.
            */ /**
               * Set the active cell in the workbook by row and column.
               * @param {number} rowNumber - The row number of the cell.
               * @param {string|number} columnNameOrNumber - The column name or number of the cell.
               * @returns {Sheet} The sheet.
               */

    }, {
        key: "activeCell",
        value: function activeCell() {
            var _this2 = this;

            var sheetViewNode = this._getOrCreateSheetViewNode();
            var selectionNode = xmlq.findChild(sheetViewNode, "selection");
            return new ArgHandler('Sheet.activeCell').case(function () {
                var cellAddress = selectionNode ? selectionNode.attributes.activeCell : "A1";
                return _this2.cell(cellAddress);
            }).case(['number', '*'], function (rowNumber, columnNameOrNumber) {
                var cell = _this2.cell(rowNumber, columnNameOrNumber);
                return _this2.activeCell(cell);
            }).case('*', function (cell) {
                if (!selectionNode) {
                    selectionNode = {
                        name: "selection",
                        attributes: {},
                        children: []
                    };

                    xmlq.appendChild(sheetViewNode, selectionNode);
                }

                if (!(cell instanceof Cell)) cell = _this2.cell(cell);
                selectionNode.attributes.activeCell = selectionNode.attributes.sqref = cell.address();
                return _this2;
            }).handle(arguments);
        }

        /**
         * Gets the cell with the given address.
         * @param {string} address - The address of the cell.
         * @returns {Cell} The cell.
         */ /**
            * Gets the cell with the given row and column numbers.
            * @param {number} rowNumber - The row number of the cell.
            * @param {string|number} columnNameOrNumber - The column name or number of the cell.
            * @returns {Cell} The cell.
            */

    }, {
        key: "cell",
        value: function cell() {
            var _this3 = this;

            return new ArgHandler('Sheet.cell').case('string', function (address) {
                var ref = addressConverter.fromAddress(address);
                if (ref.type !== 'cell') throw new Error('Sheet.cell: Invalid address.');
                return _this3.row(ref.rowNumber).cell(ref.columnNumber);
            }).case(['number', '*'], function (rowNumber, columnNameOrNumber) {
                return _this3.row(rowNumber).cell(columnNameOrNumber);
            }).handle(arguments);
        }

        /**
         * Gets a column in the sheet.
         * @param {string|number} columnNameOrNumber - The name or number of the column.
         * @returns {Column} The column.
         */

    }, {
        key: "column",
        value: function column(columnNameOrNumber) {
            var columnNumber = typeof columnNameOrNumber === "string" ? addressConverter.columnNameToNumber(columnNameOrNumber) : columnNameOrNumber;

            // If we're already created a column for this column number, return it.
            if (this._columns[columnNumber]) return this._columns[columnNumber];

            // We need to create a new column, which requires a backing col node. There may already exist a node whose min/max cover our column.
            // First, see if there is an existing col node.
            var existingColNode = this._colNodes[columnNumber];

            var colNode = void 0;
            if (existingColNode) {
                // If the existing node covered earlier columns than the new one, we need to have a col node to cover the min up to our new node.
                if (existingColNode.attributes.min < columnNumber) {
                    // Clone the node and set the max to the column before our new col.
                    var beforeColNode = _.cloneDeep(existingColNode);
                    beforeColNode.attributes.max = columnNumber - 1;

                    // Update the col nodes cache.
                    for (var i = beforeColNode.attributes.min; i <= beforeColNode.attributes.max; i++) {
                        this._colNodes[i] = beforeColNode;
                    }
                }

                // Make a clone for the new column. Set the min/max to the column number and cache it.
                colNode = _.cloneDeep(existingColNode);
                colNode.attributes.min = columnNumber;
                colNode.attributes.max = columnNumber;
                this._colNodes[columnNumber] = colNode;

                // If the max of the existing node is greater than the nre one, create a col node for that too.
                if (existingColNode.attributes.max > columnNumber) {
                    var afterColNode = _.cloneDeep(existingColNode);
                    afterColNode.attributes.min = columnNumber + 1;
                    for (var _i = afterColNode.attributes.min; _i <= afterColNode.attributes.max; _i++) {
                        this._colNodes[_i] = afterColNode;
                    }
                }
            } else {
                // The was no existing node so create a new one.
                colNode = {
                    name: 'col',
                    attributes: {
                        min: columnNumber,
                        max: columnNumber
                    },
                    children: []
                };

                this._colNodes[columnNumber] = colNode;
            }

            // Create the new column and cache it.
            var column = new Column(this, colNode);
            this._columns[columnNumber] = column;
            return column;
        }

        /**
         * Gets a defined name scoped to the sheet.
         * @param {string} name - The defined name.
         * @returns {undefined|string|Cell|Range|Row|Column} What the defined name refers to or undefined if not found. Will return the string formula if not a Row, Column, Cell, or Range.
         */ /**
            * Set a defined name scoped to the sheet.
            * @param {string} name - The defined name.
            * @param {string|Cell|Range|Row|Column} refersTo - What the name refers to.
            * @returns {Workbook} The workbook.
            */

    }, {
        key: "definedName",
        value: function definedName() {
            var _this4 = this;

            return new ArgHandler("Workbook.definedName").case('string', function (name) {
                return _this4.workbook().scopedDefinedName(_this4, name);
            }).case(['string', '*'], function (name, refersTo) {
                _this4.workbook().scopedDefinedName(_this4, name, refersTo);
                return _this4;
            }).handle(arguments);
        }

        /**
         * Deletes the sheet and returns the parent workbook.
         * @returns {Workbook} The workbook.
         */

    }, {
        key: "delete",
        value: function _delete() {
            this.workbook().deleteSheet(this);
            return this.workbook();
        }

        /**
         * Find the given pattern in the sheet and optionally replace it.
         * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
         * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced.
         * @returns {Array.<Cell>} The matching cells.
         */

    }, {
        key: "find",
        value: function find(pattern, replacement) {
            pattern = regexify(pattern);

            var matches = [];
            this._rows.forEach(function (row) {
                if (!row) return;
                matches = matches.concat(row.find(pattern, replacement));
            });

            return matches;
        }

        /**
         * Gets a value indicating whether this sheet's grid lines are visible.
         * @returns {boolean} True if selected, false if not.
         */ /**
            * Sets whether this sheet's grid lines are visible.
            * @param {boolean} selected - True to make visible, false to hide.
            * @returns {Sheet} The sheet.
            */

    }, {
        key: "gridLinesVisible",
        value: function gridLinesVisible() {
            var _this5 = this;

            var sheetViewNode = this._getOrCreateSheetViewNode();
            return new ArgHandler('Sheet.gridLinesVisible').case(function () {
                return sheetViewNode.attributes.showGridLines === 1 || sheetViewNode.attributes.showGridLines === undefined;
            }).case('boolean', function (visible) {
                sheetViewNode.attributes.showGridLines = visible ? 1 : 0;
                return _this5;
            }).handle(arguments);
        }

        /**
         * Gets a value indicating if the sheet is hidden or not.
         * @returns {boolean|string} True if hidden, false if visible, and 'very' if very hidden.
         */ /**
            * Set whether the sheet is hidden or not.
            * @param {boolean|string} hidden - True to hide, false to show, and 'very' to make very hidden.
            * @returns {Sheet} The sheet.
            */

    }, {
        key: "hidden",
        value: function hidden() {
            var _this6 = this;

            return new ArgHandler('Sheet.hidden').case(function () {
                if (_this6._idNode.attributes.state === 'hidden') return true;
                if (_this6._idNode.attributes.state === 'veryHidden') return "very";
                return false;
            }).case('*', function (hidden) {
                if (hidden) {
                    var visibleSheets = _.filter(_this6.workbook().sheets(), function (sheet) {
                        return !sheet.hidden();
                    });
                    if (visibleSheets.length === 1 && visibleSheets[0] === _this6) {
                        throw new Error("This sheet may not be hidden as a workbook must contain at least one visible sheet.");
                    }

                    // If activate, activate the first other visible sheet.
                    if (_this6.active()) {
                        var activeIndex = visibleSheets[0] === _this6 ? 1 : 0;
                        visibleSheets[activeIndex].active(true);
                    }
                }

                if (hidden === 'very') _this6._idNode.attributes.state = 'veryHidden';else if (hidden) _this6._idNode.attributes.state = 'hidden';else delete _this6._idNode.attributes.state;
                return _this6;
            }).handle(arguments);
        }

        /**
         * Move the sheet.
         * @param {number|string|Sheet} [indexOrBeforeSheet] The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
         * @returns {Sheet} The sheet.
         */

    }, {
        key: "move",
        value: function move(indexOrBeforeSheet) {
            this.workbook().moveSheet(this, indexOrBeforeSheet);
            return this;
        }

        /**
         * Get the name of the sheet.
         * @returns {string} The sheet name.
         */ /**
            * Set the name of the sheet. *Note: this method does not rename references to the sheet so formulas, etc. can be broken. Use with caution!*
            * @param {string} name - The name to set to the sheet.
            * @returns {Sheet} The sheet.
            */

    }, {
        key: "name",
        value: function name() {
            var _this7 = this;

            return new ArgHandler('Sheet.name').case(function () {
                return _this7._idNode.attributes.name;
            }).case('string', function (name) {
                _this7._idNode.attributes.name = name;
                return _this7;
            }).handle(arguments);
        }

        /**
         * Gets a range from the given range address.
         * @param {string} address - The range address (e.g. 'A1:B3').
         * @returns {Range} The range.
         */ /**
            * Gets a range from the given cells or cell addresses.
            * @param {string|Cell} startCell - The starting cell or cell address (e.g. 'A1').
            * @param {string|Cell} endCell - The ending cell or cell address (e.g. 'B3').
            * @returns {Range} The range.
            */ /**
               * Gets a range from the given row numbers and column names or numbers.
               * @param {number} startRowNumber - The starting cell row number.
               * @param {string|number} startColumnNameOrNumber - The starting cell column name or number.
               * @param {number} endRowNumber - The ending cell row number.
               * @param {string|number} endColumnNameOrNumber - The ending cell column name or number.
               * @returns {Range} The range.
               */

    }, {
        key: "range",
        value: function range() {
            var _this8 = this;

            return new ArgHandler('Sheet.range').case('string', function (address) {
                var ref = addressConverter.fromAddress(address);
                if (ref.type !== 'range') throw new Error('Sheet.range: Invalid address');
                return _this8.range(ref.startRowNumber, ref.startColumnNumber, ref.endRowNumber, ref.endColumnNumber);
            }).case(['*', '*'], function (startCell, endCell) {
                if (typeof startCell === "string") startCell = _this8.cell(startCell);
                if (typeof endCell === "string") endCell = _this8.cell(endCell);
                return new Range(startCell, endCell);
            }).case(['number', '*', 'number', '*'], function (startRowNumber, startColumnNameOrNumber, endRowNumber, endColumnNameOrNumber) {
                return _this8.range(_this8.cell(startRowNumber, startColumnNameOrNumber), _this8.cell(endRowNumber, endColumnNameOrNumber));
            }).handle(arguments);
        }

        /**
         * Gets the row with the given number.
         * @param {number} rowNumber - The row number.
         * @returns {Row} The row with the given number.
         */

    }, {
        key: "row",
        value: function row(rowNumber) {
            if (this._rows[rowNumber]) return this._rows[rowNumber];

            var rowNode = {
                name: 'row',
                attributes: {
                    r: rowNumber
                },
                children: []
            };

            var row = new Row(this, rowNode);
            this._rows[rowNumber] = row;
            return row;
        }

        /**
         * Get the tab color. (See style [Color](#color).)
         * @returns {undefined|Color} The color or undefined if not set.
         */ /**
            * Sets the tab color. (See style [Color](#color).)
            * @returns {Color|string|number} color - Color of the tab. If string, will set an RGB color. If number, will set a theme color.
            */

    }, {
        key: "tabColor",
        value: function tabColor() {
            var _this9 = this;

            return new ArgHandler("Sheet.tabColor").case(function () {
                var tabColorNode = xmlq.findChild(_this9._sheetPrNode, "tabColor");
                if (!tabColorNode) return;

                var color = {};
                if (tabColorNode.attributes.hasOwnProperty('rgb')) color.rgb = tabColorNode.attributes.rgb;else if (tabColorNode.attributes.hasOwnProperty('theme')) color.theme = tabColorNode.attributes.theme;else if (tabColorNode.attributes.hasOwnProperty('indexed')) color.rgb = colorIndexes[tabColorNode.attributes.indexed];

                if (tabColorNode.attributes.hasOwnProperty('tint')) color.tint = tabColorNode.attributes.tint;

                return color;
            }).case("string", function (rgb) {
                return _this9.tabColor({ rgb: rgb });
            }).case("integer", function (theme) {
                return _this9.tabColor({ theme: theme });
            }).case("nil", function () {
                xmlq.removeChild(_this9._sheetPrNode, "tabColor");
                return _this9;
            }).case("object", function (color) {
                var tabColorNode = xmlq.appendChildIfNotFound(_this9._sheetPrNode, "tabColor");
                xmlq.setAttributes(tabColorNode, {
                    rgb: color.rgb && color.rgb.toUpperCase(),
                    indexed: null,
                    theme: color.theme,
                    tint: color.tint
                });

                return _this9;
            }).handle(arguments);
        }

        /**
         * Gets a value indicating whether this sheet is selected.
         * @returns {boolean} True if selected, false if not.
         */ /**
            * Sets whether this sheet is selected.
            * @param {boolean} selected - True to select, false to deselected.
            * @returns {Sheet} The sheet.
            */

    }, {
        key: "tabSelected",
        value: function tabSelected() {
            var _this10 = this;

            var sheetViewNode = this._getOrCreateSheetViewNode();
            return new ArgHandler('Sheet.tabSelected').case(function () {
                return sheetViewNode.attributes.tabSelected === 1;
            }).case('boolean', function (selected) {
                if (selected) sheetViewNode.attributes.tabSelected = 1;else delete sheetViewNode.attributes.tabSelected;
                return _this10;
            }).handle(arguments);
        }

        /**
         * Get the range of cells in the sheet that have contained a value or style at any point. Useful for extracting the entire sheet contents.
         * @returns {Range|undefined} The used range or undefined if no cells in the sheet are used.
         */

    }, {
        key: "usedRange",
        value: function usedRange() {
            var minRowNumber = _.findIndex(this._rows);
            var maxRowNumber = this._rows.length - 1;

            var minColumnNumber = 0;
            var maxColumnNumber = 0;
            for (var i = 0; i < this._rows.length; i++) {
                var row = this._rows[i];
                if (!row) continue;

                var minUsedColumnNumber = row.minUsedColumnNumber();
                var maxUsedColumnNumber = row.maxUsedColumnNumber();
                if (minUsedColumnNumber > 0 && (!minColumnNumber || minUsedColumnNumber < minColumnNumber)) minColumnNumber = minUsedColumnNumber;
                if (maxUsedColumnNumber > 0 && (!maxColumnNumber || maxUsedColumnNumber > maxColumnNumber)) maxColumnNumber = maxUsedColumnNumber;
            }

            // Return undefined if nothing in the sheet is used.
            if (minRowNumber <= 0 || minColumnNumber <= 0 || maxRowNumber <= 0 || maxColumnNumber <= 0) return;

            return this.range(minRowNumber, minColumnNumber, maxRowNumber, maxColumnNumber);
        }

        /**
         * Gets the parent workbook.
         * @returns {Workbook} The parent workbook.
         */

    }, {
        key: "workbook",
        value: function workbook() {
            return this._workbook;
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
            this._rows.forEach(function (row) {
                if (!row) return;
                row.clearCellsUsingSharedFormula(sharedFormulaId);
            });
        }

        /**
         * Get an existing column style ID.
         * @param {number} columnNumber - The column number.
         * @returns {undefined|number} The style ID.
         * @ignore
         */

    }, {
        key: "existingColumnStyleId",
        value: function existingColumnStyleId(columnNumber) {
            // This will work after setting Column.style because Column updates the attributes live.
            var colNode = this._colNodes[columnNumber];
            return colNode && colNode.attributes.style;
        }

        /**
         * Call a callback for each column number that has a node defined for it.
         * @param {Function} callback - The callback.
         * @returns {undefined}
         * @ignore
         */

    }, {
        key: "forEachExistingColumnNumber",
        value: function forEachExistingColumnNumber(callback) {
            _.forEach(this._colNodes, function (node, columnNumber) {
                if (!node) return;
                callback(columnNumber);
            });
        }

        /**
         * Call a callback for each existing row.
         * @param {Function} callback - The callback.
         * @returns {undefined}
         * @ignore
         */

    }, {
        key: "forEachExistingRow",
        value: function forEachExistingRow(callback) {
            _.forEach(this._rows, function (row, rowNumber) {
                if (row) callback(row, rowNumber);
            });

            return this;
        }

        /**
         * Get the hyperlink attached to the cell with the given address.
         * @param {string} address - The address of the hyperlinked cell.
         * @returns {string|undefined} The hyperlink or undefined if not set.
         * @ignore
         */ /**
            * Set the hyperlink attached to the cell with the given address.
            * @param {string} address - The address to of the hyperlinked cell.
            * @param {boolean} hyperlink - The hyperlink to set or undefined to clear.
            * @returns {Sheet} The sheet.
            * @ignore
            */

    }, {
        key: "hyperlink",
        value: function hyperlink() {
            var _this11 = this;

            return new ArgHandler('Sheet.hyperlink').case('string', function (address) {
                var hyperlinkNode = _this11._hyperlinks[address];
                if (!hyperlinkNode) return;
                var relationship = _this11._relationships.findById(hyperlinkNode.attributes['r:id']);
                return relationship && relationship.attributes.Target;
            }).case(['string', 'nil'], function (address) {
                delete _this11._hyperlinks[address];
                return _this11;
            }).case(['string', 'string'], function (address, hyperlink) {
                var relationship = _this11._relationships.add("hyperlink", hyperlink, "External");
                _this11._hyperlinks[address] = {
                    name: 'hyperlink',
                    attributes: { ref: address, 'r:id': relationship.attributes.Id },
                    children: []
                };

                return _this11;
            }).handle(arguments);
        }

        /**
         * Increment and return the max shared formula ID.
         * @returns {number} The new max shared formula ID.
         * @ignore
         */

    }, {
        key: "incrementMaxSharedFormulaId",
        value: function incrementMaxSharedFormulaId() {
            return ++this._maxSharedFormulaId;
        }

        /**
         * Get a value indicating whether the cells in the given address are merged.
         * @param {string} address - The address to check.
         * @returns {boolean} True if merged, false if not merged.
         * @ignore
         */ /**
            * Merge/unmerge cells by adding/removing a mergeCell entry.
            * @param {string} address - The address to merge.
            * @param {boolean} merged - True to merge, false to unmerge.
            * @returns {Sheet} The sheet.
            * @ignore
            */

    }, {
        key: "merged",
        value: function merged() {
            var _this12 = this;

            return new ArgHandler('Sheet.merge').case('string', function (address) {
                return _this12._mergeCells.hasOwnProperty(address);
            }).case(['string', '*'], function (address, merge) {
                if (merge) {
                    _this12._mergeCells[address] = {
                        name: 'mergeCell',
                        attributes: { ref: address },
                        children: []
                    };
                } else {
                    delete _this12._mergeCells[address];
                }

                return _this12;
            }).handle(arguments);
        }

        /**
         * Gets a Object or undefined of the cells in the given address.
         * @param {string} address - The address to check.
         * @returns {object|boolean} Object or false if not set
         * @ignore
         */ /**
            * Removes dataValidation at the given address
            * @param {string} address - The address to remove.
            * @param {boolean} obj - false to delete.
            * @returns {boolean} true if removed.
            * @ignore
            */ /**
               * Add dataValidation to cells at the given address if object or string
               * @param {string} address - The address to set.
               * @param {object|string} obj - Object or String to set
               * @returns {Sheet} The sheet.
               * @ignore
               */

    }, {
        key: "dataValidation",
        value: function dataValidation() {
            var _this13 = this;

            return new ArgHandler('Sheet.dataValidation').case('string', function (address) {
                if (_this13._dataValidations[address]) {
                    return {
                        type: _this13._dataValidations[address].attributes.type,
                        allowBlank: _this13._dataValidations[address].attributes.allowBlank,
                        showInputMessage: _this13._dataValidations[address].attributes.showInputMessage,
                        prompt: _this13._dataValidations[address].attributes.prompt,
                        promptTitle: _this13._dataValidations[address].attributes.promptTitle,
                        showErrorMessage: _this13._dataValidations[address].attributes.showErrorMessage,
                        error: _this13._dataValidations[address].attributes.error,
                        errorTitle: _this13._dataValidations[address].attributes.errorTitle,
                        operator: _this13._dataValidations[address].attributes.operator,
                        formula1: _this13._dataValidations[address].children[0].children[0],
                        formula2: _this13._dataValidations[address].children[1] ? _this13._dataValidations[address].children[1].children[0] : undefined
                    };
                } else {
                    return false;
                }
            }).case(['string', 'boolean'], function (address, obj) {
                if (_this13._dataValidations[address]) {
                    if (obj === false) return delete _this13._dataValidations[address];
                } else {
                    return false;
                }
            }).case(['string', '*'], function (address, obj) {
                if (typeof obj === 'string') {
                    _this13._dataValidations[address] = {
                        name: 'dataValidation',
                        attributes: {
                            type: 'list',
                            allowBlank: false,
                            showInputMessage: false,
                            prompt: '',
                            promptTitle: '',
                            showErrorMessage: false,
                            error: '',
                            errorTitle: '',
                            operator: '',
                            sqref: address
                        },
                        children: [{
                            name: 'formula1',
                            atrributes: {},
                            children: [obj]
                        }, {
                            name: 'formula2',
                            atrributes: {},
                            children: ['']
                        }]
                    };
                } else if ((typeof obj === "undefined" ? "undefined" : _typeof(obj)) === 'object') {
                    _this13._dataValidations[address] = {
                        name: 'dataValidation',
                        attributes: {
                            type: obj.type ? obj.type : 'list',
                            allowBlank: obj.allowBlank,
                            showInputMessage: obj.showInputMessage,
                            prompt: obj.prompt,
                            promptTitle: obj.promptTitle,
                            showErrorMessage: obj.showErrorMessage,
                            error: obj.error,
                            errorTitle: obj.errorTitle,
                            operator: obj.operator,
                            sqref: address
                        },
                        children: [{
                            name: 'formula1',
                            atrributes: {},
                            children: [obj.formula1]
                        }, {
                            name: 'formula2',
                            atrributes: {},
                            children: [obj.formula2]
                        }]
                    };
                }
                return _this13;
            }).handle(arguments);
        }

        /**
         * Convert the sheet to a collection of XML objects.
         * @returns {{}} The XML forms.
         * @ignore
         */

    }, {
        key: "toXmls",
        value: function toXmls() {
            // Shallow clone the node so we don't have to remove these children later if they don't belong.
            var node = _.clone(this._node);
            node.children = node.children.slice();

            // Add the columns if needed.
            this._colsNode.children = _.filter(this._colNodes, function (colNode, i) {
                // Columns should only be present if they have attributes other than min/max.
                return colNode && i === colNode.attributes.min && Object.keys(colNode.attributes).length > 2;
            });
            if (this._colsNode.children.length) {
                xmlq.insertInOrder(node, this._colsNode, nodeOrder);
            }

            // Add the hyperlinks if needed.
            this._hyperlinksNode.children = _.values(this._hyperlinks);
            if (this._hyperlinksNode.children.length) {
                xmlq.insertInOrder(node, this._hyperlinksNode, nodeOrder);
            }

            // Add the merge cells if needed.
            this._mergeCellsNode.children = _.values(this._mergeCells);
            if (this._mergeCellsNode.children.length) {
                xmlq.insertInOrder(node, this._mergeCellsNode, nodeOrder);
            }

            // Add the DataValidation cells if needed.
            this._dataValidationsNode.children = _.values(this._dataValidations);
            if (this._dataValidationsNode.children.length) {
                xmlq.insertInOrder(node, this._dataValidationsNode, nodeOrder);
            }

            return {
                id: this._idNode,
                sheet: node,
                relationships: this._relationships
            };
        }

        /**
         * Update the max shared formula ID to the given value if greater than current.
         * @param {number} sharedFormulaId - The new shared formula ID.
         * @returns {undefined}
         * @ignore
         */

    }, {
        key: "updateMaxSharedFormulaId",
        value: function updateMaxSharedFormulaId(sharedFormulaId) {
            if (sharedFormulaId > this._maxSharedFormulaId) {
                this._maxSharedFormulaId = sharedFormulaId;
            }
        }

        /* PRIVATE */

        /**
         * Get the sheet view node if it exists or create it if it doesn't.
         * @returns {{}} The sheet view node.
         * @private
         */

    }, {
        key: "_getOrCreateSheetViewNode",
        value: function _getOrCreateSheetViewNode() {
            var sheetViewsNode = xmlq.findChild(this._node, "sheetViews");
            if (!sheetViewsNode) {
                sheetViewsNode = {
                    name: "sheetViews",
                    attributes: {},
                    children: [{
                        name: "sheetView",
                        attributes: {
                            workbookViewId: 0
                        },
                        children: []
                    }]
                };

                xmlq.insertInOrder(this._node, sheetViewsNode, nodeOrder);
            }

            return xmlq.findChild(sheetViewsNode, "sheetView");
        }

        /**
         * Initializes the sheet.
         * @param {Workbook} workbook - The parent workbook.
         * @param {{}} idNode - The sheet ID node (from the parent workbook).
         * @param {{}} node - The sheet node.
         * @param {{}} [relationshipsNode] - The optional sheet relationships node.
         * @returns {undefined}
         * @private
         */

    }, {
        key: "_init",
        value: function _init(workbook, idNode, node, relationshipsNode) {
            var _this14 = this;

            if (!node) {
                node = {
                    name: "worksheet",
                    attributes: {
                        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                        'xmlns:r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        'xmlns:mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
                        'mc:Ignorable': "x14ac",
                        'xmlns:x14ac': "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                    },
                    children: [{
                        name: "sheetData",
                        attributes: {},
                        children: []
                    }]
                };
            }

            this._workbook = workbook;
            this._idNode = idNode;
            this._node = node;
            this._maxSharedFormulaId = -1;
            this._mergeCells = {};
            this._dataValidations = {};
            this._hyperlinks = {};

            // Create the relationships.
            this._relationships = new Relationships(relationshipsNode);

            // Delete the optional dimension node
            xmlq.removeChild(this._node, "dimension");

            // Create the rows.
            this._rows = [];
            this._sheetDataNode = xmlq.findChild(this._node, "sheetData");
            this._sheetDataNode.children.forEach(function (rowNode) {
                var row = new Row(_this14, rowNode);
                _this14._rows[row.rowNumber()] = row;
            });
            this._sheetDataNode.children = this._rows;

            // Create the columns node.
            this._columns = [];
            this._colsNode = xmlq.findChild(this._node, "cols");
            if (this._colsNode) {
                xmlq.removeChild(this._node, this._colsNode);
            } else {
                this._colsNode = { name: 'cols', attributes: {}, children: [] };
            }

            // Cache the col nodes.
            this._colNodes = [];
            _.forEach(this._colsNode.children, function (colNode) {
                var min = colNode.attributes.min;
                var max = colNode.attributes.max;
                for (var i = min; i <= max; i++) {
                    _this14._colNodes[i] = colNode;
                }
            });

            // Create the sheet properties node.
            this._sheetPrNode = xmlq.findChild(this._node, "sheetPr");
            if (!this._sheetPrNode) {
                this._sheetPrNode = { name: 'sheetPr', attributes: {}, children: [] };
                xmlq.insertInOrder(this._node, this._sheetPrNode, nodeOrder);
            }

            // Create the merge cells.
            this._mergeCellsNode = xmlq.findChild(this._node, "mergeCells");
            if (this._mergeCellsNode) {
                xmlq.removeChild(this._node, this._mergeCellsNode);
            } else {
                this._mergeCellsNode = { name: 'mergeCells', attributes: {}, children: [] };
            }

            var mergeCellNodes = this._mergeCellsNode.children;
            this._mergeCellsNode.children = [];
            mergeCellNodes.forEach(function (mergeCellNode) {
                _this14._mergeCells[mergeCellNode.attributes.ref] = mergeCellNode;
            });

            // Create the DataValidations.
            this._dataValidationsNode = xmlq.findChild(this._node, "dataValidations");
            if (this._dataValidationsNode) {
                xmlq.removeChild(this._node, this._dataValidationsNode);
            } else {
                this._dataValidationsNode = { name: 'dataValidations', attributes: {}, children: [] };
            }

            var dataValidationNodes = this._dataValidationsNode.children;
            this._dataValidationsNode.children = [];
            dataValidationNodes.forEach(function (dataValidationNode) {
                _this14._dataValidations[dataValidationNode.attributes.sqref] = dataValidationNode;
            });

            // Create the hyperlinks.
            this._hyperlinksNode = xmlq.findChild(this._node, "hyperlinks");
            if (this._hyperlinksNode) {
                xmlq.removeChild(this._node, this._hyperlinksNode);
            } else {
                this._hyperlinksNode = { name: 'hyperlinks', attributes: {}, children: [] };
            }

            var hyperlinkNodes = this._hyperlinksNode.children;
            this._hyperlinksNode.children = [];
            hyperlinkNodes.forEach(function (hyperlinkNode) {
                _this14._hyperlinks[hyperlinkNode.attributes.ref] = hyperlinkNode;
            });
        }
    }]);

    return Sheet;
}();

module.exports = Sheet;