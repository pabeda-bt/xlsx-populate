"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");

/**
 * The shared strings table.
 * @ignore
 */

var SharedStrings = function () {
    /**
     * Constructs a new instance of _SharedStrings.
     * @param {{}} node - The node.
     */
    function SharedStrings(node) {
        _classCallCheck(this, SharedStrings);

        this._stringArray = [];
        this._indexMap = {};

        this._init(node);
        this._cacheExistingSharedStrings();
    }

    /**
     * Gets the index for a string
     * @param {string|Array.<{}>} string - The string or rich text array.
     * @returns {number} The index
     */


    _createClass(SharedStrings, [{
        key: "getIndexForString",
        value: function getIndexForString(string) {
            // If the string is found in the cache, return the index.
            var key = _.isArray(string) ? JSON.stringify(string) : string;
            var index = this._indexMap[key];
            if (index >= 0) return index;

            // Otherwise, add it to the caches.
            index = this._stringArray.length;
            this._stringArray.push(string);
            this._indexMap[key] = index;

            // Append a new si node.
            this._node.children.push({
                name: "si",
                children: _.isArray(string) ? string : [{
                    name: "t",
                    attributes: { 'xml:space': "preserve" },
                    children: [string]
                }]
            });

            return index;
        }

        /**
         * Get the string for a given index
         * @param {number} index - The index
         * @returns {string} The string
         */

    }, {
        key: "getStringByIndex",
        value: function getStringByIndex(index) {
            return this._stringArray[index];
        }

        /**
         * Convert the collection to an XML object.
         * @returns {{}} The XML object.
         */

    }, {
        key: "toXml",
        value: function toXml() {
            return this._node;
        }

        /**
         * Store any existing values in the caches.
         * @private
         * @returns {undefined}
         */

    }, {
        key: "_cacheExistingSharedStrings",
        value: function _cacheExistingSharedStrings() {
            var _this = this;

            this._node.children.forEach(function (node, i) {
                var content = node.children[0];
                if (content.name === "t") {
                    var string = content.children[0];
                    _this._stringArray.push(string);
                    _this._indexMap[string] = i;
                } else {
                    // TODO: Properly support rich text nodes in the future. For now just store the object as a placeholder.
                    _this._stringArray.push(node.children);
                    _this._indexMap[JSON.stringify(node.children)] = i;
                }
            });
        }

        /**
         * Initialize the node.
         * @param {{}} [node] - The shared strings node.
         * @private
         * @returns {undefined}
         */

    }, {
        key: "_init",
        value: function _init(node) {
            if (!node) node = {
                name: "sst",
                attributes: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                },
                children: []
            };

            this._node = node;

            delete this._node.attributes.count;
            delete this._node.attributes.uniqueCount;
        }
    }]);

    return SharedStrings;
}();

module.exports = SharedStrings;

/*
xl/sharedStrings.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="13" uniqueCount="4">
	<si>
		<t>Foo</t>
	</si>
	<si>
		<t>Bar</t>
	</si>
	<si>
		<t>Goo</t>
	</si>
	<si>
		<r>
			<t>s</t>
		</r><r>
			<rPr>
				<b/>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr><t>d;</t>
		</r><r>
			<rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr><t>lfk;l</t>
		</r>
	</si>
</sst>
*/