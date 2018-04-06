"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");

/**
 * A content type collection.
 * @ignore
 */

var ContentTypes = function () {
    /**
     * Creates a new instance of ContentTypes
     * @param {{}} node - The node.
     */
    function ContentTypes(node) {
        _classCallCheck(this, ContentTypes);

        this._node = node;
    }

    /**
     * Add a new content type.
     * @param {string} partName - The part name.
     * @param {string} contentType - The content type.
     * @returns {{}} The new content type.
     */


    _createClass(ContentTypes, [{
        key: "add",
        value: function add(partName, contentType) {
            var node = {
                name: "Override",
                attributes: {
                    PartName: partName,
                    ContentType: contentType
                }
            };

            this._node.children.push(node);
            return node;
        }

        /**
         * Find a content type by part name.
         * @param {string} partName - The part name.
         * @returns {{}|undefined} The matching content type or undefined if not found.
         */

    }, {
        key: "findByPartName",
        value: function findByPartName(partName) {
            return _.find(this._node.children, function (node) {
                return node.attributes.PartName === partName;
            });
        }

        /**
         * Convert the collection to an XML object.
         * @returns {{}} The XML.
         */

    }, {
        key: "toXml",
        value: function toXml() {
            return this._node;
        }
    }]);

    return ContentTypes;
}();

module.exports = ContentTypes;

/*
[Content_Types].xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
*/