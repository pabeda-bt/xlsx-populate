"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");

var RELATIONSHIP_SCHEMA_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/**
 * A relationship collection.
 * @ignore
 */

var Relationships = function () {
    /**
     * Creates a new instance of _Relationships.
     * @param {{}} node - The node.
     */
    function Relationships(node) {
        _classCallCheck(this, Relationships);

        this._init(node);
        this._getStartingId();
    }

    /**
     * Add a new relationship.
     * @param {string} type - The type of relationship.
     * @param {string} target - The target of the relationship.
     * @param {string} [targetMode] - The target mode of the relationship.
     * @returns {{}} The new relationship.
     */


    _createClass(Relationships, [{
        key: "add",
        value: function add(type, target, targetMode) {
            var node = {
                name: "Relationship",
                attributes: {
                    Id: "rId" + this._nextId++,
                    Type: "" + RELATIONSHIP_SCHEMA_PREFIX + type,
                    Target: target
                }
            };

            if (targetMode) {
                node.attributes.TargetMode = targetMode;
            }

            this._node.children.push(node);
            return node;
        }

        /**
         * Find a relationship by ID.
         * @param {string} id - The relationship ID.
         * @returns {{}|undefined} The matching relationship or undefined if not found.
         */

    }, {
        key: "findById",
        value: function findById(id) {
            return _.find(this._node.children, function (node) {
                return node.attributes.Id === id;
            });
        }

        /**
         * Find a relationship by type.
         * @param {string} type - The type to search for.
         * @returns {{}|undefined} The matching relationship or undefined if not found.
         */

    }, {
        key: "findByType",
        value: function findByType(type) {
            return _.find(this._node.children, function (node) {
                return node.attributes.Type === "" + RELATIONSHIP_SCHEMA_PREFIX + type;
            });
        }

        /**
         * Convert the collection to an XML object.
         * @returns {{}|undefined} The XML or undefined if empty.
         */

    }, {
        key: "toXml",
        value: function toXml() {
            if (!this._node.children.length) return;
            return this._node;
        }

        /**
         * Get the starting relationship ID to use for new relationships.
         * @private
         * @returns {undefined}
         */

    }, {
        key: "_getStartingId",
        value: function _getStartingId() {
            var _this = this;

            this._nextId = 1;
            this._node.children.forEach(function (node) {
                var id = parseInt(node.attributes.Id.substr(3));
                if (id >= _this._nextId) _this._nextId = id + 1;
            });
        }

        /**
         * Initialize the node.
         * @param {{}} [node] - The relationships node.
         * @private
         * @returns {undefined}
         */

    }, {
        key: "_init",
        value: function _init(node) {
            if (!node) node = {
                name: "Relationships",
                attributes: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
                },
                children: []
            };

            this._node = node;
        }
    }]);

    return Relationships;
}();

module.exports = Relationships;

/*
xl/_rels/workbook.xml.rels

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
*/