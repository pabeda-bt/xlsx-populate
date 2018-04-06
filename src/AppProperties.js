"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");
var xmlq = require("./xmlq");
var ArgHandler = require("./ArgHandler");

/**
 * App properties
 * @ignore
 */

var AppProperties = function () {
    /**
     * Creates a new instance of AppProperties
     * @param {{}} node - The node.
     */
    function AppProperties(node) {
        _classCallCheck(this, AppProperties);

        this._node = node;
    }

    _createClass(AppProperties, [{
        key: "isSecure",
        value: function isSecure(value) {
            var _this = this;

            return new ArgHandler("Range.formula").case(function () {
                var docSecurityNode = xmlq.findChild(_this._node, "DocSecurity");
                if (!docSecurityNode) return false;
                return docSecurityNode.children[0] === 1;
            }).case('boolean', function (value) {
                var docSecurityNode = xmlq.appendChildIfNotFound(_this._node, "DocSecurity");
                docSecurityNode.children = [value ? 1 : 0];
                return _this;
            }).handle(arguments);
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

    return AppProperties;
}();

module.exports = AppProperties;

/*
docProps/app.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Microsoft Excel</Application>
<DocSecurity>1</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
    <vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>1</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="1" baseType="lpstr">
    <vt:lpstr>Sheet1</vt:lpstr>
</vt:vector>
</TitlesOfParts>
<Company/>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0300</AppVersion>
</Properties>
 */