"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var sax = require("sax");
var externals = require("./externals");

// Regex to check if string is all whitespace.
var allWhitespaceRegex = /^\s+$/;

/**
 * XML parser.
 * @private
 */

var XmlParser = function () {
    function XmlParser() {
        _classCallCheck(this, XmlParser);
    }

    _createClass(XmlParser, [{
        key: "parseAsync",

        /**
         * Parse the XML text into a JSON object.
         * @param {string} xmlText - The XML text.
         * @returns {{}} The JSON object.
         */
        value: function parseAsync(xmlText) {
            var _this = this;

            return new externals.Promise(function (resolve, reject) {
                // Create the SAX parser.
                var parser = sax.parser(true);

                // Parsed is the full parsed object. Current is the current node being parsed. Stack is the current stack of
                // nodes leading to the current one.
                var parsed = void 0,
                    current = void 0;
                var stack = [];

                // On error: Reject the promise.
                parser.onerror = reject;

                // On text nodes: If it is all whitespace, do nothing. Otherwise, try to convert to a number and add as a child.
                parser.ontext = function (text) {
                    if (allWhitespaceRegex.test(text)) {
                        if (current && current.attributes['xml:space'] === 'preserve') {
                            current.children.push(text);
                        }
                    } else {
                        current.children.push(_this._covertToNumberIfNumber(text));
                    }
                };

                // On open tag start: Create a child element. If this is the root element, set it as parsed. Otherwise, add
                // it as a child to the current node.
                parser.onopentagstart = function (node) {
                    var child = { name: node.name, attributes: {}, children: [] };
                    if (current) {
                        current.children.push(child);
                    } else {
                        parsed = child;
                    }

                    stack.push(child);
                    current = child;
                };

                // On close tag: Pop the stack.
                parser.onclosetag = function (node) {
                    stack.pop();
                    current = stack[stack.length - 1];
                };

                // On attribute: Try to convert the value to a number and add to the current node.
                parser.onattribute = function (attribute) {
                    current.attributes[attribute.name] = _this._covertToNumberIfNumber(attribute.value);
                };

                // On end: Resolve the promise.
                parser.onend = function () {
                    return resolve(parsed);
                };

                // Start parsing the text.
                parser.write(xmlText).close();
            });
        }

        /**
         * Convert the string to a number if it looks like one.
         * @param {string} str - The string to convert.
         * @returns {string|number} The number if converted or the string if not.
         * @private
         */

    }, {
        key: "_covertToNumberIfNumber",
        value: function _covertToNumberIfNumber(str) {
            var num = Number(str);
            return num.toString() === str ? num : str;
        }
    }]);

    return XmlParser;
}();

module.exports = XmlParser;