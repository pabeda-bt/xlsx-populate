"use strict";

/* eslint camelcase:off */

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var ArgHandler = require("./ArgHandler");
var _ = require("lodash");
var xmlq = require("./xmlq");
var colorIndexes = require("./colorIndexes");

/**
 * A style.
 * @ignore
 */

var Style = function () {
    /**
     * Creates a new instance of _Style.
     * @constructor
     * @param {StyleSheet} styleSheet - The styleSheet.
     * @param {number} id - The style ID.
     * @param {{}} xfNode - The xf node.
     * @param {{}} fontNode - The font node.
     * @param {{}} fillNode - The fill node.
     * @param {{}} borderNode - The border node.
     */
    function Style(styleSheet, id, xfNode, fontNode, fillNode, borderNode) {
        _classCallCheck(this, Style);

        this._styleSheet = styleSheet;
        this._id = id;
        this._xfNode = xfNode;
        this._fontNode = fontNode;
        this._fillNode = fillNode;
        this._borderNode = borderNode;
    }

    /**
     * Gets the style ID.
     * @returns {number} The ID.
     */


    _createClass(Style, [{
        key: "id",
        value: function id() {
            return this._id;
        }

        /**
         * Gets or sets a style.
         * @param {string} name - The style name.
         * @param {*} [value] - The value to set.
         * @returns {*|Style} The value if getting or the style if setting.
         */

    }, {
        key: "style",
        value: function style() {
            var _this = this;

            return new ArgHandler("_Style.style").case('string', function (name) {
                var getterName = "_get_" + name;
                if (!_this[getterName]) throw new Error("_Style.style: '" + name + "' is not a valid style");
                return _this[getterName]();
            }).case(['string', '*'], function (name, value) {
                var setterName = "_set_" + name;
                if (!_this[setterName]) throw new Error("_Style.style: '" + name + "' is not a valid style");
                _this[setterName](value);
                return _this;
            }).handle(arguments);
        }
    }, {
        key: "_getColor",
        value: function _getColor(node, name) {
            var child = xmlq.findChild(node, name);
            if (!child || !child.attributes) return;

            var color = {};
            if (child.attributes.hasOwnProperty('rgb')) color.rgb = child.attributes.rgb;else if (child.attributes.hasOwnProperty('theme')) color.theme = child.attributes.theme;else if (child.attributes.hasOwnProperty('indexed')) color.rgb = colorIndexes[child.attributes.indexed];

            if (child.attributes.hasOwnProperty('tint')) color.tint = child.attributes.tint;

            if (_.isEmpty(color)) return;

            return color;
        }
    }, {
        key: "_setColor",
        value: function _setColor(node, name, color) {
            if (typeof color === "string") color = { rgb: color };else if (typeof color === "number") color = { theme: color };

            xmlq.setChildAttributes(node, name, {
                rgb: color && color.rgb && color.rgb.toUpperCase(),
                indexed: null,
                theme: color && color.theme,
                tint: color && color.tint
            });

            xmlq.removeChildIfEmpty(node, 'color');
        }
    }, {
        key: "_get_bold",
        value: function _get_bold() {
            return xmlq.hasChild(this._fontNode, 'b');
        }
    }, {
        key: "_set_bold",
        value: function _set_bold(bold) {
            if (bold) xmlq.appendChildIfNotFound(this._fontNode, "b");else xmlq.removeChild(this._fontNode, 'b');
        }
    }, {
        key: "_get_italic",
        value: function _get_italic() {
            return xmlq.hasChild(this._fontNode, 'i');
        }
    }, {
        key: "_set_italic",
        value: function _set_italic(italic) {
            if (italic) xmlq.appendChildIfNotFound(this._fontNode, "i");else xmlq.removeChild(this._fontNode, 'i');
        }
    }, {
        key: "_get_underline",
        value: function _get_underline() {
            var uNode = xmlq.findChild(this._fontNode, 'u');
            return uNode ? uNode.attributes.val || true : false;
        }
    }, {
        key: "_set_underline",
        value: function _set_underline(underline) {
            if (underline) {
                var uNode = xmlq.appendChildIfNotFound(this._fontNode, "u");
                var val = typeof underline === 'string' ? underline : null;
                xmlq.setAttributes(uNode, { val: val });
            } else {
                xmlq.removeChild(this._fontNode, 'u');
            }
        }
    }, {
        key: "_get_strikethrough",
        value: function _get_strikethrough() {
            return xmlq.hasChild(this._fontNode, 'strike');
        }
    }, {
        key: "_set_strikethrough",
        value: function _set_strikethrough(strikethrough) {
            if (strikethrough) xmlq.appendChildIfNotFound(this._fontNode, "strike");else xmlq.removeChild(this._fontNode, 'strike');
        }
    }, {
        key: "_getFontVerticalAlignment",
        value: function _getFontVerticalAlignment() {
            return xmlq.getChildAttribute(this._fontNode, 'vertAlign', "val");
        }
    }, {
        key: "_setFontVerticalAlignment",
        value: function _setFontVerticalAlignment(alignment) {
            xmlq.setChildAttributes(this._fontNode, 'vertAlign', { val: alignment });
            xmlq.removeChildIfEmpty(this._fontNode, 'vertAlign');
        }
    }, {
        key: "_get_subscript",
        value: function _get_subscript() {
            return this._getFontVerticalAlignment() === "subscript";
        }
    }, {
        key: "_set_subscript",
        value: function _set_subscript(subscript) {
            this._setFontVerticalAlignment(subscript ? "subscript" : null);
        }
    }, {
        key: "_get_superscript",
        value: function _get_superscript() {
            return this._getFontVerticalAlignment() === "superscript";
        }
    }, {
        key: "_set_superscript",
        value: function _set_superscript(superscript) {
            this._setFontVerticalAlignment(superscript ? "superscript" : null);
        }
    }, {
        key: "_get_fontSize",
        value: function _get_fontSize() {
            return xmlq.getChildAttribute(this._fontNode, 'sz', "val");
        }
    }, {
        key: "_set_fontSize",
        value: function _set_fontSize(size) {
            xmlq.setChildAttributes(this._fontNode, 'sz', { val: size });
            xmlq.removeChildIfEmpty(this._fontNode, 'sz');
        }
    }, {
        key: "_get_fontFamily",
        value: function _get_fontFamily() {
            return xmlq.getChildAttribute(this._fontNode, 'name', "val");
        }
    }, {
        key: "_set_fontFamily",
        value: function _set_fontFamily(family) {
            xmlq.setChildAttributes(this._fontNode, 'name', { val: family });
            xmlq.removeChildIfEmpty(this._fontNode, 'name');
        }
    }, {
        key: "_get_fontColor",
        value: function _get_fontColor() {
            return this._getColor(this._fontNode, "color");
        }
    }, {
        key: "_set_fontColor",
        value: function _set_fontColor(color) {
            this._setColor(this._fontNode, "color", color);
        }
    }, {
        key: "_get_horizontalAlignment",
        value: function _get_horizontalAlignment() {
            return xmlq.getChildAttribute(this._xfNode, 'alignment', "horizontal");
        }
    }, {
        key: "_set_horizontalAlignment",
        value: function _set_horizontalAlignment(alignment) {
            xmlq.setChildAttributes(this._xfNode, 'alignment', { horizontal: alignment });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_get_justifyLastLine",
        value: function _get_justifyLastLine() {
            return xmlq.getChildAttribute(this._xfNode, 'alignment', "justifyLastLine") === 1;
        }
    }, {
        key: "_set_justifyLastLine",
        value: function _set_justifyLastLine(justifyLastLine) {
            xmlq.setChildAttributes(this._xfNode, 'alignment', { justifyLastLine: justifyLastLine ? 1 : null });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_get_indent",
        value: function _get_indent() {
            return xmlq.getChildAttribute(this._xfNode, 'alignment', "indent");
        }
    }, {
        key: "_set_indent",
        value: function _set_indent(indent) {
            xmlq.setChildAttributes(this._xfNode, 'alignment', { indent: indent });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_get_verticalAlignment",
        value: function _get_verticalAlignment() {
            return xmlq.getChildAttribute(this._xfNode, 'alignment', "vertical");
        }
    }, {
        key: "_set_verticalAlignment",
        value: function _set_verticalAlignment(alignment) {
            xmlq.setChildAttributes(this._xfNode, 'alignment', { vertical: alignment });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_get_wrapText",
        value: function _get_wrapText() {
            return xmlq.getChildAttribute(this._xfNode, 'alignment', "wrapText") === 1;
        }
    }, {
        key: "_set_wrapText",
        value: function _set_wrapText(wrapText) {
            xmlq.setChildAttributes(this._xfNode, 'alignment', { wrapText: wrapText ? 1 : null });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_get_shrinkToFit",
        value: function _get_shrinkToFit() {
            return xmlq.getChildAttribute(this._xfNode, 'alignment', "shrinkToFit") === 1;
        }
    }, {
        key: "_set_shrinkToFit",
        value: function _set_shrinkToFit(shrinkToFit) {
            xmlq.setChildAttributes(this._xfNode, 'alignment', { shrinkToFit: shrinkToFit ? 1 : null });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_get_textDirection",
        value: function _get_textDirection() {
            var readingOrder = xmlq.getChildAttribute(this._xfNode, 'alignment', "readingOrder");
            if (readingOrder === 1) return "left-to-right";
            if (readingOrder === 2) return "right-to-left";
            return readingOrder;
        }
    }, {
        key: "_set_textDirection",
        value: function _set_textDirection(textDirection) {
            var readingOrder = void 0;
            if (textDirection === "left-to-right") readingOrder = 1;else if (textDirection === "right-to-left") readingOrder = 2;
            xmlq.setChildAttributes(this._xfNode, 'alignment', { readingOrder: readingOrder });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_getTextRotation",
        value: function _getTextRotation() {
            return xmlq.getChildAttribute(this._xfNode, 'alignment', "textRotation");
        }
    }, {
        key: "_setTextRotation",
        value: function _setTextRotation(textRotation) {
            xmlq.setChildAttributes(this._xfNode, 'alignment', { textRotation: textRotation });
            xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
        }
    }, {
        key: "_get_textRotation",
        value: function _get_textRotation() {
            var textRotation = this._getTextRotation();

            // Negative angles in Excel correspond to values > 90 in OOXML.
            if (textRotation > 90) textRotation = 90 - textRotation;
            return textRotation;
        }
    }, {
        key: "_set_textRotation",
        value: function _set_textRotation(textRotation) {
            // Negative angles in Excel correspond to values > 90 in OOXML.
            if (textRotation < 0) textRotation = 90 - textRotation;
            this._setTextRotation(textRotation);
        }
    }, {
        key: "_get_angleTextCounterclockwise",
        value: function _get_angleTextCounterclockwise() {
            return this._getTextRotation() === 45;
        }
    }, {
        key: "_set_angleTextCounterclockwise",
        value: function _set_angleTextCounterclockwise(value) {
            this._setTextRotation(value ? 45 : null);
        }
    }, {
        key: "_get_angleTextClockwise",
        value: function _get_angleTextClockwise() {
            return this._getTextRotation() === 135;
        }
    }, {
        key: "_set_angleTextClockwise",
        value: function _set_angleTextClockwise(value) {
            this._setTextRotation(value ? 135 : null);
        }
    }, {
        key: "_get_rotateTextUp",
        value: function _get_rotateTextUp() {
            return this._getTextRotation() === 90;
        }
    }, {
        key: "_set_rotateTextUp",
        value: function _set_rotateTextUp(value) {
            this._setTextRotation(value ? 90 : null);
        }
    }, {
        key: "_get_rotateTextDown",
        value: function _get_rotateTextDown() {
            return this._getTextRotation() === 180;
        }
    }, {
        key: "_set_rotateTextDown",
        value: function _set_rotateTextDown(value) {
            this._setTextRotation(value ? 180 : null);
        }
    }, {
        key: "_get_verticalText",
        value: function _get_verticalText() {
            return this._getTextRotation() === 255;
        }
    }, {
        key: "_set_verticalText",
        value: function _set_verticalText(value) {
            this._setTextRotation(value ? 255 : null);
        }
    }, {
        key: "_get_fill",
        value: function _get_fill() {
            var _this2 = this;

            var patternFillNode = xmlq.findChild(this._fillNode, 'patternFill'); // jq.get(this._fillNode, "patternFill[0]");
            var gradientFillNode = xmlq.findChild(this._fillNode, 'gradientFill'); // jq.get(this._fillNode, "gradientFill[0]");
            var patternType = patternFillNode && patternFillNode.attributes.patternType; // jq.get(patternFillNode, "$.patternType");

            if (patternType === "solid") {
                return {
                    type: "solid",
                    color: this._getColor(patternFillNode, "fgColor")
                };
            }

            if (patternType) {
                return {
                    type: "pattern",
                    pattern: patternType,
                    foreground: this._getColor(patternFillNode, "fgColor"),
                    background: this._getColor(patternFillNode, "bgColor")
                };
            }

            if (gradientFillNode) {
                var gradientType = gradientFillNode.attributes.type || "linear";
                var fill = {
                    type: "gradient",
                    gradientType: gradientType,
                    stops: _.map(gradientFillNode.children, function (stop) {
                        return {
                            position: stop.attributes.position,
                            color: _this2._getColor(stop, "color")
                        };
                    })
                };

                if (gradientType === "linear") {
                    fill.angle = gradientFillNode.attributes.degree;
                } else {
                    fill.left = gradientFillNode.attributes.left;
                    fill.right = gradientFillNode.attributes.right;
                    fill.top = gradientFillNode.attributes.top;
                    fill.bottom = gradientFillNode.attributes.bottom;
                }

                return fill;
            }
        }
    }, {
        key: "_set_fill",
        value: function _set_fill(fill) {
            var _this3 = this;

            this._fillNode.children = [];

            // No fill
            if (_.isNil(fill)) return;

            // Pattern fill
            if (fill.type === "pattern") {
                var _patternFill = {
                    name: 'patternFill',
                    attributes: { patternType: fill.pattern },
                    children: []
                };
                this._fillNode.children.push(_patternFill);
                this._setColor(_patternFill, "fgColor", fill.foreground);
                this._setColor(_patternFill, "bgColor", fill.background);
                return;
            }

            // Gradient fill
            if (fill.type === "gradient") {
                var gradientFill = { name: 'gradientFill', attributes: {}, children: [] };
                this._fillNode.children.push(gradientFill);
                xmlq.setAttributes(gradientFill, {
                    type: fill.gradientType === "path" ? "path" : undefined,
                    left: fill.left,
                    right: fill.right,
                    top: fill.top,
                    bottom: fill.bottom,
                    degree: fill.angle
                });

                _.forEach(fill.stops, function (fillStop, i) {
                    var stop = {
                        name: 'stop',
                        attributes: { position: fillStop.position },
                        children: []
                    };
                    gradientFill.children.push(stop);
                    _this3._setColor(stop, 'color', fillStop.color);
                });

                return;
            }

            // Solid fill (really a pattern fill with a solid pattern type).
            if (!_.isObject(fill)) fill = { type: "solid", color: fill };else if (fill.hasOwnProperty('rgb') || fill.hasOwnProperty("theme")) fill = { color: fill };

            var patternFill = {
                name: 'patternFill',
                attributes: { patternType: 'solid' }
            };
            this._fillNode.children.push(patternFill);
            this._setColor(patternFill, "fgColor", fill.color);
        }
    }, {
        key: "_getBorder",
        value: function _getBorder() {
            var _this4 = this;

            var result = {};
            ["left", "right", "top", "bottom", "diagonal"].forEach(function (side) {
                var sideNode = xmlq.findChild(_this4._borderNode, side);
                var sideResult = {};

                var style = xmlq.getChildAttribute(_this4._borderNode, side, 'style');
                if (style) sideResult.style = style;
                var color = _this4._getColor(sideNode, 'color');
                if (color) sideResult.color = color;

                if (side === "diagonal") {
                    var up = _this4._borderNode.attributes.diagonalUp;
                    var down = _this4._borderNode.attributes.diagonalDown;
                    var direction = void 0;
                    if (up && down) direction = "both";else if (up) direction = "up";else if (down) direction = "down";
                    if (direction) sideResult.direction = direction;
                }

                if (!_.isEmpty(sideResult)) result[side] = sideResult;
            });

            return result;
        }
    }, {
        key: "_setBorder",
        value: function _setBorder(settings) {
            var _this5 = this;

            _.forOwn(settings, function (setting, side) {
                if (typeof setting === "boolean") {
                    setting = { style: setting ? "thin" : null };
                } else if (typeof setting === "string") {
                    setting = { style: setting };
                } else if (setting === null || setting === undefined) {
                    setting = { style: null, color: null, direction: null };
                }

                if (setting.hasOwnProperty("style")) {
                    xmlq.setChildAttributes(_this5._borderNode, side, { style: setting.style });
                }

                if (setting.hasOwnProperty("color")) {
                    var sideNode = xmlq.findChild(_this5._borderNode, side);
                    _this5._setColor(sideNode, "color", setting.color);
                }

                if (side === "diagonal") {
                    xmlq.setAttributes(_this5._borderNode, {
                        diagonalUp: setting.direction === "up" || setting.direction === "both" ? 1 : null,
                        diagonalDown: setting.direction === "down" || setting.direction === "both" ? 1 : null
                    });
                }
            });
        }
    }, {
        key: "_get_border",
        value: function _get_border() {
            return this._getBorder();
        }
    }, {
        key: "_set_border",
        value: function _set_border(settings) {
            if (_.isObject(settings) && !settings.hasOwnProperty("style") && !settings.hasOwnProperty("color")) {
                settings = _.defaults(settings, {
                    left: null,
                    right: null,
                    top: null,
                    bottom: null,
                    diagonal: null
                });
                this._setBorder(settings);
            } else {
                this._setBorder({
                    left: settings,
                    right: settings,
                    top: settings,
                    bottom: settings
                });
            }
        }
    }, {
        key: "_get_borderColor",
        value: function _get_borderColor() {
            return _.mapValues(this._getBorder(), function (value) {
                return value.color;
            });
        }
    }, {
        key: "_set_borderColor",
        value: function _set_borderColor(color) {
            if (_.isObject(color)) {
                this._setBorder(_.mapValues(color, function (color) {
                    return { color: color };
                }));
            } else {
                this._setBorder({
                    left: { color: color },
                    right: { color: color },
                    top: { color: color },
                    bottom: { color: color },
                    diagonal: { color: color }
                });
            }
        }
    }, {
        key: "_get_borderStyle",
        value: function _get_borderStyle() {
            return _.mapValues(this._getBorder(), function (value) {
                return value.style;
            });
        }
    }, {
        key: "_set_borderStyle",
        value: function _set_borderStyle(style) {
            if (_.isObject(style)) {
                this._setBorder(_.mapValues(style, function (style) {
                    return { style: style };
                }));
            } else {
                this._setBorder({
                    left: { style: style },
                    right: { style: style },
                    top: { style: style },
                    bottom: { style: style }
                });
            }
        }
    }, {
        key: "_get_diagonalBorderDirection",
        value: function _get_diagonalBorderDirection() {
            var border = this._getBorder().diagonal;
            return border && border.direction;
        }
    }, {
        key: "_set_diagonalBorderDirection",
        value: function _set_diagonalBorderDirection(direction) {
            this._setBorder({ diagonal: { direction: direction } });
        }
    }, {
        key: "_get_numberFormat",
        value: function _get_numberFormat() {
            var numFmtId = this._xfNode.attributes.numFmtId || 0;
            return this._styleSheet.getNumberFormatCode(numFmtId);
        }
    }, {
        key: "_set_numberFormat",
        value: function _set_numberFormat(formatCode) {
            this._xfNode.attributes.numFmtId = this._styleSheet.getNumberFormatId(formatCode);
        }
    }]);

    return Style;
}();

["left", "right", "top", "bottom", "diagonal"].forEach(function (side) {
    Style.prototype["_get_" + side + "Border"] = function () {
        return this._getBorder()[side];
    };

    Style.prototype["_set_" + side + "Border"] = function (settings) {
        this._setBorder(_defineProperty({}, side, settings));
    };

    Style.prototype["_get_" + side + "BorderColor"] = function () {
        var border = this._getBorder()[side];
        return border && border.color;
    };

    Style.prototype["_set_" + side + "BorderColor"] = function (color) {
        this._setBorder(_defineProperty({}, side, { color: color }));
    };

    Style.prototype["_get_" + side + "BorderStyle"] = function () {
        var border = this._getBorder()[side];
        return border && border.style;
    };

    Style.prototype["_set_" + side + "BorderStyle"] = function (style) {
        this._setBorder(_defineProperty({}, side, { style: style }));
    };
});

// IE doesn't support function names so explicitly set it.
if (!Style.name) Style.name = "Style";

module.exports = Style;