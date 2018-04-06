"use strict";

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require("lodash");

/**
 * A formula error (e.g. #DIV/0!).
 */

var FormulaError = function () {
  // /**
  //  * Creates a new instance of Formula Error.
  //  * @param {string} error - The error code.
  //  */
  function FormulaError(error) {
    _classCallCheck(this, FormulaError);

    this._error = error;
  }

  /**
   * Get the error code.
   * @returns {string} The error code.
   */


  _createClass(FormulaError, [{
    key: "error",
    value: function error() {
      return this._error;
    }
  }]);

  return FormulaError;
}();

/**
 * \#DIV/0! error.
 * @type {FormulaError}
 */


FormulaError.DIV0 = new FormulaError("#DIV/0!");

/**
 * \#N/A error.
 * @type {FormulaError}
 */
FormulaError.NA = new FormulaError("#N/A");

/**
 * \#NAME? error.
 * @type {FormulaError}
 */
FormulaError.NAME = new FormulaError("#NAME?");

/**
 * \#NULL! error.
 * @type {FormulaError}
 */
FormulaError.NULL = new FormulaError("#NULL!");

/**
 * \#NUM! error.
 * @type {FormulaError}
 */
FormulaError.NUM = new FormulaError("#NUM!");

/**
 * \#REF! error.
 * @type {FormulaError}
 */
FormulaError.REF = new FormulaError("#REF!");

/**
 * \#VALUE! error.
 * @type {FormulaError}
 */
FormulaError.VALUE = new FormulaError("#VALUE!");

/**
 * Get the matching FormulaError object.
 * @param {string} error - The error code.
 * @returns {FormulaError} The matching FormulaError or a new object if no match.
 * @ignore
 */
FormulaError.getError = function (error) {
  return _.find(FormulaError, function (value) {
    return value instanceof FormulaError && value.error() === error;
  }) || new FormulaError(error);
};

module.exports = FormulaError;