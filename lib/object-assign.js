"use strict"

if (typeof Object.assign != "function") {
  (function () {
    Object.assign = function (target) {
      if (target === undefined || target === null) {
        throw new TypeError("Cannot convert undefined or null to object");
      }

      var output = Object(target);
      for (var i = 1, n = arguments.length; i < n; i++) {
        var source = arguments[i];
        if (source !== undefined && source !== null) {
          for (var nextKey in source) {
            if (source.hasOwnProperty(nextKey)) {
              output[nextKey] = source[nextKey];
            }
          }
        }
      }
      return output;
    };
  })();
}
