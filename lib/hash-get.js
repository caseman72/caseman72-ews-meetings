require("caseman72-string-utils");

var moment = require("moment-timezone");

// helpers
var toString   = Object.prototype.toString;
var isArray    = function(v) { return toString.apply(v) === "[object Array]"; };
var isFunction = function(v) { return typeof v === "function"; };
var isObject   = function(v) { return v != null && typeof v === "object"; };

var formatUrl  = function(p) { return p && p.length === 2 ? "https://{0}".format(p[1]) : ""; };
var formatDate = function(d, tz) { return d ? moment.tz(d, tz).format("YYYY-MM-DD HH:mm:ss") : ""; };
var formatTime = function(d, tz) { return d ? moment.tz(d, tz).format("h:mm a z") : ""; };
var nextDay    = function(d, tz) { return moment.tz(d, tz).add(24, "hours").toDate(); };

// real helper
var hashGet = function(hash, key, defaultResult) {
	var current = hash;

	var keys = "{0}".format(key).split(".");
	for(var i=0,n=keys.length; i<n; i++) {
		if ((isObject(current) || isFunction(current)) && keys[i] !== "" && typeof current[keys[i]] !== "undefined") {
			current = current[keys[i]];
			if (isArray(current) && current.length === 1 && keys[i] !== "CalendarItem") {
				current = current[0];
			}
		}
		else {
			current = hash;
			break;
		}
	}
	return current === hash ? defaultResult : current;
};

module.exports = {
	toString: toString,
	isArray: isArray,
	isFunction: isFunction,
	isObject: isObject,
	hashGet: hashGet,
	nextDay: nextDay,
	formatUrl: formatUrl,
	formatDate: formatDate,
	formatTime: formatTime
};
