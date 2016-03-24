// helpers
var toString = Object.prototype.toString;
var isArray = function(value){return toString.apply(value) === "[object Array]";};
var isFunction = function(value){return typeof value === "function";};
var isObject = function(value){return value != null && typeof value === "object";};

// real helper
var hash_get = function(hash, key, default_result) {
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
	return current === hash ? default_result : current;
};

module.exports = {
	toString: toString,
	isArray: isArray,
	isFunction: isFunction,
	isObject: isObject,
	hash_get: hash_get
};
