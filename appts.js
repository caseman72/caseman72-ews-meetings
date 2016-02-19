// requires
require("date-utils");
require("string-utils-cwm");

var _ = require("underscore");
var async = require("async");
var exec = require("child_process").exec
var xml2js = require("xml2js").parseString;

var process_env = _.pick(process.env, "EWS_DOMAIN", "EWS_USER", "EWS_PASS", "EWS_URL");
var config = {
	domain: process_env.EWS_DOMAIN || "",
	user: process_env.EWS_USER     || "",
	passwd: process_env.EWS_PASS   || "",
	url: process_env.EWS_URL       || ""
};

Object.keys(config).forEach(function(prop) {
	if (!config[prop]) {
		console.log("Error: env variable '{0}' not set!".format(prop));
		process.exit(1);
	}
});

// curl_get_calendar_items
var curl_get_calendar_items = [
	'curl --ntlm -k -s',
	'-H "Method: POST"',
	'-H "Connection: Keep-Alive"',
	'-H "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)"',
	'-H "Content-Type: text/xml; charset=utf-8"',
	'-H \'SOAPAction: "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem"\'',
	'-u \'{domain}\\{user}:{passwd}\'',
	'-d \'<?xml version="1.0" encoding="UTF-8"?><SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:ns2="http://schemas.microsoft.com/exchange/services/2006/messages"><SOAP-ENV:Header><ns1:RequestServerVersion Version="Exchange2010"/></SOAP-ENV:Header><SOAP-ENV:Body><ns2:FindItem Traversal="Shallow"><ns2:ItemShape><ns1:BaseShape>Default</ns1:BaseShape></ns2:ItemShape><ns2:CalendarView StartDate="{start}" EndDate="{end}"/><ns2:ParentFolderIds><ns1:DistinguishedFolderId Id="calendar"/></ns2:ParentFolderIds></ns2:FindItem></SOAP-ENV:Body></SOAP-ENV:Envelope>\'',
	'https://{url}/EWS/Exchange.asmx'
].join(" ");

// curl_get_calendar_item
var curl_get_calendar_item = [
	'curl --ntlm -k -s',
	'-H "Method: POST"',
	'-H "Connection: Keep-Alive"',
	'-H "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)"',
	'-H "Content-Type: text/xml; charset=utf-8"',
	'-H \'SOAPAction: "http://schemas.microsoft.com/exchange/services/2006/messages/GetItem"\'',
	'-u \'{domain}\\{user}:{passwd}\'',
	'-d \'<?xml version="1.0" encoding="UTF-8"?><SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:ns2="http://schemas.microsoft.com/exchange/services/2006/messages"><SOAP-ENV:Header><ns1:RequestServerVersion Version="Exchange2010"/></SOAP-ENV:Header><SOAP-ENV:Body><ns2:GetItem><ns2:ItemShape><ns1:BaseShape>AllProperties</ns1:BaseShape></ns2:ItemShape><ns2:ItemIds><ns1:ItemId Id="{id}" ChangeKey="{change_key}"/></ns2:ItemIds></ns2:GetItem></SOAP-ENV:Body></SOAP-ENV:Envelope>\'',
	'https://{url}/EWS/Exchange.asmx'
].join(" ");

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

// get items
var get_calendar_items = function(start, end, callback) {
	start = start.toISOString().replace(/\.[0-9]{3}Z/, "Z");
	end = end.toISOString().replace(/\.[0-9]{3}Z/, "Z");

	exec(curl_get_calendar_items.format(config).format({start: start, end: end}), function(err_not_used, stdout/*, stderr*/) {
		stdout = (""+stdout).replace(/[smt]:/g, "");
		xml2js(stdout, function(err_not_used, result) {
			var items = hash_get(result, "Envelope.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder.Items.CalendarItem", []);
			async.each(items, get_calendar_item, function() { callback(items) });
		});
	});
};

// get item (body) -> G2M
var get_calendar_item = function(item, done) {
	var attrs = hash_get(item, "ItemId.$", {});
	var id = hash_get(attrs, "Id", "");
	var change_key = hash_get(attrs, "ChangeKey", "");

	exec(curl_get_calendar_item.format(config).format({id: id, change_key: change_key}), function(err_not_used, stdout/*, stderr*/) {
		stdout = (""+stdout).replace(/[smt]:/g, "");
		xml2js(stdout, function(err_not_used, result) {
			var calendar_item = hash_get(result, "Envelope.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage.Items.CalendarItem", []).pop();
			var body = hash_get(calendar_item, "Body._", "");

			// special ones
			item["G2M"] = (function(p){return p && p.length === 2 ? "https://{0}".format(p[1]) : ""})(/([^\/]+\.gotomeeting\.com\/join\/[0-9]{9})/.exec(body));
			item["Webex"] = (function(p){return p && p.length === 2 ? "https://{0}".format(p[1]) : ""})(/([^\/]+\.webex\.com[^<>"]+)/.exec(body));

			// simple values
			item["ItemId"] = id;
			item["ChangeKey"] = change_key;
			item["Organizer"] = hash_get(item, "Organizer.Mailbox.Name", "");

			// local dates
			item["StartLocal"] = (function(d){return d.toFormat("YYYY-MM-DD HH:MI:SS ") + d.getUTCOffset()})(new Date(item["Start"]));
			item["EndLocal"] = (function(d){return d.toFormat("YYYY-MM-DD HH:MI:SS ") + d.getUTCOffset()})(new Date(item["End"]));

			// reduce the arrays of 1 to values
			_.each(item, function(value, key, _list) {
				if (isArray(value) && value.length === 1) {
					_list[key] = value = value[0];
				}

				// t/f to 1|0 instead of strings
				if (value === "true") { _list[key] = true; }
				if (value === "false") { _list[key] = false; }
			});

			done();
		});
	});
};

get_calendar_items(Date.today(), Date.today().add({days: 1}), function(results) {
	var text = [];
	_.each(_.indexBy(results, "StartLocal"), function(meeting/*, key, _list*/) {
		var time_until = new Date().getMinutesBetween(new Date(meeting["Start"]));
		// nice message
		var until_msg = "[{0} minutes]".format(time_until);
		if (time_until < 2 && time_until > -2) {
			until_msg = "[now]";
		}
		else if (time_until < -10) {
			until_msg = "[completed]";
		}
		else if (time_until > 480) {
			until_msg = "";
		}

		text.push("Subject {0}".format(meeting["Subject"]))
		if (meeting["G2M"]) {
			text.push("    URL {0}".format(meeting["G2M"]));
		}
		if (meeting["Webex"]) {
			text.push("    URL {0}".format(meeting["Webex"]));
		}
		text.push("   Time {0} {1}".format(meeting["StartLocal"], until_msg));
		text.push(" Length {0} minutes".format(
			new Date(meeting["Start"]).getMinutesBetween(new Date(meeting["End"]))
		));

		text.push("");
	});

	console.log(text.join("\n"));
});
