// requires
require("date-utils");
require("string-utils-cwm");
require("./lib/object-assign");

var _ = require("underscore");
var xml2js = require("xml2js").parseString;
var httpntlm = require("httpntlm");

var async = require("async");
var request = require("request");
var helper = require("./lib/hash-get");

var process_env = _.pick(process.env, "EWS_DOMAIN", "EWS_USER", "EWS_PASS", "EWS_URL");
var config = {
	domain: process_env.EWS_DOMAIN || "",
	user: process_env.EWS_USER     || "",
	passwd: process_env.EWS_PASS   || "",
	url: process_env.EWS_URL       || "",

	start: Date.today(),
	start_string: Date.today().toISOString().replace(/\.[0-9]{3}Z/, "Z"),
	end: Date.today().add({days: 1}),
	end_string: Date.today().add({days: 1}).toISOString().replace(/\.[0-9]{3}Z/, "Z")
};

Object.keys(config).forEach(function(prop) {
	if (!config[prop]) {
		console.log("Error: env variable '{0}' not set!".format(prop));
		process.exit(1);
	}
});

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = '0';

var post_headers = {
	"User-Agent": "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
	"Content-Type": "text/xml; charset=utf-8",
	"SOAPAction": '"{soap_action}"'
};

var post_items = {
	soap_action: "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
	body: [
		'<?xml version="1.0" encoding="UTF-8"?>',
		'<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:ns2="http://schemas.microsoft.com/exchange/services/2006/messages">',
		'  <SOAP-ENV:Header>',
		'    <ns1:RequestServerVersion Version="Exchange2010"/>',
		'  </SOAP-ENV:Header>',
		'  <SOAP-ENV:Body>',
		'    <ns2:FindItem Traversal="Shallow">',
		'      <ns2:ItemShape>',
		'        <ns1:BaseShape>Default</ns1:BaseShape>',
		'      </ns2:ItemShape>',
		'      <ns2:CalendarView StartDate="{start_string}" EndDate="{end_string}"/>',
		'      <ns2:ParentFolderIds>',
		'        <ns1:DistinguishedFolderId Id="calendar"/>',
		'      </ns2:ParentFolderIds>',
		'    </ns2:FindItem>',
		'  </SOAP-ENV:Body>',
		'</SOAP-ENV:Envelope>'
	].join("").format(config)
};

Object.defineProperty(post_items, "headers", {
	get: function() {
		var headers = Object.assign({}, post_headers);
		headers.SOAPAction = headers.SOAPAction.format(this.soap_action);;
		return headers;
	}
});


var post_item = {
	soap_action: "http://schemas.microsoft.com/exchange/services/2006/messages/GetItem",
	body: [
		'<?xml version="1.0" encoding="UTF-8"?>',
		'<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:ns2="http://schemas.microsoft.com/exchange/services/2006/messages">',
		'  <SOAP-ENV:Header>',
		'    <ns1:RequestServerVersion Version="Exchange2010"/>',
		'  </SOAP-ENV:Header>',
		'  <SOAP-ENV:Body>',
		'    <ns2:GetItem>',
		'      <ns2:ItemShape>',
		'        <ns1:BaseShape>AllProperties</ns1:BaseShape>',
		'      </ns2:ItemShape>',
		'      <ns2:ItemIds>',
		'        <ns1:ItemId Id="{id}" ChangeKey="{change_key}"/>',
		'      </ns2:ItemIds>',
		'    </ns2:GetItem>',
		'  </SOAP-ENV:Body>',
		'</SOAP-ENV:Envelope>',
	].join("")
};

Object.defineProperty(post_item, "headers", {
	get: function() {
		var headers = Object.assign({}, post_headers);
		headers.SOAPAction = headers.SOAPAction.format(this.soap_action);;
		return headers;
	}
});

helper.format_url = function(p){return p && p.length === 2 ? "https://{0}".format(p[1]) : ""};
helper.format_date = function(d){return d ? d.toFormat("YYYY-MM-DD HH:MI:SS ") : ""};
helper.format_time = function(d){return d ? d.toFormat("H:MI P") : ""};

httpntlm.post({
		url: "https://{url}/EWS/Exchange.asmx".format(config),
		username: config.user,
		password: config.passwd,
		domain: config.domain,
		headers: post_items.headers,
		body: post_items.body
	},
	function (err, res) {
		var body = (""+res.body).replace(/[smt]:/g, "");
		xml2js(body, function(err_not_used, result) {
			var items = helper.hash_get(result, "Envelope.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder.Items.CalendarItem", []);

			async.map(items,
				function(item, callback) {
					var attrs = helper.hash_get(item, "ItemId.$", {});
					var params = {
						id: helper.hash_get(attrs, "Id", ""),
						change_key: helper.hash_get(attrs, "ChangeKey", "")
					};

					httpntlm.post({
							url: "https://{url}/EWS/Exchange.asmx".format(config),
							username: config.user,
							password: config.passwd,
							domain: config.domain,
							headers: post_item.headers,
							body: post_item.body.format(params)
						},
						function (err, res) {
							var body = (""+res.body).replace(/[smt]:/g, "");
							xml2js(body, function(err_not_used, result) {
								var calendar_item = helper.hash_get(result, "Envelope.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage.Items.CalendarItem", []).pop();
								var body = helper.hash_get(calendar_item, "Body._", "");

								// special ones
								item["G2M"] = helper.format_url(/([^\/]+\.gotomeeting\.com\/join\/[0-9]{9})/.exec(body));
								item["Webex"] = helper.format_url(/([^\/]+\.webex\.com[^<>"]+)/.exec(body));

								// simple values
								item["ItemId"] = params.id;
								item["ChangeKey"] = params.change_key;
								item["Organizer"] = helper.hash_get(item, "Organizer.Mailbox.Name", "");

								// local dates
								item["StartLocal"] = helper.format_date(new Date(item["Start"]));
								item["EndLocal"] = helper.format_date(new Date(item["End"]));

								// reduce the arrays of 1 to values
								_.each(item, function(value, key, _list) {
									if (helper.isArray(value) && value.length === 1) {
										_list[key] = value = value[0];
									}

									// t/f to 1|0 instead of strings
									if (value === "true") { _list[key] = true; }
									if (value === "false") { _list[key] = false; }
								});

								callback(null, item);
							});
						}
					);
				},
				function(err, results) {

					var text = [];
					var standup = [];

					_.each(_.indexBy(results, "StartLocal"), function(meeting/*, key, _list*/) {
						var time_until = new Date().getMinutesBetween(new Date(meeting.Start));

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


						var display = [ "- ", "{0}".format(meeting.Subject).replace(/\[[^\]]*\]/g, "").trim() ];

						text.push("Subject {0}".format(meeting.Subject).trim())

						if (meeting.Webex) {
							text.push("    URL {0}".format(meeting.Webex));
							display.push("@ <{0}|{1}>".format(meeting.Webex, helper.format_time(new Date(meeting.Start))));
						}
						else if (meeting.G2M) {
							text.push("    URL {0}".format(meeting.G2M));
							display.push("@ <{0}|{1}>".format(meeting.G2M, helper.format_time(new Date(meeting.Start))));
						}
						else {
							display.push("@ {0}".format(helper.format_time(new Date(meeting.Start))));
						}

						text.push("   Time {0} {1}".format(meeting.StartLocal, until_msg));
						text.push(" Length {0} minutes".format(
							new Date(meeting.Start).getMinutesBetween(new Date(meeting.End))
						));

						standup.push(display.join(" "));
						text.push("");
					});

					//console.log("\n", text.join("\n"), "\n");

					if (standup.length) {
						console.log("\"* Meeting{0}: \",".format(standup.length > 1 ? "s" : ""));
						console.log("\"" + standup.join("\",\n\"") + "\"", "\n\n" );
					}
					else {
						console.log("\"* No Meetings\"" , "\n\n");
					}

				}
			);

		});
	});

