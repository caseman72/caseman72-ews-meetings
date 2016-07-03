// requires
var xml2js = require("xml2js").parseString;
var post = require("request").post
var async = require("async");
var _ = require("underscore");
var helper = require("./hash-get");

// skip env/cert issues
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

module.exports = function(args, callback) {
	// args -> config
	var process_env = _.pick(args, "EWS_TLD", "EWS_AUTH", "EWS_START", "EWS_TZ");
	var config = {
		tld: process_env.EWS_TLD || "",
		auth: process_env.EWS_AUTH || "",
		start: process_env.EWS_START || "",
		tz: process_env.EWS_TZ || ""
	};

	if (!config.tld || !config.auth) {
		throw new Error("Error: env variable '{0}' not set!".format(!config.tld ? "EWS_TLD" : "EWS_AUTH"));
	}

	Object.defineProperty(config, "startDate", {
		enumerable: true,
		get: function() {
			return helper.gmtStart(this.start ? this.start : new Date(), config.tz);
		}
	});
	Object.defineProperty(config, "startString", {
		enumerable: true,
		get: function() {
			return helper.formatIso(this.startDate, config.tz);
		}
	});
	Object.defineProperty(config, "endDate", {
		enumerable: true,
		get: function() {
			return helper.gmtEnd(this.startDate, config.tz);
		}
	});
	Object.defineProperty(config, "endString", {
		enumerable: true,
		get: function() {
			return helper.formatIso(this.endDate, config.tz);
		}
	});

	//console.log( JSON.stringify(config, null, 2)); process.exit();

	var postHeaders = {
		"User-Agent": "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
		"Content-Type": "text/xml; charset=utf-8",
		"SOAPAction": '"{soapAction}"'
	};

	var postItems = {
		soapAction: "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
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
			'      <ns2:CalendarView StartDate="{startString}" EndDate="{endString}"/>',
			'      <ns2:ParentFolderIds>',
			'        <ns1:DistinguishedFolderId Id="calendar"/>',
			'      </ns2:ParentFolderIds>',
			'    </ns2:FindItem>',
			'  </SOAP-ENV:Body>',
			'</SOAP-ENV:Envelope>'
		].join("").format(config)
	};

	Object.defineProperty(postItems, "headers", {
		get: function() {
			var headers = _.extend({}, postHeaders);
			headers.SOAPAction = headers.SOAPAction.format(this.soapAction);;
			headers.Authorization = "Basic {auth}".format(config)
			return headers;
		}
	});

	var postItem = {
		soapAction: "http://schemas.microsoft.com/exchange/services/2006/messages/GetItem",
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
			'        <ns1:ItemId Id="{id}" ChangeKey="{changeKey}"/>',
			'      </ns2:ItemIds>',
			'    </ns2:GetItem>',
			'  </SOAP-ENV:Body>',
			'</SOAP-ENV:Envelope>',
		].join("")
	};

	Object.defineProperty(postItem, "headers", {
		get: function() {
			var headers = Object.assign({}, postHeaders);
			headers.SOAPAction = headers.SOAPAction.format(this.soapAction);;
			headers.Authorization = "Basic {auth}".format(config)
			return headers;
		}
	});

	post({
			url: "https://{tld}/EWS/Exchange.asmx".format(config),
			headers: postItems.headers,
			body: postItems.body
		},
		function (err, res, itemsBody) {
			xml2js(itemsBody.replace(/[smt]:/g, ""), function(err_not_used, result) {
				var items = helper.hashGet(result, "Envelope.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder.Items.CalendarItem", []);

				async.map(items,
					function(item, done) {
						var attrs = helper.hashGet(item, "ItemId.$", {});
						var params = {
							id: helper.hashGet(attrs, "Id", ""),
							changeKey: helper.hashGet(attrs, "ChangeKey", "")
						};

						post({
								url: "https://{tld}/EWS/Exchange.asmx".format(config),
								headers: postItem.headers,
								body: postItem.body.format(params)
							},
							function (err, res, itemBody) {
								xml2js(itemBody.replace(/[smt]:/g, ""), function(err_not_used, result) {
									var calendarItem = helper.hashGet(result, "Envelope.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage.Items.CalendarItem", []).pop();
									var calendarBody = helper.hashGet(calendarItem, "Body._", "");

									// special ones
									item["G2M"] = helper.formatUrl(/([^\/]+\.gotomeeting\.com\/join\/[0-9]{9})/.exec(calendarBody));
									item["Webex"] = helper.formatUrl(/([^\/]+\.webex\.com[^<>"]+mtid[^<>"]+)/i.exec(calendarBody));

									// simple values
									item["ItemId"] = params.id;
									item["ChangeKey"] = params.changeKey;
									item["Organizer"] = helper.hashGet(item, "Organizer.Mailbox.Name", "");

									// local dates
									item["StartLocal"] = helper.formatDate(new Date(item["Start"]), config.tz);
									item["EndLocal"] = helper.formatDate(new Date(item["End"]), config.tz);

									// local times
									item["StartTime"] = helper.formatTime(new Date(item["Start"]), config.tz);
									item["EndTime"] = helper.formatTime(new Date(item["End"]), config.tz);

									// long meetings
									item["LongMeeting"] = item["StartLocal"].replace(/[ T].*$/, "") !== config.startString.replace(/[ T].*$/, "");

									// reduce the arrays of 1 to values
									_.each(item, function(value, key, _list) {
										if (helper.isArray(value) && value.length === 1) {
											_list[key] = value = value[0];
										}

										// t/f to 1|0 instead of strings
										if (value === "true") { _list[key] = true; }
										if (value === "false") { _list[key] = false; }
									});

									done(null, item);
								});
							}
						);
					},
					function(err, meetings) {
						var standup = [];
						var prevDay = false;

						_.each(_.sortBy(meetings, "StartLocal"), function(meeting, index/*, key, _list*/) {
							// id prev day values
							if (prevDay && !meeting.LongMeeting) {
								prevDay = false;
								standup.push("-  --------------^ yesterday ^--------------");
							}
							else if (!prevDay && meeting.LongMeeting) {
								prevDay = true;
							}

							var startTime = meeting.StartTime;
							if (meeting.Webex) {
								startTime = "{0}\n      ~> {1}".format(meeting.StartTime, meeting.Webex);
							}
							else if (meeting.G2M) {
								startTime = "{0}\n      ~> {1}".format(meeting.StartTime, meeting.G2M); 
							}

							standup.push("-  {0} @ {1}".format(
								meeting.Subject
									.replace(/\.com/g, "..com")
									.replace(/\[[^\]]*\]/g, "")
									.replace(/\s*\(\s*/g, " (")
									.replace(/\s*\)\s*/g, ") ")
									.trim(),
								startTime
							));
						});

						if (standup.length) {
							callback(
								":spiral_calendar_pad: {0}\n* Meeting{1}:\n{2}"
									.format(
										config.startString.replace(/[ T].*$/, ""),
										standup.length > 1 ? "s" : "",
										standup.join("\n"))
											.replace(/^[*][ ]/mg, "\u2022 ") // bullets
											.replace(/^[-][ -]/mg, "\u2013 ") // en-dash
							);
						}
						else {
							callback("* No Meetings");
						}
					}
				);
			});
		});
};
//module.exports(process.env, function(m) { console.log(m); process.exit() });
