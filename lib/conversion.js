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
	var process_env = _.pick(args, "EWS_TLD", "EWS_AUTH");
	var config = {
		tld: process_env.EWS_TLD || "",
		auth: process_env.EWS_AUTH || ""
	};
	if (!config.tld || !config.auth) {
		throw new Error("Error: env variable '{0}' not set!".format(!config.tld ? "EWS_TLD" : "EWS_AUTH"));
	}

	var postHeaders = {
		"User-Agent": "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
		"Content-Type": "text/xml; charset=utf-8",
		"SOAPAction": '"{soapAction}"'
	};

	var postFolders = {
		soapAction: "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
		body: [
			'<?xml version="1.0" encoding="UTF-8"?>',
			'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">',
			'  <soap:Header>',
			'    <t:RequestServerVersion Version="Exchange2010" />',
			'  </soap:Header>',
			'  <soap:Body>',
			'    <m:FindFolder Traversal="Deep">',
			'      <m:FolderShape>',
			'        <t:BaseShape>IdOnly</t:BaseShape>',
			'        <t:AdditionalProperties>',
			'          <t:FieldURI FieldURI="folder:DisplayName" />',
			'        </t:AdditionalProperties>',
			'      </m:FolderShape>',
			'      <m:IndexedPageFolderView MaxEntriesReturned="50" Offset="0" BasePoint="Beginning" />',
			'      <m:Restriction>',
			'        <t:IsGreaterThan>',
			'          <t:FieldURI FieldURI="folder:TotalCount" />',
			'          <t:FieldURIOrConstant>',
			'            <t:Constant Value="0" />',
			'          </t:FieldURIOrConstant>',
			'        </t:IsGreaterThan>',
			'      </m:Restriction>',
			'      <m:ParentFolderIds>',
			'        <t:DistinguishedFolderId Id="root" />',
			'      </m:ParentFolderIds>',
			'    </m:FindFolder>',
			'  </soap:Body>',
			'</soap:Envelope>'
		].join("").format(config)
	};

	Object.defineProperty(postFolders, "headers", {
		get: function() {
			var headers = _.extend({}, postHeaders);
			headers.SOAPAction = headers.SOAPAction.format(this.soapAction);;
			headers.Authorization = "Basic {auth}".format(config)
			return headers;
		}
	});

	var postEmailIds = {
		soapAction: "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem",
		body: [
			'<?xml version="1.0" encoding="UTF-8"?>',
			'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">',
			'  <soap:Header>',
			'    <t:RequestServerVersion Version="Exchange2010" />',
			'  </soap:Header>',
			'  <soap:Body>',
			'    <m:FindItem Traversal="Shallow">',
			'      <m:ItemShape>',
			'        <t:BaseShape>IdOnly</t:BaseShape>',
			'        <t:AdditionalProperties>',
			'          <t:FieldURI FieldURI="item:Subject" />',
			'          <t:FieldURI FieldURI="item:DateTimeReceived" />',
			'        </t:AdditionalProperties>',
			'      </m:ItemShape>',
			'      <m:IndexedPageItemView MaxEntriesReturned="21" Offset="0" BasePoint="Beginning" />',
			'      <m:Restriction>',
			'        <t:IsEqualTo>',
			'          <t:FieldURI FieldURI="item:Subject" />',
			'          <t:FieldURIOrConstant>',
			'            <t:Constant Value="BETA Sam\'s Club Daily Mobile Metrics" />',
			'         </t:FieldURIOrConstant>',
			'        </t:IsEqualTo>',
			'      </m:Restriction>',
			'      <m:SortOrder>',
			'        <t:FieldOrder Order="Descending">',
			'          <t:FieldURI FieldURI="item:DateTimeReceived" />',
			'        </t:FieldOrder>',
			'      </m:SortOrder>',
			'      <m:ParentFolderIds>',
			'        <t:FolderId Id="{folderId}" ChangeKey="{folderChangeKey}" />',
			'      </m:ParentFolderIds>',
			'    </m:FindItem>',
			'  </soap:Body>',
			'</soap:Envelope>'
		].join("")
	};

	Object.defineProperty(postEmailIds, "headers", {
		get: function() {
			var headers = Object.assign({}, postHeaders);
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
			'        <ns1:BaseShape>Default</ns1:BaseShape>',
			'        <ns1:BodyType>Text</ns1:BodyType>',
			'      </ns2:ItemShape>',
			'      <ns2:ItemIds>',
			'        <ns1:ItemId Id="{id}" ChangeKey="{changeKey}"/>',
			'      </ns2:ItemIds>',
			'      <ns2:ParentFolderIds>',
			'        <ns2:FolderId Id="{folderId}" ChangeKey="{folderChangeKey}" />',
			'      </ns2:ParentFolderIds>',
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
			headers: postFolders.headers,
			body: postFolders.body
		},
		function (err, res, foldersBody) {
			xml2js(foldersBody.replace(/[smt]:/g, ""), function(err_not_used, result) {
				var folders = helper.hashGet(result, "Envelope.Body.FindFolderResponse.ResponseMessages.FindFolderResponseMessage.RootFolder.Folders.Folder", []);
				var folder = _.filter(folders, function(f) {
					return helper.hashGet(f, "DisplayName") === "SamsClub mWeb";
				}).pop();

				if (!folder) {
					throw new Error("WTF?");
				}

				var folderAttrs = helper.hashGet(folder, "FolderId.$", {});
				var folderParams = {
					folderId: helper.hashGet(folderAttrs, "Id", ""),
					folderChangeKey: helper.hashGet(folderAttrs, "ChangeKey", "")
				};

				post({
						url: "https://{tld}/EWS/Exchange.asmx".format(config),
						headers: postEmailIds.headers,
						body: postEmailIds.body.format(folderParams)
					},
					function (err, res, emailIdsBody) {

						xml2js(emailIdsBody.replace(/[smt]:/g, ""), function(err_not_used, result) {

							var emailIds = helper.hashGet(result, "Envelope.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder.Items.Message", []);

							async.mapLimit(emailIds, 7,
								function(emailId, done) {
									var attrs = helper.hashGet(emailId, "ItemId.$", {});
									var params = {
										id: helper.hashGet(attrs, "Id", ""),
										changeKey: helper.hashGet(attrs, "ChangeKey", "")
									};

									post({
											url: "https://{tld}/EWS/Exchange.asmx".format(config),
											headers: postItem.headers,
											body: postItem.body.format(params).format(folderParams)
										},
										function (err, res, itemBody) {
											xml2js(itemBody.replace(/[smt]:/g, ""), function(err_not_used, result) {

												var emailItem = helper.hashGet(result, "Envelope.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage.Items.Message", []);
												var emailBody = helper.hashGet(emailItem, "Body._", "").replace(/[\t\r\n ]+/g, " ").replace(/%/g, "%\n");

												var data = {
													date: "",
													conv: ""
												};

												if (/Conversion Yesterday Vs\. Day before Vs\. 1wk Ago Vs\. 1yr Ago mWeb ([^%]+)[%]/.test(emailBody)) {
													data.conv = +RegExp.$1;
													data.conv = data.conv.toFixed(2);
												}
												if (/Report Date: ([0-9\/]+)/.test(emailBody)) {
													data.date = RegExp.$1;
													data.date = data.date.replace(/\/([0-9])/g, "-0$1").replace(/\-0([0-9]{2})/g, "-$1");
												}

												done(null, data);
											});
										}
									);
								},
								function(err, data) {
									callback(
										":chart_with_upwards_trend:\n* mWeb conversion:\n{0}"
											.format( _.map(data, function(d) {
													if (d.date) {
														return "-  {date}  {conv}%".format(d);
													}
												}).join("\n")
											)
											.replace(/^[*][ ]/mg, "\u2022 ") // bullets
											.replace(/^[-][ -]/mg, "\u2013 ") // en-dash
									);
								}
							);
						})
					}
				);
			});
		});
};
//module.exports(process.env, function(m) { console.log(m); process.exit() });
