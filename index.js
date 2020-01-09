//
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
//
const crypto = require('crypto');
const sharedSecret = "CrDKeWlKabngHJbqbSMFu/AIniR61xb83wUn+LrqHO8="; // e.g. "+ZaRRMC8+mpnfGaGsBOmkIFt98bttL5YQRq3p2tXgcE="
const bufSecret = Buffer(sharedSecret, "base64");

var request = require("request");
var http = require('http');
var fs = require('fs');
var PORT = process.env.port || process.env.PORT || 8080;

const sendPayload = function(){
	var options = { method: 'POST',
	  url: 'https://outlook.office.com/webhook/1d8f5e75-6e0e-4e8e-9432-eb48d2550a02@221ba5e5-7011-40b1-bf29-b717749b7124/IncomingWebhook/19dcc0af4fd141fbb3e2d202af1e3cee/bedd77d6-353d-4858-8ffc-458d8733d523',
	  headers: 
	   { 'Content-Type': 'application/json' },
	  body: 
	   { title: 'Fill Details to create a Ticket',
		 text: '* fields are mandatory.',
		 '@type': 'MessageCard',
		 '@context': 'http://schema.org/extensions',
		 themeColor: '0076D7',
		 summary: 'Larry Bryant created a new task',
		 sections: 
		  [ { activityTitle: '![TestImage](https://47a92947.ngrok.io/Content/Images/default.png)Larry Bryant created a new task',
			  activitySubtitle: 'On Project Tango',
			  activityImage: 'https://teamsnodesample.azurewebsites.net/static/img/image5.png',
			  facts: 
			   [ { name: 'Assigned to', value: 'Unassigned' },
				 { name: 'Due date',
				   value: 'Mon May 01 2017 17:07:18 GMT-0700 (Pacific Daylight Time)' },
				 { name: 'Status', value: 'Not started' } ],
			  markdown: true } ],
		 potentialAction: 
		  [ { '@type': 'ActionCard',
			  name: 'Add a comment',
			  inputs: 
			   [ { '@type': 'TextInput',
				   id: 'comment',
				   isMultiline: false,
				   title: 'Add a comment here for this task' } ],
			  actions: 
			   [ { '@type': 'HttpPOST',
				   name: 'Add comment',
				   target: 'http://...' } ] },
			{ '@type': 'ActionCard',
			  name: 'Set due date',
			  inputs: 
			   [ { '@type': 'DateInput',
				   id: 'dueDate',
				   title: 'Enter a due date for this task' } ],
			  actions: 
			   [ { '@type': 'HttpPOST',
				   name: 'Save',
				   target: 'https://www.google.com' } ] },
			{ '@type': 'ActionCard',
			  name: 'Change status',
			  inputs: 
			   [ { '@type': 'MultichoiceInput',
				   id: 'list',
				   title: 'Select a status',
				   isMultiSelect: 'false',
				   choices: 
					[ { display: 'In Progress', value: '1' },
					  { display: 'Active', value: '2' },
					  { display: 'Closed', value: '3' } ] } ],
			  actions: [ { '@type': 'HttpPOST', name: 'Save', target: 'http://...' } ] } ] },
	  json: true };
	
	request(options, function (error, response, body) {
	  if (error) throw new Error(error);
	
	  console.log(body);
	});
	
}

fs.readFile('./index.html', function (err, html) {
    if (err) {
        throw err; 
    }       
    http.createServer(function(req, response) { 
		var payload = '';
		// Process the req
		req.on('data', function (data) {
			payload += data;
		});
		
		// Respond to the req
		req.on('end', function() {
			sendPayload();
			try {
				// Retrieve authorization HMAC information
				var auth = this.headers['authorization'];
				// Calculate HMAC on the message we've received using the shared secret			
				var msgBuf = Buffer.from(payload, 'utf8');
				var msgHash = "HMAC " + crypto.createHmac('sha256', bufSecret).update(msgBuf).digest("base64");
				// console.log("Computed HMAC: " + msgHash);
				// console.log("Received HMAC: " + auth);
				
				response.writeHead(200);
				if (msgHash === auth) {
					var receivedMsg = JSON.parse(payload);
					var responseMsg = '{ "type": "message", "text": "You typed: ' + receivedMsg.text + '" }';	
				} else {
					var responseMsg = '{ "type": "message", "text": "Error: message sender cannot be authenticated." }';
				}
				console.log(responseMsg);
				return response.write(html);
			}
			catch (err) {
				return response.write(html);
				// response.writeHead(400);
				// return response.end("Error: " + err + "\n" + err.stack);
			}
		});
			
	}).listen(PORT);
});

console.log('Listening on port %s', PORT);
