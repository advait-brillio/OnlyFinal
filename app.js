const restify = require('restify');
const builder = require('botbuilder');

var authHelper = require('./authHelper');

var outlook = require('node-outlook');

const querystring = require('querystring');

const request = require('request');

var url = require('url');

var http = require('http');

var bodyParser = require('body-parser')

//--------------RESTIFY SERVER-----------------------------------------------------------------------------------------------------
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 8282, function () {
    console.log('%s listening to %s', server.name, server.url);
});

console.log('started...')
var connector = new builder.ChatConnector({
appId: "e9788c79-66c4-4414-8da7-48317ae91b64",
appPassword: "gkNSLEC389[^ewwnaHR90@="
});
var bot = new builder.UniversalBot(connector);
server.post('/api/messages', connector.listen());
//-------------------------------------------------------------------------------------------------------------------


var cookies = []



var emails = []


//======================================================================================================================

//---------------------------------------BOT DIALOGUES--------------------------------------------------------------



//FIRST---------------

bot.dialog('/', [



    (session, args, next) => {



        if (!cookies[3]) {
            session.beginDialog('signinPrompt');

        } else {



            next();



        }



    },



    (session, results, next) => {



        if (cookies[3]) {

            var input = ["email", "calendar", "contacts", "quit", "logout"], options = ['Get Mails', 'Get Events', 'Get Contacts', 'Quit', 'LogOut'];

            // They're logged in

            session.send('Welcome ' + cookies[3] + "." + " How can I help you?");



            // builder.Prompts.text(session,  "* To get the latest Emails, type 'email'.\n\n* To get Calendar Events type 'calendar',\n\n* For Contacts type 'contacts'\n\n* To Quit, type 'quit'. \n\n* To Log Out, type 'logout'. ");



            // create the card based on selection



            sendoptionCard(session, input, options);

            // builder.Prompts.text(session, msg) ;




        } else {



            session.endConversation("Goodbye.");



        }



    },



    (session, results, next) => {

        console.log('results...' + results);

        var resp = results.response;



        if (resp === 'Show my mails') {



            // session.beginDialog('workPrompt');



            session.beginDialog('sendMails');



        } else if (resp === 'Show me the calendar events') {



            session.beginDialog('calendar');



        } else if (resp === 'Show my contacts') {



            session.beginDialog('contacts');



        } else if (resp === 'quit') {



            session.endConversation("Goodbye.");



        } else if (resp === 'logout') {

            cookies = [];

            session.userData.loginData = null;



            session.userData.userName = null;



            session.userData.accessToken = null;



            session.userData.refreshToken = null;



            session.endConversation("You have logged out. Goodbye.");



        } else {



            next();



        }



    },



    (session, results) => {



        session.replaceDialog('/');



    }



]);



//SECOND=========================================



bot.dialog('signinPrompt', [



    (session, args) => {

        login(session);

    },



    (session, results) => {



        if (results.response === `login`) {



            // session.beginDialog('validateCode');

            if (cookies[0]) {



                session.endDialogWithResult({ response: true });



            } else {



                session.send("hmm... Looks like that was an invalid code. Please try again.");



                session.replaceDialog('signinPrompt', { invalid: true });



            }



        } else {

            session.send('Please type "login" again.')

            session.replaceDialog('signinPrompt', { invalid: true });



        }



    },



    (session, results) => {



        if (results.response) {



            session.endDialogWithResult({ response: true });



        } else {



            session.endDialogWithResult({ response: false });



        }



    }



]);



//===============================================   



// bot.dialog('validateCode', [



//     (session) => {



//         builder.Prompts.text(session, "Please type 'ok' to access outlook. ");



//     },



//     (session, results) => {



//         // const code = results.response;

//         const code = cookies[0]



//         console.log(code)



//         if (code == 'quit') {



//             session.endDialogWithResult({ response: false });



//         } else {



//             if (code == cookies[0]) {



//                 session.endDialogWithResult({ response: true });



//             } else {



//                 session.send("hmm... Looks like You are logged out. Please try again.");



//                 session.replaceDialog('validateCode');



//             }



//         }



//     }



// ]);



bot.dialog('sendMails', [



    (session, args, next) => {



        mail(session);

        session.send("Okay these are your latest recieved mails.");



    }

    // ,



    // (session, results) => {



    //     session.replaceDialog('/');



    // }



]);



bot.dialog('calendar', [



    (session, args) => {



        calendar(session);

        session.send(" Here's your outlook calendar events.");



    }

    // ,



    // (session, results) => {



    //     session.replaceDialog('/');



    // }



]);



bot.dialog('contacts', [



    (session, args) => {



        contacts(session);



    }

    // ,



    // (session, results) => {



    //     session.replaceDialog('/');



    // }



]);






//==============================FUNCTIONS===========================================================================================



//-------------------------------------------------------------------------------------------------------------------

//when signin button clicked in the bot ==> localhost 3000==>homepage

server.get("/", function home(response, request, next) {

    console.log('Request handler \'home\' was called.');

    response.writeHead(200, { 'Content-Type': 'text/html' });

    response.end();

    next();

});



//THIRD =====



function login(session) {

    var link = authHelper.getAuthUrl()

    var msg = new builder.Message(session)

        .attachments([

            new builder.SigninCard(session)

                .text("Welcome! Please click on the below link to signin to outlook.")

                .button("signin", link)

        ]);

    session.send(msg);

    builder.Prompts.text(session, "Please type 'login' to continue.");

}




bot.dialog('signin', [



    (session, results) => {



        console.log('signin callback: ' + results);



        session.endDialog();



    }



]);





server.get("/authorize", function authorize(response, request, next) {



    console.log('Request handler \'authorize\' was called.');



    // console.log(response._url.query);
    console.log("response " + response);


    // The authorization code is passed as a query parameter

    var url_parts = response._url.query;



    var code = url_parts.replace("code=", "")
    var code_arr = code.split("&session_state=")
    code_arr[0]


    //console.log(url_parts)

    // console.log("url part:"+ url_parts);

    console.log("Code " + code_arr[0])



    // console.log('Code: ' + code);



    authHelper.getTokenFromCode(code_arr[0], tokenReceived, response);



});




function tokenReceived(response, error, token) {
    if (error) {
        console.log('Access token error: ', error.message);
    } else {

        getUserEmail(token.token.access_token, function (error, email) {



            if (error) {



                console.log('getUserEmail returned an error: ' + error);



            } else if (email) {



                cookies = [token.token.access_token, token.token.refresh_token, token.token.expires_at.getTime(), email];



                // response.writeHead(302, { 'Location': 'http://localhost:8080/code' });



                // response.end();



            }



        });



    }



}




function getUserEmail(token, callback) {



    // Set the API endpoint to use the v2.0 endpoint



    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');




    // Set up oData parameters



    var queryParams = {



        '$select': 'DisplayName, EmailAddress',



    };




    outlook.base.getUser({ token: token, odataParams: queryParams }, function (error, user) {



        if (error) {



            callback(error, null);



        } else {



            callback(null, user.EmailAddress);



        }



    });



}



function getValueFromCookie(valueName, cookie) {



    if (cookie.indexOf(valueName) !== -1) {



        var start = cookie.indexOf(valueName) + valueName.length + 1;



        var end = cookie.indexOf(';', start);



        end = end === -1 ? cookie.length : end;



        return cookie.substring(start, end);



    }



}




function getAccessToken(request, response, callback) {



    var expiration = new Date(parseFloat(cookies[2]));




    if (expiration <= new Date()) {



        // refresh token



        console.log('TOKEN EXPIRED, REFRESHING');



        var refresh_token = cookies[1];



        authHelper.refreshAccessToken(refresh_token, function (error, newToken) {



            if (error) {



                callback(error, null);



            } else if (newToken) {



                cookies = [newToken.token.access_token, newToken.token.refresh_token, newToken.token.expires_at.getTime()];



                callback(null, newToken.token.access_token);



            }



        });



    } else {



        // Return cached token



        var access_token = cookies[0];



        callback(null, access_token);



    }



}




server.get("/code", function code(response, request) {



    getAccessToken(request, response, function (error, token) {



        console.log('Token found in cookie: ', token);



        var email = cookies[3]



        console.log('Email found in cookie: ', email);



        if (token) {



            response.writeHead(200, { 'Content-Type': 'text/html' });



            response.write('<div align="center"><h1>Welcome  ' + email + '</h1></div>');



            response.write("<div align='center'><h3>Please go back to the bot. You'll be able to access your Outlook Account now. </h3></div>");

            response.end();



        } else {



            response.writeHead(200, { 'Content-Type': 'text/html' });



            response.write('<p> No token found in cookie!</p>');



            response.end();



        }



    });



});




function mail(session, response, request) {



    getAccessToken(request, response, function (error, token) {



        console.log('Token found in cookie: ', token);



        var email = cookies[3]



        console.log('Email found in cookie: ', email);



        if (token) {



            var queryParams = {



                '$select': 'Subject,ReceivedDateTime,From,IsRead, BodyPreview',



                '$orderby': 'ReceivedDateTime desc',



                '$top': 10



            };



            // Set the API endpoint to use the v2.0 endpoint



            outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');



            // Set the anchor mailbox to the user's SMTP address



            outlook.base.setAnchorMailbox(email);



            outlook.mail.getMessages({ token: token, folderId: 'inbox', odataParams: queryParams },



                function (error, result) {



                    if (error) {



                        console.log('getMessages returned an error: ' + error);



                    }



                    else if (result) {



                        // console.log('getMessages returned ' + result.value.length + ' messages.');



                        // var i = 0;

                        // session.send("Okay Iam in your Inbox now.");

                        // result.value.forEach(function (message) {



                        //     console.log(' Subject: ' + message.Subject);



                        //     var from = message.From ? message.From.EmailAddress.Name : 'NONE';



                        //     emails[i] = "From :" + from + " Subject :" + message.Subject + " on " + message.ReceivedDateTime.toString();



                        //     session.send(emails[i]);



                        //     i++;

                        console.log('getMessages returned ' + result.value.length + ' messages.');

                        var i = 0, sub = [], tim = [], fromadd = [], body = [];

                        result.value.forEach(function (message) {

                            console.log(' Subject: ' + message.Subject);

                            console.log('message body:' + message.BodyPreview);

                            var from = message.From ? message.From.EmailAddress.Name : 'NONE';

                            sub[i] = message.Subject

                            tim[i] = message.ReceivedDateTime.toString();

                            fromadd[i] = from;

                            body[i] = message.BodyPreview;

                            // emails[i]="From :"+from+" Subject :"+message.Subject+" on "+message.ReceivedDateTime.toString();

                            // session.send(emails[i]);

                            i++;

                        });

                        sendCardMail(session, fromadd, tim, sub, body);

                        session.endDialogWithResult({

                            resumed: builder.ResumeReason.notCompleted

                        });



                        // session.replaceDialog('/');

                    }

                });



        } else {



            console.log('No token found in cookie!');



        }



    });



}




function calendar(session, response, request) {



    var token = cookies[0];



    console.log('Token found in cookie: ', token);



    var email = cookies[3];



    console.log('Email found in cookie: ', email);



    if (token) {



        var queryParams = {



            '$select': 'Subject,Start,End,Attendees, BodyPreview',



            '$orderby': 'Start/DateTime desc',



            '$top': 10



        };



        // Set the API endpoint to use the v2.0 endpoint



        outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');



        // Set the anchor mailbox to the user's SMTP address



        outlook.base.setAnchorMailbox(email);



        // Set the preferred time zone.



        // The API will return event date/times in this time zone.



        outlook.base.setPreferredTimeZone('Eastern Standard Time');



        outlook.calendar.getEvents({ token: token, odataParams: queryParams },



            function (error, result) {



                if (error) {



                    console.log('getEvents returned an error: ' + error);



                } else if (result) {

                    console.log('getEvents returned ' + result.value.length + ' events.');



                    var i = 0, sub = [], tim = [], attend = [], body = [];

                    result.value.forEach(function (event) {

                        console.log(' Subject: ' + event.Subject);

                        console.log(' Starting Time: ' + event.Start.DateTime.toString());

                        console.log(' Ending Time: ' + event.End.DateTime.toString());

                        console.log(' Attendees: ' + buildAttendeeString(event.Attendees));

                        console.log(' Event dump: ' + JSON.stringify(event));

                        body[i] = event.BodyPreview

                        sub[i] = event.Subject

                        tim[i] = event.Start.DateTime.toString() + ' to ' + event.End.DateTime.toString()

                        attend[i] = buildAttendeeString(event.Attendees);

                        i++;

                    });

                    sendCardCalendar(session, sub, tim, attend, body);

                    // session.replaceDialog('/');

                    session.endDialogWithResult({

                        resumed: builder.ResumeReason.notCompleted

                    });



                }

            });

    }

    else {



        console.log('No token found in cookie!');



    }



}




function contacts(session, request, response) {



    var token = cookies[0]



    console.log('Token found in cookie: ', token);



    var email = cookies[3]



    console.log('Email found in cookie: ', email);



    if (token) {



        var queryParams = {



            '$select': 'GivenName,Surname,EmailAddresses',



            '$orderby': 'GivenName asc',



            '$top': 10



        };



        // Set the API endpoint to use the v2.0 endpoint



        outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');



        // Set the anchor mailbox to the user's SMTP address



        outlook.base.setAnchorMailbox(email);



        outlook.contacts.getContacts({ token: token, odataParams: queryParams },



            function (error, result) {



                if (error) {



                    console.log('getContacts returned an error: ' + error);

                    session.replaceDialog('/');



                } else if (result) {



                    console.log('getContacts returned ' + result.value.length + ' contacts.');

                    var i = 0, firstName = [], lastName = [], mail = [];

                    result.value.forEach(function (contact) {



                        var email = contact.EmailAddresses[0] ? contact.EmailAddresses[0].Address : 'NONE';



                        console.log('First name: ' + contact.GivenName);



                        console.log('Last name: ' + contact.Surname);



                        console.log('Email: ' + email);

                        firstName[i] = contact.GivenName;

                        lastName[i] = contact.Surname;

                        mail[i] = email;

                        i++;

                    });

                    session.send("You have " + result.value.length + " outlook contacts. Here they are...");

                    sendcardContacts(session, firstName, lastName, mail);

                    // session.send('First name: ' + contact.GivenName + ' Last name: ' + contact.Surname + ' Email: ' + email);



                    session.endDialogWithResult({

                        resumed: builder.ResumeReason.notCompleted

                    });





                }

            });



    }



    else {



        console.log('No token found in cookie!');

        session.replaceDialog('/');

    }



}





function buildAttendeeString(attendees) {



    var attendeeString = 'wut';



    if (attendees) {



        attendeeString = '';




        attendees.forEach(function (attendee) {



            attendeeString += attendee.EmailAddress.Name + "<br>";




            // attendeeString += ' Type:' + attendee.Type;



            // attendeeString += ' Response:' + attendee.Status.Response;



            // attendeeString += ' Respond time:' + attendee.Status.Time;



        });



    }




    return attendeeString;



}



///===========================card attachment==================================

function sendCardMail(session, fromadd, tim, sub, body) {

    var attachments = [];



    var msg = new builder.Message(session);

    msg.attachmentLayout(builder.AttachmentLayout.carousel);

    var i = 0

    while (sub[i] != null) {



        var card = {

            'contentType': 'application/vnd.microsoft.card.adaptive',

            'content': {

                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',

                'type': 'AdaptiveCard',

                'version': '1.0',

                'body': [

                    {

                        "type": "TextBlock",

                        "text": "From: " + fromadd[i],

                        "size": "medium",

                        "weight": "bolder"

                    },

                    {

                        "type": "TextBlock",

                        "text": "Recieved at: " + tim[i],

                        "wrap": true

                    },

                    {

                        "type": "TextBlock",

                        "text": "Subject: ",

                        "size": "medium",

                        "weight": "bolder",



                    },

                    {

                        "type": "TextBlock",

                        "text": sub[i],

                        "size": "medium",



                        "wrap": true

                    },



                ],

                "actions": [

                    {

                        "type": "Action.ShowCard",

                        "title": "View...",

                        "card": {

                            "type": "AdaptiveCard",

                            "body": [

                                {

                                    "type": "TextBlock",

                                    "text": "Content: ",

                                    "size": "medium",

                                    "wrap": true,

                                    "weight": "bolder"

                                },

                                {

                                    "type": "TextBlock",

                                    "width": "stretch",

                                    "height": "stretch",

                                    "text": body[i],

                                    "size": "medium",

                                    "wrap": true

                                }

                            ]

                        }

                    }

                ]

            }

        }



        attachments.push(card);

        i++;

    }

    msg.attachments(attachments)

    session.send(msg);



}



function sendCardCalendar(session, sub, tim, attend, body) {

    var attachments = [];

    var msg = new builder.Message(session);

    msg.attachmentLayout(builder.AttachmentLayout.carousel);

    var i = 0

    while (sub[i] != null) {

        var card = {

            'contentType': 'application/vnd.microsoft.card.adaptive',

            'content': {

                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',

                'type': 'AdaptiveCard',

                'version': '1.0',

                'body': [

                    {

                        "type": "TextBlock",

                        "text": "Subject: " + sub[i],

                        "size": "medium",

                        "weight": "bolder",

                        "wrap": true

                    },

                    {

                        "type": "TextBlock",

                        "text": "From: " + tim[i],

                        "wrap": true

                    }




                ],

                "actions": [

                    {

                        "type": "Action.ShowCard",

                        "title": "Event Details",

                        "card": {

                            "type": "AdaptiveCard",

                            "body": [

                                {

                                    "type": "TextBlock",

                                    "text": "Event: ",

                                    "size": "medium",

                                    "weight": "bolder",

                                    "wrap": true,



                                },

                                {

                                    "type": "TextBlock",

                                    "text": body[i],

                                    "size": "medium",

                                    "wrap": true,



                                }

                            ]

                        }

                    },

                    {

                        "type": "Action.ShowCard",

                        "title": "View Attendies",

                        "card": {

                            "type": "AdaptiveCard",

                            "body": [

                                {

                                    "type": "TextBlock",

                                    "text": "Attendies",

                                    "size": "medium",

                                    "weight": "bolder",

                                    "wrap": true

                                },

                                {

                                    "type": "TextBlock",

                                    "text": attend[i],

                                    "size": "medium",



                                    "wrap": true

                                }

                            ]

                        }

                    }

                ]

            }

        }

        // var card = new builder.HeroCard(session)

        //     .title(sub[i])

        //     .subtitle(tim[i])

        //     .text(attend[i])

        //body[i]

        attachments.push(card);

        i++;

    }

    msg.attachments(attachments)

    session.send(msg);

}



function sendcardContacts(session, firstName, lastName, mail) {

    var attachments = [];

    var msg = new builder.Message(session);

    msg.attachmentLayout(builder.AttachmentLayout.carousel);

    var i = 0

    while (firstName[i] != null) {

        var card = {

            'contentType': 'application/vnd.microsoft.card.adaptive',

            'content': {

                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',

                'type': 'AdaptiveCard',

                'version': '1.0',

                'body': [

                    {

                        "type": "TextBlock",

                        "text": "Name: " + firstName[i] + " " + lastName[i],

                        "size": "medium",

                        "weight": "bolder",

                        "wrap": true

                    },

                    {

                        "type": "TextBlock",

                        "text": "Email ID: " + mail[i],

                        "weight": "bolder",

                        "wrap": true

                    }



                ],

            }

        }

        // var card = new builder.HeroCard(session)

        //     .title(sub[i])

        //     .subtitle(tim[i])

        //     .text(attend[i])

        //body[i]

        attachments.push(card);

        i++;

    }

    msg.attachments(attachments)

    session.send(msg);

}



function sendoptionCard(session, input, options) {

    // console.log('im in')

    // var attachments = [];

    // var msg = new builder.Message(session);

    // msg.attachmentLayout(builder.AttachmentLayout.carousel);



    var i = 0

    while (input[i] != null) {



        // var card = new builder.HeroCard(session)

        //     .buttons([

        //         builder.CardAction.postBack(session, input[i], options[i])



        //     ])

        var msg = new builder.Message(session)



            .suggestedActions(

            builder.SuggestedActions.create(

                session, [

                    builder.CardAction.imBack(session, "Show my mails", "Get my mails"),

                    builder.CardAction.imBack(session, "Show me the calendar events", "View Events"),

                    builder.CardAction.imBack(session, "Show my contacts", "View Contacts"),

                    builder.CardAction.imBack(session, "logout", "Logout")

                ]

            ));

        //   session.send(msg);



        // attachments.push(card);

        i++;

    }

    // msg.attachments(attachments)



    builder.Prompts.text(session, msg);



}

