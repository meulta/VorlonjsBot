var restify = require('restify');
var builder = require('botbuilder');

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat bot
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});
var bot = new builder.UniversalBot(connector);
server.post('/api/messages', connector.listen());

//=========================================================
// Bots Dialogs
//=========================================================

bot.dialog('/', function (session) {
    switch(session.message.text.toLowerCase()){
        case "create":
            session.send("Here is your Vorlon.js instance: ...");
            break;
        case "delete":
            session.send("I deleted your Vorlon.js instance: ...");
            break;
        case "reset":
            session.send("Here is your NEW Vorlon.js instance: ...");
            break;
        default:
            session.send("I don't know. You can say *create* *delete* or *reset* to create, reset or delete your Vorlon.js instance!");
            break;
    }
});