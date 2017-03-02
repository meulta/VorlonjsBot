'use strict';

const envx = require("envx");

const restify = require('restify');
const builder = require('botbuilder');
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const expressSession = require('express-session');
const crypto = require('crypto');
const querystring = require('querystring');
const https = require('https');
const request = require('request');

//bot application identity
const MICROSOFT_APP_ID = envx("MICROSOFT_APP_ID");
const MICROSOFT_APP_PASSWORD = envx("MICROSOFT_APP_PASSWORD");

//oauth details
const AZUREAD_APP_ID = envx("AZUREAD_APP_ID");
const AZUREAD_APP_PASSWORD = envx("AZUREAD_APP_PASSWORD");
const AZUREAD_APP_REALM = envx("AZUREAD_APP_REALM");
const AUTHBOT_CALLBACKHOST = envx("AUTHBOT_CALLBACKHOST");

//VaaS variables
const VAAS_ROOT_URL = envx("VAAS_ROOT_URL");

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat bot
console.log('started...')
console.log(MICROSOFT_APP_ID);

var connector = new builder.ChatConnector({
  appId: MICROSOFT_APP_ID,
  appPassword: MICROSOFT_APP_PASSWORD
});

var bot = new builder.UniversalBot(connector);

server.post('/api/messages', connector.listen());

server.get('/', restify.serveStatic({
  'directory': __dirname,
  'default': 'index.html'
}));

//=========================================================
// Bots Dialogs
//=========================================================

bot.dialog('/', [
  (session, args, next) => {
    if (!(session.userData.userName && session.userData.accessToken && session.userData.refreshToken)) {
      session.send("Welcome! This bot will create a Vorlon.js instance for you. Login first!");
      session.beginDialog('signinPrompt');
    } else {
      next();
    }
  },
  (session, results, next) => {
    if (session.userData.userName && session.userData.accessToken && session.userData.refreshToken) {
        handleConversation(session, results, next);
    } else {
      session.endConversation("Goodbye.");
    }
  },
  (session, results) => {
    session.replaceDialog('/');
  }
]);

var handleConversation = function(session, results, next){
    var resp = session.message.text;
    
    switch(resp.toLowerCase()){
        case "create": {
          session.send("I am creating your Vorlon.js instance...");
          vaaSApiCall("create", session.userData.loginData.alias, (err, url, created) =>{
            if(url != null){
              if(created)
                session.send("Done! You can access it using: " + url);
              else
                session.send("You alread have one here: " + url);
            }
            else
              session.send("I did not manage to create your Vorlon.js instance. Try again or contact vorlonjs@microsoft.com");
          });
          break;
        }
        case "delete": {
          session.send("Deleting your instance...");
          vaaSApiCall("remove", session.userData.loginData.alias, (err, url) =>{
            session.send("Done! Feel free to create a new one by saying 'create'");
          });
          break;
        }
        case "logout": {
          session.userData.loginData = null;
          session.userData.userName = null;
          session.userData.accessToken = null;
          session.userData.refreshToken = null;
          session.endConversation("You have logged out. Goodbye.");
          break;
        }
        default: {
          session.send("Welcome " + session.userData.loginData.email + "! You are currently logged in. Say 'create' to create a vorlon.js instance, 'delete' to delete your instance, 'logout' to disconnect!");
        }
    }
}

//=========================================================
// Vorlon.js as a Service API
//=========================================================

var vaaSApiCall = function(command, instanceName, done){

  var rootUrl = VAAS_ROOT_URL;

  var options = {
      method: 'POST',
      url: rootUrl + '/api/instance/' + command,
      body: {
          "serviceName": instanceName
      },
      json: true,
      headers: { 'Content-Type' : 'application/json' }
  };

  request(options, function (err, res, body) {
      if (err) { 
        done(err, null);
      }
      else if (res.statusCode === 201 || res.statusCode === 200) {
        done(null, rootUrl + "/" + instanceName, true);
      }
      else if(res.statusCode === 400) {
        done(null, rootUrl + "/" + instanceName, false);
      }
      else done(null, null, null);
  }); 
}

//=========================================================
// Auth Setup
//=========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(expressSession({ secret: 'keyboard cat', resave: true, saveUninitialized: false }));
server.use(passport.initialize());

server.get('/login', function (req, res, next) {
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', customState: req.query.address, resourceURL: process.env.MICROSOFT_RESOURCE }, function (err, user, info) {
    if (err) {
      return next(err);
    }
    if (!user) {
      return res.redirect('/login');
    }
    req.logIn(user, function (err) {
      if (err) {
        return next(err);
      } else {
        return res.send('Welcome ' + req.user.displayName + ' ' + req.user.email);
      }
    });
  })(req, res, next);
});

server.get('/api/OAuthCallback/',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login' }),
  (req, res) => {
    const address = JSON.parse(req.query.state);
    const magicCode = crypto.randomBytes(4).toString('hex');
    const messageData = { magicCode: magicCode, accessToken: req.user.accessToken, refreshToken: req.user.refreshToken, userId: address.user.id, name: req.user.displayName, email: req.user.upn };
    var authCodePage = '<p style="font-family: Verdana; font-size: 0.9em; text-align: center; margin-top: 40vh">You have to login with your @microsoft.com account.<br/><br/>Go here to logout and login again: <a href="http://account.microsoft.com">account.microsoft.com</a> </p>';
                       
    
    if(messageData.email){
      var mailParts = messageData.email.split('@');
      
      if(mailParts.length === 2 && mailParts[1] === "microsoft.com"){

        messageData.alias = mailParts[0];

        var continueMsg = new builder.Message().address(address).text(JSON.stringify(messageData));
        bot.receive(continueMsg.toMessage());

        authCodePage =  '<p style="font-family: Verdana; font-size: 0.9em; text-align: center; margin-top: 40vh">' + 
                                'Welcome <strong>' + req.user.displayName + '! </strong> <br /><br /> Please copy this number and paste it back to your chat so your authentication can complete: <br /><br /><strong>' + 
                              magicCode +
                            '</strong></p>';
      }
    }

    res.writeHead(200, {
        'Content-Length': Buffer.byteLength(authCodePage),
        'Content-Type': 'text/html'
    });

    res.write(authCodePage);
});

passport.serializeUser(function(user, done) {
  done(null, user);
});

passport.deserializeUser(function(id, done) {
  done(null, id);
});

var realm = AZUREAD_APP_REALM; 

// Use the v1 endpoint (applications configured by manage.windowsazure.com)
// This works against Azure AD
let strategy = {
  redirectUrl: AUTHBOT_CALLBACKHOST +'/api/OAuthCallback',
  realm: realm,
  clientID: AZUREAD_APP_ID,
  clientSecret: AZUREAD_APP_PASSWORD,
  validateIssuer: false,
  oidcIssuer: undefined,
  identityMetadata: 'https://login.microsoftonline.com/' + realm + '/.well-known/openid-configuration',
  skipUserProfile: true,
  responseType: 'code',
  responseMode: 'query',
  passReqToCallback: true
};

passport.use(new OIDCStrategy(strategy,
  (req, iss, sub, profile, accessToken, refreshToken, done) => {
    if (!profile.displayName) {
      return done(new Error("No oid found"), null);
    }
    // asynchronous verification, for effect...
    process.nextTick(() => {
      profile.accessToken = accessToken;
      profile.refreshToken = refreshToken;
      return done(null, profile);
    });
  }
));

//=========================================================
// Bots Dialogs
//=========================================================
function login(session) {
  // Generate signin link
  const address = session.message.address;

  // TODO: Encrypt the address string
  const link = AUTHBOT_CALLBACKHOST + '/login?address=' + querystring.escape(JSON.stringify(address));
  

  var msg = new builder.Message(session) 
    .attachments([ 
        new builder.SigninCard(session) 
            .text("Please click this link to sign in first.") 
            .button("signin", link) 
    ]); 
  session.send(msg);
  builder.Prompts.text(session, "You must first sign into your account.");
}

bot.dialog('signin', [
  (session, results) => {
    console.log('signin callback: ' + results);
    session.endDialog();
  }
]);

bot.dialog('signinPrompt', [
  (session, args) => {
    if (args && args.invalid) {
      // Re-prompt the user to click the link
      builder.Prompts.text(session, "please click the signin link.");
    } else {
      login(session);
    }
  },
  (session, results) => {
    //resuming
    session.userData.loginData = JSON.parse(results.response);
    if (session.userData.loginData && session.userData.loginData.magicCode && session.userData.loginData.accessToken) {
      session.beginDialog('validateCode');
    } else {
      session.replaceDialog('signinPrompt', { invalid: true });
    }
  },
  (session, results) => {
    if (results.response) {
      //code validated
      session.userData.userName = session.userData.loginData.name;
      session.endDialogWithResult({ response: true });
    } else {
      session.endDialogWithResult({ response: false });
    }
  }
]);

bot.dialog('validateCode', [
  (session) => {
    builder.Prompts.text(session, "Please enter the code you received or type 'quit' to end. ");
  },
  (session, results) => {
    const code = results.response;
    if (code === 'quit') {
      session.endDialogWithResult({ response: false });
    } else {
      if (code === session.userData.loginData.magicCode) {
        // Authenticated, save
        session.userData.accessToken = session.userData.loginData.accessToken;
        session.userData.refreshToken = session.userData.loginData.refreshToken;

        session.endDialogWithResult({ response: true });
      } else {
        session.send("hmm... Looks like that was an invalid code. Please try again.");
        session.replaceDialog('validateCode');
      }
    }
  }
]);

function getAccessTokenWithRefreshToken(refreshToken, callback){
  var data = 'grant_type=refresh_token' 
        + '&refresh_token=' + refreshToken
        + '&client_id=' + AZUREAD_APP_ID
        + '&client_secret=' + encodeURIComponent(AZUREAD_APP_PASSWORD) 

  var options = {
      method: 'POST',
      url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      body: data,
      json: true,
      headers: { 'Content-Type' : 'application/x-www-form-urlencoded' }
  };

  request(options, function (err, res, body) {
      if (err) return callback(err, body, res);
      if (parseInt(res.statusCode / 100, 10) !== 2) {
          if (body.error) {
              return callback(new Error(res.statusCode + ': ' + (body.error.message || body.error)), body, res);
          }
          if (!body.access_token) {
              return callback(new Error(res.statusCode + ': refreshToken error'), body, res);
          }
          return callback(null, body, res);
      }
      callback(null, {
          accessToken: body.access_token,
          refreshToken: body.refresh_token
      }, res);
  }); 
}