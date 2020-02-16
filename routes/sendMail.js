var express = require('express');
var router = express.Router();
var MicrosoftGraph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

/* GET /authorize. */
router.post('/', async function(req, res, next) {
    // Get auth code
    let token;
    const tokenAvailable = req.headers.authorization ||
    req.headers['x-access-token'];

    if (req.headers.authorization) {
      [, token] = req.headers.authorization.split(' ');
    } else {
      token = tokenAvailable;
    }

    if (token) {
        // Create a Graph client
        var client = MicrosoftGraph.Client.init({
            authProvider: (done) => {
                // Just return the token
                done(null, token);
            }
        });

        const subject = req.body.subject;
        const messageBody = req.body.messageBody;
        const receiverEmail = req.body.receiverEmail;

        const messageRequest = {
            message: {
                subject,
                body: {
                    contentType: "Text",
                    content: messageBody
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: receiverEmail
                        }
                    }
                ]
            },
            saveToSentItems: "true"
        };

       // console.log(messageRequest, '====>')
      try {
        // Send Mail
        let response = await client
        .api('/me/sendMail')
        .post(messageRequest)
        res.send({
            message: 'Email sent successfully!',
            response
        })
      } catch (error) {
          console.log(error, '----->')
      }
        
    } else {
        res.send({
            error: "No access token"
        })
    }
  });

module.exports = router;