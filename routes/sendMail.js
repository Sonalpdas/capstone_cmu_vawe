/**
 * http-server -a localhost -p 8081
 * @type {createApplication}
 */

var express = require('express');
var router = express.Router();
var MicrosoftGraph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

/* GET /authorize. */
router.post('/', async function (req, res, next) {
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
        const receiverEmailArray = req.body.receiverEmail;
        //here we assume json request pass an array of string: "receiverEmail": ["mia@gmail.com", "sonal@gmail.com"]
        //Assume pre-processing is done in alexa.
        //Email de-duplication needed?

        try {
            for (const receiverEmail of receiverEmailArray) {
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

                // Send Mail
                let response = await client
                    .api('/me/sendMail')
                    .post(messageRequest)
                //response is empty "" after success sending.
            }
            res.send({
                message: 'Email sent successfully!'
            })
        } catch
            (error) {
            console.log(error, '----->')
        }

    } else {
        res.send({
            error: "No access token"
        })
    }
});

module.exports = router;