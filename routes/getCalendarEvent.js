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
        /*
        const date = new Date();
        const today =  date.toISOString();
        const tomorrow = new Date(today);
        tomorrow.setDate(tomorrow.getDate() + 1).toISOString();
        */
        Date.prototype.addHours = function(h){
            this.setHours(this.getHours()+h);
            return this;
        }
        const date = new Date();
        const today =  date.toISOString();
        const tomorrow = date.addHours(24).toISOString();
        try {
            // Get Calendar event
            let response = await client
                .api('me/calendarview?startdatetime='+today+'&enddatetime='+tomorrow)
                .header('Prefer','outlook.timezone="Eastern Standard Time"')
                .select('subject,body,bodyPreview,organizer,attendees,start,end,location')
                .get();
            res.send({
                message: 'Calendar Event fetched successfully!',
                response
            })

             console.log("Calendar response: " +response);

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