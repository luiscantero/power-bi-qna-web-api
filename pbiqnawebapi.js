/// <reference path="node_modules/@types/node/index.d.ts" />

"use strict";

// Config.
const CONFIG = require('./config');
const LOAD_DELAY = CONFIG.LOAD_DELAY,
    RENDER_DELAY = CONFIG.RENDER_DELAY;

const express = require('express'),
    fs = require('fs'),
    app = express(),
    puppeteer = require('puppeteer');

const https = require('https');
var browser, page;

function getPbiQna(query) {
    return new Promise(async (resolve, reject) => {
        // Note: Currently we use a delay to allow for the result to load before making the screenshot.
        // TODO: A better approach would be to react to the qna's "rendered" event, but it currently doesn't work,
        // it needs to be fixed by the Power BI team.
        // https://stackoverflow.com/questions/49432205/power-bi-embedded-qa-events-do-not-fire-is-it-a-bug

        console.log(`Rendering query (${RENDER_DELAY} ms) ...`);
        await page.evaluate(query => {
            return setQuestion(query);
        }, query);
        await page.waitFor(RENDER_DELAY); // ms
        //await page.waitForNavigation({ 'waitUntil': 'networkidle2' }); // 0/2 both hang.

        console.log("Making screenshot ...");
        const container = await page.$('#qnaContainer iframe'); // Select frame only.
        var png = await container.screenshot();

        console.log("Done");

        resolve(png);
    });
}

// GET: api/pbiqna/query
app.get('/api/pbiqna/:query', async (req, res) => {
    var query = req.params.query;
    console.log("Query: " + query);

    try {
        var png = await getPbiQna(query);
        writeResponse(res, 200, png);
    } catch (e) {
        writeResponse(res, 400, e);
    }

    // Red dot data URL for testing (works).
    // writeResponse(res, 200, 'data:image/jpeg;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==');
    // Apparently the bot emulator has a limit for data URLs.
    // https://stackoverflow.com/questions/37915171/how-do-i-display-images-in-microsoft-bot-framework-with-only-the-base64-encoded
    // An image can be sent to the bot client directly as a hosted URL,
    // but sending as data URL has the advantage that a "Loading" message can be sent to the user before the
    // Power BI response is ready to be sent.
});

var server = app.listen(8081, async () => {
    // Launch puppeteer.
    console.log('Launching puppeteer ...');
    browser = await puppeteer.launch();
    page = await browser.newPage();

    console.log('Setting config ...');
    var html = await readFile('pbi.html');
    html = html.replace('<EMBED_URL>', CONFIG.EMBED_URL);
    html = html.replace('<DATASET>', CONFIG.DATASET);

    process.stdout.write('Getting token ... ');
    // var token = CONFIG.TOKEN; // Get token from config.
    var token = await getSampleToken(); // Get token from public sample.
    console.log(`${token.substring(0, 30)}...`);
    html = html.replace("<TOKEN>", token);

    html = html.replace("<INITIAL_QUERY>", CONFIG.INITIAL_QUERY);

    console.log(`Setting content (${LOAD_DELAY} ms) ...`);
    await page.setContent(html);
    await page.waitFor(LOAD_DELAY); // ms

    console.log(`Server running at http://${server.address().address}:${server.address().port}/`);
}).on('error', (err) => {
    console.log(err);
});

async function getSampleToken() {
    var sampleQna = await httpGet('https://powerbilivedemobe.azurewebsites.net/api/Datasets/SampleQna');
    return JSON.parse(sampleQna.toString('utf-8')).embedToken.token;
}

function writeResponse(res, status, output) {
    res.writeHead(status, {
        'Access-Control-Allow-Origin': '*', // Enable CORS.
        'Content-Type': 'image/png'
    });
    res.end(output);
}

function httpGet(url) {
    return new Promise((resolve, reject) => {
        https.get(url, (res) => {
            // TODO: Test for success status code.
            //console.log(`Response code: ${res.statusCode}`);

            res.on('data', (chunk) => {
                resolve(chunk);
            });
        }).on('error', (e) => {
            reject(`Error: ${e.message}`);
        });
    });
}

function readFile(file) {
    return new Promise((resolve, reject) => {
        fs.readFile(file, 'utf-8', (err, data) => {
            if (err) return reject(err);
            resolve(data);
        });
    });
}


// --------------------------------------------------------------------------------


// Set configuration to test approach #1 and #2.

// Read embed application token from textbox
// var txtAccessToken = 'H4sIAAAAAAAEAB2Wx86EXJJE3-Xf0lLhTUu9AIqL954dUHjvzWjefb6edYZycRSZEf_zj5U-_ZT-_vn3P5hBjcsm72gxf9zv866-HJi8Bn3tHntiFCso9bBtIGqFPc9z41UyQwc4BSlXV1ixXLLU2LkoMmi_Nru3JoOsXpWAotoKukwJJeOPn6G0WFTftMetOn5G9-ZRe3oIit5rpKHyliDCSqWet9by1pNgKSDk97JGzNmPqaYJn6zsfEgYnDzutEKxJmlLu1vksW9J1rFPpxeghxzF3hYeO17sNRguWgg8hDOO_iF-u8EHaHdogUHEdfLYB1R4or6qvfgailSiHYmd9e0qGOT0Y8AinKtXVFfQmsjWtsZeAWEGhcz3Q8ITmWvhVtn4iGB3cBHLditUGqyRPwcZ8CUJ7y-1Qcy8GYjA6kO6dLHqj7_xk9vzU_eWTMpjU0lP4OK2bxfF-npXdawoWM1bg2PavXJmYJfWDFHSxpaupJVulWtQFgp9eIJbGyOsKrqLAv-jGfceS4r6dKbi8eeGYwSHVTAGOrOFlsJGIb4DjI9FT85n5vkISrYOaBylB_2t80OrtKE691Kyv9qnqBYTsUqOKVgeEQ1YARAcNzbj3nWWzQjwZvT8Xl7A43oZIjdSiBWjZlf3EJE2Ciz_UtkVps5zUxhXFHB6xemc4Ym1lMog6kNMNWNNVWodq43KNgG9GYZIP2oaF3RJ-EOO0BfSq0wWMyC8GemmsirYcAL_TqnBJW9LyoFXYJxiHhyTpo2vcbNANJVYHOIP5xpeWspcS6ywTrTY5NXrQPJZJWmN7dlJP8xxnGtarA1zbJMuo6zskdaO227UVcdh1uc0LBUTlydJGVwe9ByLsTyMfe6YIa7acHdbhqp1jsNjvh5F1KfHbCdK-64OFS0H_lp2G-VrfWmTEb4pk0fxAINQG9v3jnCJdVBf7wqqGfxfOuDvWuak4O0L3gDi4sS0NqttLHdMR1uCvfaO_qBI2qNU-kH8cZr1lyhp2T9G5srzmH3rrvWxUQemqg5dZjTDDyakx-LnC7QA6MxL4Pe99jWM6mSnF7cZ9CVdQphENmQ2HNNguLhyXlEjGm7knAXXOmPvyn87A1QqyEoal4Z9OIcBnwJq7_Ul-45p02zv9TPN7GsvUmk-7gXRr2tlnT84FS88JBdtZ7Cd6DoJBTu5k-1k0fHNdYW8UdGdD6RluMtd6GSKRSV6eFG7yv3e93m9CbzfTsQEJdRaOFMIgjKP6BPt2Lt8Wf5T8rnX6LGsfYd1gK71M03dr8grPfH_pu-nFM6tcvAW8VEHbURGOqw2yeUJzsMa_4wTQioyASEoeL7mvvHhd66O7e-JjCFCghCB5_iymJHaMWl90_ODaznTSGheC-l4KvjHOOr43dondaV52ycPKfYTunlc_W1L5Zx-FcD7F9VOfOrfnIAiP0Z_ByfJjHeGF1o1UJ8HcShd2TNs9-cXfsoOLoXkgpspWhWBrWkHnNCC-58J8tZcgFn92E8tNFKSR1mPT2jmx3-zLxViSyPr0fEkuyYPLjXcXt3T6DJ3QL5rujMbIgP8HrCTbyjxwKTwncM7-tPSInMkmm7hzuHoEHv3UR8bV9oSZRwGuRqRFL2HYZ3zjjc0a1uNLEDDmfzRrE9GClS4yK2WHS48BtUCSmaLX3ZbvJFKg-APFut0eAcEVNkAv-xQCmPNd5ip72onP82m7cpCID8kUuiyortj0uazeaSIaoLMqdISnaS9rw5Wpo1RBBPK0nXU7ogT9YeK-4lnMBxkEQCUsCNy_lZVKvz-0KF19nzjiMN4jZSbDR3tdnx7Q6fu-4f9_IIkWlVx8uJF4hSm1wYX3yQY19l1vXIqSsH-5ftTds0ubi_08njCufvFKvVwahyHH0e0r5Gy3DkyN6khVJP2hfZ9-fLFlPPM34E39a80FbJHd_UTrPSsSJ4nYU4QbWqB-u2ifJX3tNg5Si4L8XxIyhfApmw3HaiWJ7KQNbmaYQ3Xp6F7JY-4HQp9p0xmliRY2YRCPwoovLlC9G6AS95OyvzZWCbtUi2ktkHrPesYp5T3VZ_5bsXHwhPsUDTtpFNEEtfMjbyqn1v5SChyd5LAPzRiTtpZ6mn4Enlq7bBaMNwLhNw___qHX595n9Ti-asD_DWJQWKmu3DJiqe1B4d4Z2evm1-CM6aTUO8emE_wWwESzHgIv-TZ7Mf6mgK2lIM1mxljQc5oudgKFYWM0iL5gvx03e7TYthZYegpeCLFaggTGF7er-BhWwaAfTeKZVl_6IQoUdiDLcR5BPOX1dHoavMGHF2bTa5faaEs8q87rCeDf7unPB01csS_EHTSfPXsobtVqUJMUUZqDNIZ0DDa6FdqPuNpYtVpEwSjWnGLBmA-mJhMNdzb41vjU_dBzy0csORBkkjZ4k8WhVI5o4vkYLixN8FfhflRAGIyg1O_JzsUtXciwnb4StDPqMVrN0WmX6C-9k-VG-qlmjx0gDqB9PrPf_6L-ZnrYpWDP8qSdQVA9RZDe9C2oEoI4daN_X-V21Rjuh9r8ScTdDlyeFXAERtJ7fwDIqkvztUfmVEMfA4rf_OIWTr50SmM9dnm1NIXCP41H73o9hEeaVoJXMv3Yx4dt9V1J5-7aXFLFmBZpEBs8sDlgoFT-w9B_kyAGly5PH18fR-WWt9aWYz19yHGmL34xvKiD5OkGKXSiQSZ5nvP2YlSyU9Rm8cx04iUzMjo3LDQ-e6anUwYEEA3HazzDcJcUidTxDCvWS3uTP5NRw3yyHz5lU2H78GozMM2FZIC_uJ_4SLuHCOK0pHSlySB_RRAYI6ZtafMDyvo6C_TtCrsywVC8oWiaZ1H-3xwHlg-bAddL6YHThaxNp76uha8eOCIqmkMbeD2Kf4X8__-H8ZuMyMCCwAA';

// // Read embed URL from textbox
// var txtEmbedUrl = 'https://app.powerbi.com/qnaEmbed?groupId=be8908da-da25-452e-b220-163f52476cdd';

// // Read dataset Id from textbox
// var txtDatasetId = '88cc0a3e-e7ed-4fbd-8861-992d7fefa61c';

// // Read question from textbox
// var txtQuestion = 'show sales';

// // Read Qna mode
// var qnaMode = "ResultOnly";


// --------------------------------------------------------------------------------


// Approach #1: Use powerbi-api, doesn't seem to have a qna mode.

// var powerbi = require('powerbi-api'); //"powerbi-api": "^1.2.3",
// var msrest = require('ms-rest'); //"ms-rest": "^2.3.6",

// var credentials = new msrest.TokenCredentials('{AccessKey}', "AppKey");
// var client = new powerbi.PowerBIClient(credentials);

// // Example API call
// // client.workspaces.getWorkspacesByCollectionName('{WorkspaceCollection}', function(err, result) {
// //     // Your code here
// // });

// // Get models. models contains enums that can be used.
// var models = client.models;

// // Embed configuration used to describe the what and how to embed.
// // This object is used when calling powerbi.embed.
// // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Embed-Configuration-Details.
// var config= {
//     type: 'qna',
//     tokenType: models.TokenType.Embed,
//     accessToken: txtAccessToken,
//     embedUrl: txtEmbedUrl,
//     datasetIds: [txtDatasetId],
//     viewMode: models.QnaMode[qnaMode],
//     question: txtQuestion
// };

// // Get a reference to the embedded QNA HTML element
// var qnaContainer = $('#qnaContainer')[0];

// // Embed the QNA and display it within the div container.
// var qna = powerbi.embed(qnaContainer, config);

// // qna.off removes a given event handler if it exists.
// qna.off("loaded");

// // qna.on will add an event handler which prints to Log window.
// qna.on("loaded", function(event) {
//     Log.logText("QNA loaded event");
//     Log.log(event.detail);
// });


// --------------------------------------------------------------------------------


// Approach #2: Use powerbi-client, doesn't seem to have a way to retrieve Power BI response through REST etc. requires window object (browser).
//"powerbi-client": "^2.6.0",

// // Get models. models contains enums that can be used.
// var models = window['powerbi-client'].models;

// // Embed configuration used to describe the what and how to embed.
// // This object is used when calling powerbi.embed.
// // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Embed-Configuration-Details.
// var config= {
//     type: 'qna',
//     tokenType: models.TokenType.Embed,
//     accessToken: txtAccessToken,
//     embedUrl: txtEmbedUrl,
//     datasetIds: [txtDatasetId],
//     viewMode: models.QnaMode[qnaMode],
//     question: txtQuestion
// };

// // Get a reference to the embedded QNA HTML element
// var qnaContainer = $('#qnaContainer')[0];

// // Embed the QNA and display it within the div container.
// var qna = powerbi.embed(qnaContainer, config);

// // qna.off removes a given event handler if it exists.
// qna.off("loaded");

// // qna.on will add an event handler which prints to Log window.
// qna.on("loaded", function(event) {
//     Log.logText("QNA loaded event");
//     Log.log(event.detail);
// });