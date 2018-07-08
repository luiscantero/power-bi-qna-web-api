/// <reference path="node_modules/@types/node/index.d.ts" />

"use strict";

// Config.
const CONFIG = require('./config');
const RENDER_DELAY = CONFIG.RENDER_DELAY;

const express = require('express'),
    fs = require('fs'),
    app = express(),
    puppeteer = require('puppeteer');

const https = require('https'); // For getSampleToken.
var browser, page; // Puppeteer.

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
    browser = await puppeteer.launch(); // For testing: { headless: false }
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

    console.log(`Setting content ...`);
    await page.goto(`data:text/html;charset=UTF-8,${html}`, { waitUntil: 'networkidle0' }); // setContent does not support with waitUntil.

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