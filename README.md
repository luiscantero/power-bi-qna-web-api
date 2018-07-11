# Node.js Express Web API server to render Power BI Q&A as image
## Description
- Server that forwards a natural language query to a headless browser that renders the Power BI answer
- A screenshot of the answer (plot or text) is returned to the client
- Can be used e.g. to render Power BI data in a chat bot

## Steps
1. Run: `node pbiqnawebapi.js`
2. Send GET query e.g. using a browser: `http://localhost:8081/api/<QUERY>`
3. Sample query: http://localhost:8081/api/pbiqna/show%20average%20sales%20by%20city%20as%20pie%20chart

### License
[MIT](http://opensource.org/licenses/MIT)

## Reference
- [Q&A in Power BI Embedded](https://docs.microsoft.com/en-us/power-bi/developer/qanda)
- [Sample Q&A](https://microsoft.github.io/PowerBI-JavaScript/demo/v2-demo/index.html)
- [Puppeteer API](https://github.com/GoogleChrome/puppeteer/blob/master/docs/api.md)