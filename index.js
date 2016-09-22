var express = require('express');
var powerbi = require('powerbi-api');
var msrest = require('ms-rest');
var path = require('path');

const config = {
    port: process.env.PORT || 3000,
    accessKey: process.env.POWERBI_ACCESS_KEY,
    workspaceCollection: process.env.POWERBI_WORKSPACE_COLLECTION,
    workspaceId: process.env.POWERBI_WORKSPACE_ID
};

const app = express();
const credentials = new msrest.TokenCredentials(config.accessKey, "AppKey");
const powerbiClient = new powerbi.PowerBIClient(credentials); 

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname + '/index.html'));
});

app.get('/api/reports', (req, res) =>{
    powerbiClient.reports.getReports(config.workspaceCollection, config.workspaceId, (err, response) => {
        if(err){
            res.send(500, err.message);
            return;
        }
        res.send(response.value);
    });
});

app.get('/api/reports/:id', (req, res) => {
    powerbiClient.reports.getReports(config.workspaceCollection, config.workspaceId, (err, response) => {
        if(err){
            res.send(500, err.message);
            return;
        }
        var reportID = req.params.id;   // from URI 
        var reports = response.value;        // get reports from API's response above
        var filteredReports = reports.filter( report => report.id === reportID);    // filter out to get only specific report

        if(filteredReports.length !== 1){
            res.send(404, `Report with ID: ${reportID} is not found.`);
            return;
        }

        var report = filteredReports[0];    // here, we already found the requested report.
        // we need to generate a token for embed purpose. This is bound to a single report ID.
        var embedToken = powerbi.PowerBIToken.createReportEmbedToken(config.workspaceCollection, config.workspaceId,report.id);
        // and then we need to generate an access token from the embed token. This is to obtain necessary access credentials.
        var accessToken = embedToken.generate(config.accessKey);
        var embedConfig = Object.assign({
            type: 'report',
            accessToken
        }, report);

        res.send(embedConfig);
    });
});

app.listen(config.port, function () {
  console.log(`PowerBI API Server running on localhost:${config.port}`);
});