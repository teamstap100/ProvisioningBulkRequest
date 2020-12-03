'use strict';
var express = require('express');
var request = require('request');
var router = express.Router();

let auth = process.env["TVT_API_KEY"];

/* GET home page. */
router.get('/', function (req, res) {
    res.render('public', { title: 'TAP Provisioning - R1.5' });
});

router.get('/internal', function (req, res) {
    res.render('internal', { title: 'TAP Provisioning - Internal' });
});

router.get('/r3', function (req, res) {
    res.render('public-r3', { title: "TAP Provisioning - R3" });
});

router.get('/config', function (req, res) {
    res.render('config');
});

router.post("/api/tenants", function (req, res) {
    console.log(req.body);
    console.log("Not yet implemented");

    let tenantInfoParams = {
        url: "https://tap-validation-tab-admin-2.azurewebsites.net/api/tenants/email/" + req.body.email,
        headers: {
            'X-API-KEY': auth
        }
    }

    console.log(tenantInfoParams);

    request.get(tenantInfoParams, function (err, resp, body) {
        console.log(resp.statusCode);
        console.log(body);
        return res.json(body);

    });
});

router.get("/api/tenants/:tid", function (req, res) {
    console.log(req.params);

    let tenantInfoParams = {
        url: "https://tap-validation-tab-admin-2.azurewebsites.net/api/tenants/" + req.params.tid,

        headers: {
            'X-API-KEY': auth
        }
    }

    console.log(tenantInfoParams);

    request.get(tenantInfoParams, function (err, resp, body) {
        console.log(resp.statusCode);
        console.log(body);
        return res.json(body);
    });
});

router.get("/api/validations", function (req, res) {
    console.log("Not yet implemented");
    return res.json({});
});

module.exports = router;
