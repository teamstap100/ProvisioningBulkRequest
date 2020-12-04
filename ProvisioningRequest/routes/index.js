'use strict';
var express = require('express');
var request = require('request');
var router = express.Router();

let auth = process.env["TVT_API_KEY"];

function cleanEmail(email) {
    console.log("Cleaning email");
    console.log(email);

    // Deal with undefined email
    if (!email) {
        return email;
    }

    email = email.toLowerCase();
    console.log(email);
    email = email.replace("#ext#@microsoft.onmicrosoft.com", "");
    console.log(email);
    if (email.includes("@")) {
        return email;

    } else if (email.includes("_")) {
        console.log("Going the underscore route");
        var underscoreParts = email.split("_");
        var domain = underscoreParts.pop();
        var tenantString = domain.split(".")[0];

        if (underscoreParts.length > 1) {
            email = underscoreParts.join("_") + "@" + domain;
        } else {
            email = underscoreParts[0] + "@" + domain;
        }
    }
    return email;
}

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

    let userEmail = cleanEmail(req.body.email);

    let tenantInfoParams = {
        url: "https://tap-validation-tab-admin-2.azurewebsites.net/api/tenants/email/" + userEmail,
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
