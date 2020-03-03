'use strict';
var express = require('express');
var router = express.Router();

/* GET home page. */
router.get('/', function (req, res) {
    res.render('public', { title: 'TAP Provisioning - R1.5' });
});

router.get('/internal', function (req, res) {
    res.render('internal', { title: 'TAP Provisioning - Internal' });
})

router.get('/r3', function (req, res) {
    res.render('public-r3', { title: "TAP Provisioning - R3" });
})

router.get('/config', function (req, res) {
    res.render('config');
})

module.exports = router;
