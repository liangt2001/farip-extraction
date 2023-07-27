var express = require('express');
var router = express.Router();

const index_controller = require("../controllers/indexController.js");
const index_controller_updated = require("../controllers/indexController_updated.js");
const index_controller_callback = require("../controllers/indexController_callback.js");
const new_controller = require("../controllers/new_indexController.js");

/* GET home page. */
// router.get('/', function(req, res, next) {
//   res.render('index', { title: 'Express' });
// });

router.get("/", index_controller.index_get);

router.post("/import", new_controller.index_post);

router.get("/download", new_controller.download_get);

module.exports = router;
