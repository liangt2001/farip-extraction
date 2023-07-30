var express = require('express');
var router = express.Router();

const index_controller = require("../controllers/new_indexController.js");

/* GET home page. */
// router.get('/', function(req, res, next) {
//   res.render('index', { title: 'Express' });
// });

router.get("/", index_controller.index_get);

router.post("/import", index_controller.index_post);

router.get("/download", index_controller.download_get);

module.exports = router;
