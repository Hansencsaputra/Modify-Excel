const { deleteRowsController } = require('./excelController');

const express = require('express');

const router = express.Router();

router.get('/readExcel', deleteRowsController);

module.exports = router;