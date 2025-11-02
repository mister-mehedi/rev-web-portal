const express = require('express');
const bodyParser = require('body-parser');
const dotenv = require('dotenv');
const indexRouter = require('./routes/index');

dotenv.config();
const app = express();

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));
app.use('/', indexRouter);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
