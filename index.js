var express = require('express');
var mongoose = require('mongoose');
require('dotenv').config()

var app = express();
var PORT = process.env.PORT || 3001;

mongoose.connect("mongodb://127.0.0.1:27017/azamatbot?directConnection=true&serverSelectionTimeoutMS=2000&appName=mongosh+1.9.1",
  {
    useNewUrlParser: true,
    useUnifiedTopology: true,
  }
)
  .then(() => {
    console.log("Database connected");
    app.listen(PORT, () => {
      console.log(`Server started on port ${PORT}`);
      require("./bot.js");
    });
  });