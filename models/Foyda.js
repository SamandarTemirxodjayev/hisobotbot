var mongoose = require("mongoose");

var foydaSchema = new mongoose.Schema({
  name: {
    type: String
  },
  amount: {
    type: Number
  },
  date: {
    day: {
      type: Number
    },
    month: {
      type: Number
    },
    year: {
      type: Number
    }
  },
  time: {
    hour: {
      type: Number
    },
    minute: {
      type: Number
    },
    second: {
      type: Number
    }
  }
});
var Foydas = mongoose.model("foyda", foydaSchema);

module.exports = Foydas;
