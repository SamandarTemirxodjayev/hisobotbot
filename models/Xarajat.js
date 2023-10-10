var mongoose = require("mongoose");

var xarajatSchema = new mongoose.Schema({
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
var Xarajats = mongoose.model("xarajat", xarajatSchema);

module.exports = Xarajats;
