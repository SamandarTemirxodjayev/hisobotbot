console.log("bot is running!");
require('dotenv').config()
var TelegramBot = require("node-telegram-bot-api");
const DataValue = require('./models/DataValue');
const Xarajats = require('./models/Xarajat');
const Foydas = require('./models/Foyda');
const ExcelJS = require('exceljs');
const bot = new TelegramBot(process.env.BOT_TOKEN, { polling: true });

const stripos = (haystack, needle) => {
  const haystackLower = haystack.toLowerCase();
  const needleLower = needle.toLowerCase();
  return haystackLower.indexOf(needleLower);
};

bot.on("message", async (msg) => {
  const userData = await DataValue.findOne({ tgId: msg.chat.id });
  if (userData && userData.value){
    if(userData.value == "xarajat"){
      bot.sendMessage(msg.chat.id, "Xarajat Summasini Yuboring");
      userData.value = `xarajat||${msg.text}`;
      await userData.save();
    }else if(stripos(userData.value, "xarajat||")!== -1){
      const maqsad = userData.value.split("||")[1];
      console.log(maqsad);
      if(isNaN(msg.text)){
        return bot.sendMessage(msg.chat.id, "Iltimos sonda kiriting");
      }
      const newXarajat = await Xarajats({
        name: maqsad,
        amount: msg.text,
        date: {
          day: new Date().getDate(),
          month: new Date().getMonth() + 1,
          year: new Date().getFullYear()
        },
        time: {
          hour: new Date().getHours(),
          minute: new Date().getMinutes(),
          second: new Date().getSeconds()
        }
      });
      await newXarajat.save();
      await DataValue.deleteOne({tgId: msg.chat.id});
      bot.sendMessage(msg.chat.id, "Xarajat qo'shildi");
    }
    if(userData.value == "foyda"){
      bot.sendMessage(msg.chat.id, "Foyda Summasini Yuboring");
      userData.value = `foyda||${msg.text}`;
      await userData.save();
    }else if(stripos(userData.value, "foyda||")!== -1){
      const maqsad = userData.value.split("||")[1];
      console.log(maqsad);
      if(isNaN(msg.text)){
        return bot.sendMessage(msg.chat.id, "Iltimos sonda kiriting");
      }
      const newXarajat = await Foydas({
        name: maqsad,
        amount: msg.text,
        date: {
          day: new Date().getDate(),
          month: new Date().getMonth() + 1,
          year: new Date().getFullYear()
        },
        time: {
          hour: new Date().getHours(),
          minute: new Date().getMinutes(),
          second: new Date().getSeconds()
        }
      });
      await newXarajat.save();
      await DataValue.deleteOne({tgId: msg.chat.id});
      bot.sendMessage(msg.chat.id, "Foyda qo'shildi");
    }
  }
  if(msg.text == "/start"){
    bot.sendMessage(msg.chat.id, "Welcome to Azamat Bot", {
      reply_markup: {
        resize_keyboard: true,
        keyboard: [
          [{ text: "Xarajat Qo'shish" }],
          [{ text: "Foyda Qo'shish" }],
          [{ text: "EXCEL" }]
        ]
      }
    });
  }
  if(msg.text == "Xarajat Qo'shish"){
    bot.sendMessage(msg.chat.id, "Xarajat Maqsadini Yuboring");
    const newDataValue = await DataValue({
      tgId: msg.chat.id,
      value: "xarajat"
    });
    await newDataValue.save();
  }
  if(msg.text == "Foyda Qo'shish"){
    bot.sendMessage(msg.chat.id, "Foyda komentariyasi");
    const newDataValue = await DataValue({
      tgId: msg.chat.id,
      value: "foyda"
    });
    await newDataValue.save();
  }
  if(msg.text == "EXCEL"){
    bot.sendMessage(msg.chat.id, "Tayyorlanmoqda...");
    const cellStyle = {
      font: { bold: true },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: {
        top: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
      },
    };
    const labels = [
      "Yanvar",
      "Fevral",
      "Mart",
      "Aprel",
      "May",
      "Iyun",
      "Iyul",
      "Avgust",
      "Sentyabr",
      "Oktyabr",
      "Noyabr",
      "Dekabr",
    ];
    const now = new Date();
    const year = now.getFullYear();
    const workbook = new ExcelJS.Workbook();
    const worksheet1 = workbook.addWorksheet('Foyda');
    const worksheet2 = workbook.addWorksheet('Xarajat');
    const worksheet3 = workbook.addWorksheet('Hisobot');
    worksheet1.getCell(1, 1).value = 'ID';
    worksheet1.getCell(1, 2).value = 'Maqsad';
    worksheet1.getCell(1, 3).value = 'Summa';
    worksheet1.getCell(1, 4).value = 'Sana va Vaqti';
    worksheet1.mergeCells('F1:G1');
    worksheet1.getCell(1, 6).value = `Foyda ${year}`;
    const foydas = await Foydas.find();

    worksheet2.getCell(1, 1).value = 'ID';
    worksheet2.getCell(1, 2).value = 'Maqsad';
    worksheet2.getCell(1, 3).value = 'Summa';
    worksheet2.getCell(1, 4).value = 'Sana va Vaqti';
    worksheet2.mergeCells('F1:G1');
    worksheet2.getCell(1, 6).value = `Xarajat ${year}`;
    const xarajats = await Xarajats.find();

    worksheet1.getColumn(1).width = 5; 
    worksheet1.getColumn(2).width = 20;
    worksheet1.getColumn(3).width = 20;
    worksheet1.getColumn(4).width = 30;
    worksheet1.getColumn(6).width = 25;
    worksheet1.getColumn(7).width = 25;

    worksheet2.getColumn(1).width = 5; 
    worksheet2.getColumn(2).width = 20;
    worksheet2.getColumn(3).width = 20;
    worksheet2.getColumn(4).width = 30;
    worksheet2.getColumn(6).width = 25;
    worksheet2.getColumn(7).width = 25;

    let i = 0;
    let foydaSummary = 0;

    foydas.forEach((item, index) => {
      i++;
      worksheet1.getCell(index + 2, 1).value = index + 1;
      worksheet1.getCell(index + 2, 2).value = item.name;
      worksheet1.getCell(index + 2, 3).value = `${item.amount} so'm`;
      foydaSummary += item.amount;
      worksheet1.getCell(index + 2, 4).value = `${item.date.day}.${item.date.month}.${item.date.year} ${item.time.hour}:${item.time.minute}:${item.time.second}`;
    });

    let j = 0;
    let xarajatSummary = 0;

    xarajats.forEach((item, index) => {
      j++;
      worksheet2.getCell(index + 2, 1).value = index + 1;
      worksheet2.getCell(index + 2, 2).value = item.name;
      worksheet2.getCell(index + 2, 3).value = `${item.amount} so'm`;
      xarajatSummary += item.amount;
      worksheet2.getCell(index + 2, 4).value = `${item.date.day}.${item.date.month}.${item.date.year} ${item.time.hour}:${item.time.minute}:${item.time.second}`;
    });
    
    worksheet1.getCell(i+2, 2).value = 'Umumiy Summa';
    worksheet1.getCell(i+2, 3).value = `${foydaSummary} so'm`;

    worksheet2.getCell(j+2, 2).value = 'Umumiy Summa';
    worksheet2.getCell(j+2, 3).value = `${xarajatSummary} so'm`;

    const result = await Foydas.aggregate([
      {
        $group: {
          _id: {
            year: "$date.year",
            month: "$date.month",
          },
          totalAmount: { $sum: "$amount" },
        },
      },
      {
        $sort: {
          "_id.year": 1,
          "_id.month": 1,
        },
      },
      {
        $project: {
          _id: 0,
          year: "$_id.year",
          month: "$_id.month",
          totalAmount: 1,
        },
      },
    ]);

    const monthlySummary = {};

    for (let i = 1; i <= 12; i++) {
      const month = i;
      const matchingMonth = result.find((item) => item.month === month && item.year === year);
      const totalAmount = matchingMonth ? matchingMonth.totalAmount : 0;
      monthlySummary[month] = totalAmount;
    }

    labels.forEach((item, index) => {
      worksheet1.getCell(index + 2, 6).value = item;
      worksheet1.getCell(index + 2, 7).value = `${monthlySummary[index + 1]} so'm`;
    });

    const lastRowIndex = labels.length + 2;
    worksheet1.getCell(lastRowIndex, 6).value = 'Umumiy Summa';
    worksheet1.getCell(lastRowIndex, 7).value = `${foydaSummary} so'm`;

    const result2 = await Xarajats.aggregate([
      {
        $group: {
          _id: {
            year: "$date.year",
            month: "$date.month",
          },
          totalAmount: { $sum: "$amount" },
        },
      },
      {
        $sort: {
          "_id.year": 1,
          "_id.month": 1,
        },
      },
      {
        $project: {
          _id: 0,
          year: "$_id.year",
          month: "$_id.month",
          totalAmount: 1,
        },
      },
    ]);

    const monthlySummary2 = {};

    for (let i = 1; i <= 12; i++) {
      const month = i;
      const matchingMonth = result2.find((item) => item.month === month && item.year === year);
      const totalAmount = matchingMonth ? matchingMonth.totalAmount : 0;
      monthlySummary2[month] = totalAmount;
    }

    labels.forEach((item, index) => {
      worksheet2.getCell(index + 2, 6).value = item;
      worksheet2.getCell(index + 2, 7).value = `${monthlySummary2[index + 1]} so'm`;
    });

    worksheet2.getCell(lastRowIndex, 6).value = 'Umumiy Summa';
    worksheet2.getCell(lastRowIndex, 7).value = `${xarajatSummary} so'm`;

    worksheet3.mergeCells('A1:B1');
    worksheet3.mergeCells('A4:B4');
    worksheet3.getColumn(1).width = 25;
    worksheet3.getColumn(2).width = 25;

    worksheet3.getCell(1, 1).value = 'Hisobot';
    worksheet3.getCell(2, 1).value = 'Umumiy Xarajat:';
    worksheet3.getCell(2, 2).value = `${xarajatSummary} so'm`;
    worksheet3.getCell(3, 1).value = 'Umumiy Foyda:';
    worksheet3.getCell(3, 2).value = `${foydaSummary} so'm`;
    worksheet3.getCell(4, 1).value = `Umumiy: ${(foydaSummary - xarajatSummary).toFixed(2)} so'm`;

    worksheet1.eachRow((row) => {
      row.eachCell((cell) => {
        cell.style = cellStyle;
      });
    });
    worksheet2.eachRow((row) => {
      row.eachCell((cell) => {
        cell.style = cellStyle;
      });
    });
    worksheet3.eachRow((row) => {
      row.eachCell((cell) => {
        cell.style = cellStyle;
      });
    });
    await workbook.xlsx.writeFile('./docs/HISOBOT.xlsx');
    await bot.sendDocument(msg.chat.id, "./docs/HISOBOT.xlsx");
  }
});