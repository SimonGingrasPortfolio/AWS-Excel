'use strict'
var AWS = require('aws-sdk');
var S3 = new AWS.S3();
var SES = new AWS.SES({region: 'us-east-1'});
var nodemailer = require('nodemailer');
var excel = require('excel4node');
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');

exports.handler = (event, context, callback) => {
  var docClient = new AWS.DynamoDB.DocumentClient({ region: 'us-east-2' });
  var params = { TableName: 'Users' };
  docClient.scan(params, (err, data) => {
    if (err) {
      callback(err, null);
    } else {
      var items = data.Items;
      var date = new Date();
      date = date.toLocaleString('en-US', { dateStyle: 'long', timeZone: 'America/Los_Angeles' });
      var dateParts = /(.*?)\s(\d*)\,\s(\d*)/g.exec(date);
      var year = dateParts[3];
      var month = dateParts[1].substring(0, 3);
      var day = dateParts[2];
      day = (day.toString()).length > 1 ? day : '0' + day;
      date = `${year} ${month} ${day}`;
      // Headers + Columns
      worksheet.cell(1, 1).string('Date').style({font: {bold: true}});
      worksheet.cell(1, 2).string('Email').style({font: {bold: true}});
      worksheet.cell(1, 3).string('Name').style({font: {bold: true}});
      worksheet.cell(1, 4).string('Birthday').style({font: {bold: true}});
      // Rows
      items.sort((a, b) => (a.date > b.date) ? -1 : 1);
      items.forEach((item, i) => {
        worksheet.cell(i + 2, 1).string(item.date);
        worksheet.cell(i + 2, 2).string(item.email);
        worksheet.cell(i + 2, 3).string(item.name);
        worksheet.cell(i + 2, 4).string(item.birthday);
      });
      workbook.writeToBuffer().then(buffer => {
        var params = {
          Bucket: 'user-data',
          Key: `xlsx/${date}.xlsx`,
          Body: buffer,
          ACL: 'public-read'
        }
        S3.upload(params, function(err, data) {
          if (err) {
            console.log(err, err.stack);
          } else {
            var options = {
              from: 'donotreply@yourdomain.com',
              subject: 'User Report',
              to: 'info@yourdomain.com',
              attachments: [{
                filename: `${date}.xlsx`,
                content: buffer
              }]
            };
            var transporter = nodemailer.createTransport({ SES });
            transporter.sendMail(options, (err, info) => { // Send Email
              if (err) {
                console.log('Error sending report');
                callback(err);
              } else {
                callback();
              }
            });
          }
        });
      })
    }
  });
};
