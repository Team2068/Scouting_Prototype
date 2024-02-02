const express = require('express');
const fs = require('fs');
const path = require('path');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');

const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));

app.post('/submit-form', (req, res) => {
  const formData = req.body;

  let jsonData = [];
  const dataFilePath = path.join(__dirname, 'data.json');

  try {
    const dataFileContent = fs.readFileSync(dataFilePath, 'utf8');
    jsonData = JSON.parse(dataFileContent);
  } catch (err) {
  }

  jsonData.push(formData);

  fs.writeFileSync(dataFilePath, JSON.stringify(jsonData, null, 2), 'utf8');

  const workbook = xlsx.utils.book_new();
  const sheet = xlsx.utils.json_to_sheet(jsonData);
  xlsx.utils.book_append_sheet(workbook, sheet, 'Form Data');
  const spreadsheetPath = path.join(__dirname, 'form_data.xlsx');
  xlsx.writeFile(workbook, spreadsheetPath);

  //res.send('Form data submitted successfully!');
  res.redirect('/success.html');
});

app.listen(port, '0.0.0.0', () => {
  console.log(`Server is running at http://0.0.0.0:${port}`);
});
