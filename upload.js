const mongoose = require('mongoose');
const fs = require('fs');
const xlsx = require('xlsx');

mongoose.connect('mongodb+srv://tanishqagarwal595:tanishq595@cluster0.skrcz2y.mongodb.net/myDatabase?retryWrites=true&w=majority&appName=Cluster0')
    .then(() => console.log('Connected to MongoDB'))
    .catch(err => console.error('Connection error:', err));

// Define a schema for your data
const DataSchema = new mongoose.Schema({
    name: String,
    theme: String,
    subtheme: String,
    category: String
});
const DataModel = mongoose.model('Data', DataSchema);

// Read the XLSX file
const workbook = xlsx.readFile('tool_set.xlsx');
const sheetName = workbook.SheetNames[0]; // Get the first sheet
const worksheet = workbook.Sheets[sheetName];

// Convert the sheet to JSON
const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Extract themes, subthemes, and categories
const themes = jsonData[0].slice(1); // Row 1 (after the name column)
const subthemes = jsonData[1].slice(1); // Row 2
const categories = jsonData[2].slice(1); // Row 3

const processedData = [];

// Iterate through names starting from the fourth row
for (let i = 3; i < jsonData.length; i++) {
    const row = jsonData[i];
    const name = row[0]; // First column is the name

    // Iterate through the rest of the columns to find "x"
    for (let j = 1; j < row.length; j++) {
        if (row[j] === 'x') {
            const theme = themes[j - 1]; // Adjust index for themes
            const subtheme = subthemes[j - 1]; // Adjust index for subthemes
            const category = categories[j - 1]; // Adjust index for categories

            // Create an entry for the database
            processedData.push({ name, theme, subtheme, category });
        }
    }
}

// Upload processed data to MongoDB
DataModel.insertMany(processedData)
    .then(() => console.log('Data uploaded'))
    .catch(err => console.error('Upload error:', err))
    .finally(() => mongoose.connection.close());