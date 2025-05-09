// server.js
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

mongoose.connect('mongodb+srv://tanishqagarwal595:tanishq595@cluster0.skrcz2y.mongodb.net/myDatabase?retryWrites=true&w=majority&appName=Cluster0')
    .then(() => console.log('Connected to MongoDB'))
    .catch(err => console.error('Connection error:', err));

const DataSchema = new mongoose.Schema({}, { strict: false });
const DataModel = mongoose.model('Data', DataSchema);

app.get('/api/data', async (req, res) => {
    const data = await DataModel.find();
    res.json(data);
});

app.post('/api/data', async (req, res) => {
    const newData = new DataModel(req.body);
    await newData.save();
    res.json(newData);
});

app.delete('/api/data/:name', async (req, res) => {
    await DataModel.deleteOne({ name: req.params.name });
    res.json({ success: true });
});

app.listen(3000, () => console.log('Server running on port 3000'));