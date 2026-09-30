const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const app = express();

// Create uploads directory if it does not exist
const uploadsDir = path.join(__dirname, 'uploads');

if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true })); // Required to properly parse URL-encoded bodies and form data

// Serve uploaded files and frontend static files
app.use('/uploads', express.static('uploads'));
app.use(express.static(path.join(__dirname, 'public')));

// Connect Database
mongoose.connect(process.env.MONGO_URI || 'mongodb://localhost:27017/pdf-ai-suite')
  .then(() => console.log('Database connected successfully'))
  .catch(err => console.error('Database connection error:', err));

// Routes configuration
app.use('/api/auth', require('./routes/authRoutes'));
app.use('/api/pdf', require('./routes/pdfRoutes'));

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Open your app here: http://localhost:${PORT}`);
});