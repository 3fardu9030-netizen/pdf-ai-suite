require('dotenv').config();
const mongoose = require('mongoose');

const uri = process.env.MONGO_URI;

console.log('Testing connection to MongoDB...');
console.log('Using URI:', uri ? uri.substring(0, 25) + '...' : 'UNDEFINED');

mongoose.connect(uri)
  .then(() => {
    console.log('SUCCESS! MongoDB connected successfully.');
    process.exit(0);
  })
  .catch(err => {
    console.error('FAILED! Could not connect to MongoDB.');
    console.error(err.message);
    process.exit(1);
  });