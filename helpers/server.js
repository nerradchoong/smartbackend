const express = require('express');
const session = require('express-session');
const cors = require('cors');
const bodyParser = require('body-parser');
const excelRoutes = require('./routes/excelRoutes');

const app = express();
const PORT = 3001;

// Setup CORS
app.use(cors({
    origin: 'http://localhost:3000', // Specify the origin of the frontend
    credentials: true, // Enable credentials such as cookies
    optionsSuccessStatus: 200 // Some legacy browsers (IE11, various SmartTVs) choke on 204
}));

// Body parser middleware
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));

// Setup session middleware
app.use(session({
    secret: 'your_secret_key', // Replace 'your_secret_key' with a real secret string
    resave: false,
    saveUninitialized: true,
    cookie: { 
        secure: false, // Set to true if you're using https (required for cookies to be set on HTTPS)
        httpOnly: true // Helps against XSS attacks
    }
}));

// Excel Routes
app.use('/api/excel', excelRoutes);

// Start the server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
