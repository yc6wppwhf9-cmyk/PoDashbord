import express from 'express';
import imaps from 'imap-simple';
import { simpleParser } from 'mailparser';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import cron from 'node-cron';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());

// Serve the static React files built by Vite
app.use(express.static(path.join(__dirname, 'dist')));

let cachedPOData = null;
let lastFetchTime = null;

// Helper to fetch the latest PO email attachment
const fetchLatestPOFromEmail = async () => {
    const config = {
        imap: {
            user: process.env.EMAIL_USER,
            password: process.env.EMAIL_PASS ? process.env.EMAIL_PASS.replace(/\s+/g, '') : undefined,
            host: process.env.IMAP_HOST || process.env.EMAIL_IMAP_SERVER || 'imap.gmail.com',
            port: parseInt(process.env.IMAP_PORT || '993', 10),
            tls: true,
            authTimeout: 15000,
            tlsOptions: { rejectUnauthorized: false }
        }
    };

    try {
        const connection = await imaps.connect(config);
        await connection.openBox('INBOX');

        // Search for emails with the subject
        const searchCriteria = [
            ['HEADER', 'SUBJECT', 'Purchase_Order_Reports']
        ];
        
        const fetchOptions = {
            bodies: [''],
            struct: true,
            markSeen: false
        };

        console.log("Searching for emails...");
        const messages = await connection.search(searchCriteria, fetchOptions);
        
        if (!messages || messages.length === 0) {
            connection.end();
            throw new Error('IMAP Connected, but no emails found with subject "Purchase_Order_Reports".');
        }

        console.log(`Found ${messages.length} emails. Parsing the newest one...`);

        // Sort by date descending to get the newest
        messages.sort((a, b) => new Date(b.attributes.date).getTime() - new Date(a.attributes.date).getTime());
        const latestMessage = messages[0];

        // We need to fetch the full body to parse attachments
        const all = latestMessage.parts.find(part => part.which === '');
        const id = latestMessage.attributes.uid;
        const idHeader = "Imap-Id: " + id + "\r\n";
        
        const parsedEmail = await simpleParser(idHeader + all.body);
        connection.end();

        // Look for the Excel attachment
        const attachment = parsedEmail.attachments.find(att => 
            att.filename && (att.filename.endsWith('.xlsx') || att.filename.endsWith('.xls') || att.filename.endsWith('.csv'))
        );

        if (!attachment) {
            const fileNames = parsedEmail.attachments.map(a => a.filename).join(', ');
            throw new Error(`Email parsed, but no Excel attachment found. Found attachments: ${fileNames || 'None'}`);
        }

        console.log("Attachment successfully extracted!");
        return attachment.content; // Returns the Buffer
    } catch (err) {
        console.error("IMAP Fetch Error:", err);
        throw new Error(`IMAP Error: ${err.message}`);
    }
};

// Background Update Function
let isFetching = false;
const updateCacheBackground = async () => {
    if (isFetching) return;
    if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
        console.warn("Skipping background fetch: Email credentials not configured.");
        return;
    }

    try {
        isFetching = true;
        console.log(`[${new Date().toLocaleString()}] Starting background email sync...`);
        const buffer = await fetchLatestPOFromEmail();
        cachedPOData = buffer;
        lastFetchTime = Date.now();
        console.log(`[${new Date().toLocaleString()}] Background sync complete. Data cached!`);
    } catch (err) {
        console.error(`[${new Date().toLocaleString()}] Background sync failed:`, err.message);
    } finally {
        isFetching = false;
    }
};

// Schedule Cron Job to run at 9:45 AM India Standard Time
cron.schedule('45 9 * * *', () => {
    console.log("Cron triggered 9:45 AM sync.");
    updateCacheBackground();
}, {
    scheduled: true,
    timezone: "Asia/Kolkata"
});

// Run an initial fetch 3 seconds after server starts
setTimeout(updateCacheBackground, 3000);

// API Endpoint to get the PO data
app.get('/api/po-data', async (req, res) => {
    const forceRefresh = req.query.forceRefresh === 'true';
    
    if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
        return res.status(500).json({ error: "Email credentials not configured on the server." });
    }

    try {
        if (forceRefresh) {
            console.log("Client requested force refresh. Syncing now...");
            await updateCacheBackground();
        } else if (!cachedPOData) {
            console.log("Cache empty. Syncing now...");
            await updateCacheBackground();
        } else {
            console.log("Serving instantly from memory cache!");
        }

        if (!cachedPOData) {
            throw new Error("Data could not be fetched from email.");
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="po_report.xlsx"');
        res.send(cachedPOData);
    } catch (err) {
        console.error("API Error Handler:", err);
        res.status(500).json({ error: err.message || "Failed to fetch PO data from email server." });
    }
});

// Fallback for React routing
app.use((req, res) => {
    res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
