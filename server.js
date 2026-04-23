import express from 'express';
import imaps from 'imap-simple';
import { simpleParser } from 'mailparser';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';

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
            password: process.env.EMAIL_PASS,
            host: process.env.IMAP_HOST || 'imap.gmail.com',
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

        const messages = await connection.search(searchCriteria, fetchOptions);
        
        if (!messages || messages.length === 0) {
            connection.end();
            throw new Error('No emails found with subject "Purchase_Order_Reports".');
        }

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
            throw new Error('No Excel attachment found in the latest email.');
        }

        return attachment.content; // Returns the Buffer
    } catch (err) {
        console.error("IMAP Fetch Error:", err);
        throw err;
    }
};

// API Endpoint to get the PO data
app.get('/api/po-data', async (req, res) => {
    const forceRefresh = req.query.forceRefresh === 'true';
    
    if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
        return res.status(500).json({ error: "Email credentials not configured on the server." });
    }

    try {
        // Cache data for 1 hour to prevent hitting IMAP too frequently
        if (!cachedPOData || forceRefresh || (Date.now() - lastFetchTime > 1000 * 60 * 60)) {
            console.log("Connecting to email to fetch new PO data...");
            const buffer = await fetchLatestPOFromEmail();
            cachedPOData = buffer;
            lastFetchTime = Date.now();
        } else {
            console.log("Serving PO data from cache...");
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="po_report.xlsx"');
        res.send(cachedPOData);
    } catch (err) {
        res.status(500).json({ error: err.message || "Failed to fetch PO data." });
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
