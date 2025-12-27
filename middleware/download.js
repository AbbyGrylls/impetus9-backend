// middleware/download.js
const ExcelJS = require('exceljs');
const registerModel = require('../models/registerModel');
const CoordinatorLog = require('../models/coordsLogModel');

// --- Helper: vCard Generator ---
const generateVCard = (name, phone, filename) => {
    return `BEGIN:VCARD
VERSION:3.0
FN:${filename}
N:;${filename};;;
TEL;TYPE=CELL:${phone}
END:VCARD
`;
};

const handleDownload = async (req, res) => {
    try {
        const { eventName, coordsValue, coordinatorName, passkey } = req.body;

        // --- STEP 1: Identify User Role (Admin vs Coordinator) ---
        // Check if a Master Key is set in env and if the provided passkey matches it
        const masterKey = process.env.PASSKEY_MASTER;
        const isAdmin = (masterKey && passkey === masterKey);

        // If NOT Admin, strictly verify the specific Event Passkey
        if (!isAdmin) {
            const envKey = `PASSKEY_${coordsValue.toUpperCase()}`;
            if (!process.env[envKey] || process.env[envKey] !== passkey) {
                return res.status(401).json({ error: "Invalid Passkey" });
            }
        }

        // Fetch all registrations for the event
        const registrations = await registerModel.find({ eventName })
            .sort({ createdAt: -1 })
            .lean();

        if (!registrations || registrations.length === 0) {
            return res.status(200).json({ 
                success: false, 
                message: "No one registered yet!" 
            });
        }

        // --- STEP 2: Ensure Log Entry Exists ---
        // We ensure the document exists so we don't get null errors later
        let existingLog = await CoordinatorLog.findOne({ eventName });
        if (!existingLog) {
            try {
                existingLog = await CoordinatorLog.create({ eventName });
            } catch (e) {
                // Handle potential race condition if created between find and create
                existingLog = await CoordinatorLog.findOne({ eventName });
            }
        }

        // --- STEP 3: Handle Locking Logic (Branching) ---
        let isFirstDownload = false;
        let logDetails = existingLog;

        if (isAdmin) {
            // === ADMIN PATH (GHOST MODE) ===
            // 1. Do NOT update the database.
            // 2. Do NOT claim the lock.
            // 3. Just read the current state to inform the admin.
            isFirstDownload = false; // Admin is never the "first" in the DB sense
            logDetails = await CoordinatorLog.findOne({ eventName });
        } else {
            // === COORDINATOR PATH (STANDARD LOGIC) ===
            // Attempt to claim the lock (change vCardsDownloaded from false -> true)
            const lockResult = await CoordinatorLog.findOneAndUpdate(
                { eventName: eventName, vCardsDownloaded: false },
                { 
                    $set: { 
                        vCardsDownloaded: true, 
                        firstDownloaderName: coordinatorName,
                        downloadTime: new Date()
                    }
                },
                { new: true } // Return the updated document if successful
            );

            if (lockResult) {
                isFirstDownload = true; // Success! This coordinator claimed the lock.
                logDetails = lockResult;
            } else {
                isFirstDownload = false; // Failed, someone else already took it.
                logDetails = await CoordinatorLog.findOne({ eventName });
            }
        }

        // --- STEP 4: EXCEL GENERATION (Shared by both) ---
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Participants');

        let maxTeamMembers = 0;
        registrations.forEach(reg => {
            if (reg.teamMembers && reg.teamMembers.length > maxTeamMembers) {
                maxTeamMembers = reg.teamMembers.length;
            }
        });

        const columns = [
            { header: 'Team Name', key: 'teamName', width: 25 },
            { header: 'Captain Name', key: 'capName', width: 20 },
            { header: 'Captain Phone', key: 'capPhone', width: 15 },
            { header: 'Captain Roll', key: 'capRoll', width: 15 },
            { header: 'Type', key: 'type', width: 10 },
            { header: 'Registered At', key: 'regAt', width: 20 },
        ];

        for (let i = 1; i <= maxTeamMembers; i++) {
            columns.push({ header: `Mem ${i} Name`, key: `m${i}Name`, width: 20 });
            columns.push({ header: `Mem ${i} Roll`, key: `m${i}Roll`, width: 15 });
            columns.push({ header: `Mem ${i} Phone`, key: `m${i}Phone`, width: 15 });
        }

        worksheet.columns = columns;

        registrations.forEach(reg => {
            let rowData = {
                teamName: reg.teamName,
                capName: reg.capName,
                capPhone: reg.capPhone,
                capRoll: reg.participantType === 'INTERNAL' ? reg.capRoll : 'EXTERNAL',
                type: reg.participantType,
                regAt: new Date(reg.createdAt).toLocaleString()
            };

            if (reg.teamMembers && Array.isArray(reg.teamMembers)) {
                reg.teamMembers.forEach((member, index) => {
                    const i = index + 1;
                    rowData[`m${i}Name`] = member.memName || '-';
                    rowData[`m${i}Roll`] = member.memRoll || '-';
                    rowData[`m${i}Phone`] = member.memPhone || '-';
                });
            }
            worksheet.addRow(rowData);
        });

        worksheet.getRow(1).font = { bold: true };
        const buffer = await workbook.xlsx.writeBuffer();
        const base64Excel = buffer.toString('base64');

        // --- STEP 5: vCard & Message Logic ---
        let vCardContent = "";
        let message = "";

        // CONDITION: Generate vCards if (User is Admin) OR (User is the First Coordinator)
        if (isAdmin || isFirstDownload) {
            
            // Set the Message
            if (isAdmin) {
                // Admin Info Message
                if (logDetails.vCardsDownloaded) {
                    const timeStr = new Date(logDetails.downloadTime).toLocaleString();
                    message = `ADMIN MODE: Retrieved all data. (Note: Originally downloaded by ${logDetails.firstDownloaderName} at ${timeStr})`;
                } else {
                    message = "ADMIN MODE: Retrieved all data. (Status: Not yet downloaded by any coordinator)";
                }
            } else {
                // Coordinator Success Message
                message = "You are the first coordinator, You can download both Contacts and the Excel Sheet";
            }

            // Generate the cards
            registrations.forEach(reg => {
                const prefix = eventName.substring(0, 2).toLowerCase();
                let uniqueId;
                if (reg.participantType === 'INTERNAL') {
                    uniqueId = `${prefix}${reg.capRoll}`;
                } else {
                    const phoneSuffix = reg.capPhone.replace(/\D/g, '').slice(-8); 
                    uniqueId = `${prefix}EXT${phoneSuffix}`;
                }
                vCardContent += generateVCard(reg.capName, reg.capPhone, `${uniqueId}-1`);
                
                if (reg.teamMembers && reg.teamMembers[0] && reg.teamMembers[0].memPhone) {
                    vCardContent += generateVCard(
                        reg.teamMembers[0].memName, 
                        reg.teamMembers[0].memPhone, 
                        `${uniqueId}-2`
                    );
                }
            });

        } else {
            // Coordinator Failure Message (Late)
            const timeStr = new Date(logDetails.downloadTime).toLocaleString();
            message = `⚠ Alert : Contacts were ALREADY downloaded by *${logDetails.firstDownloaderName}*, At ${timeStr}.`;
        }

        return res.status(200).json({
            success: true,
            message: message,
            excelBase64: base64Excel,
            vcf: (isAdmin || isFirstDownload) ? vCardContent : null
        });

    } catch (err) {
        console.error("Download Middleware Error:", err);
        return res.status(500).json({ error: "Server Error processing download" });
    }
};

module.exports = handleDownload;