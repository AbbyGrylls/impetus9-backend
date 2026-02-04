const express=require('express')
const router=express.Router()
const axios = require('axios');
const multer = require('multer'); // Added
const path = require('path'); 
const fs = require('fs');
const registerModel=require('../models/registerModel')
const rateLimiter=require('../middleware/rateLimiter')
const allowedEvents=require('../allowedEvents')

const uploadDir = 'uploads/';
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, 'uploads/'),
    filename: (req, file, cb) => {
        let name = req.body.capName || 'user';
        let phone = req.body.capPhone || '0000';
        const shortName = name.replace(/[^a-zA-Z0-9]/g, '').substring(0, 5).toLowerCase();
        const shortPhone = phone.slice(-4); 
        const uniqueId = Date.now().toString().slice(-6) + Math.round(Math.random() * 100);
        cb(null, `${shortPhone}_${shortName}_${uniqueId}${path.extname(file.originalname)}`);
    }
});
const upload = multer({ 
    storage: storage,
    limits: { fileSize: 5 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        if (file.mimetype.startsWith('image/')) cb(null, true);
        else cb(new Error('Only image files are allowed'), false);
    }
});

//route
router.post("/",rateLimiter,upload.single('paymentScreenshot'), async(req,res)=>{
    try{
        const data = prepareRegistrationPayload(req);
        const{
            eventName,
            teamName,
            capName,
            capPhone,
            capRoll,
            teamMembers,
            deviceFingerprint,
            participantType,
        }=data;
        /* const OPEN_EVENTS = ["Photo", "BGMI","IQ Ignition"]; 
        if (!OPEN_EVENTS.includes(eventName)) {
            return res.status(400).json({ 
                error: "Registration for this event is not opened." 
            });
        } */
        if(!allowedEvents.has(eventName)){
            return res.status(400).json({error:"Invalid Event"})
        }
        if(!teamName||!capName||!capPhone||!deviceFingerprint){
            return res.status(400).json({error:"Missing required fields"});
        }
        const deviceCount=await registerModel.countDocuments({deviceFingerprint:deviceFingerprint})
        if(deviceCount>=5){
            return res.status(429).json({
                error:"Device Limit Reached: You have registered too many times from this device"
            })
        }

        //controller
        const register=new registerModel(data);
        await register.save();
        sendTelegramNotification(data).catch(err => console.error("Telegram Error:", err.message));
        if (data.participantType === "EXTERNAL" && data.paymentScreenshot) {
            return res.status(201).json({
                success: true,
                message: "Registration Submitted for Verification",
                receiptId: "PENDING", 
                status: "PENDING"
            });
        }
        return res.status(201).json({
            success:true,
            message: "Registration Successful",
            receiptId: register.receiptId,
            status: "VERIFIED"
        });

    }
    catch (err){
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        if (err.code===11000) {
            console.log("FULL DUPLICATE PATTERN:", err.keyPattern);
            const keys = Object.keys(err.keyPattern);
            if (keys.includes('capPhone')) {
                return res.status(409).json({ error: 'This Phone Number is already registered for this event.' });
            }
            if (keys.includes('capRoll')) {
                return res.status(409).json({ error: 'This Roll Number is already registered for this event.' });
            }
            return res.status(409).json({
                error: 'Duplicate Registration: This team/captain is already registered.'
            });
        }
        if(err.name==='ValidationError'){
            return res.status(400).json({error: err.message});
        }
        return res.status(400).json({
            error:err.message || "Registration Failed"
        });
    }
});

module.exports=router;
function prepareRegistrationPayload(req) {
    let payload = { ...req.body };
    if (typeof payload.teamMembers === 'string') {
        try {
            payload.teamMembers = JSON.parse(payload.teamMembers);
        } catch (e) {
            payload.teamMembers = []; 
        }
    }
    if (typeof payload.captain === 'string') {
        try {
            const capData = JSON.parse(payload.captain);
            payload = { ...payload, ...capData };
        } catch (e) {
        }
    }
    if (req.file) {
        payload.paymentScreenshot = req.file.path;
    }
    if (payload.capRoll === "" || payload.capRoll === "undefined" || payload.capRoll === "null") {
        payload.capRoll = undefined;
    }

    return payload;
}

async function sendTelegramNotification(data) {
    const BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
    if (!BOT_TOKEN) {
        console.error("❌ Telegram Token is missing in .env file");
        return;
    }

    const CHAT_IDS = [
        process.env.TELEGRAM_GROUP_ID, 
        process.env.TELEGRAM_ADMIN_ID
    ].filter(id => id); 

    const isExternal = data.participantType === "EXTERNAL";
    const filename = data.paymentScreenshot ? path.basename(data.paymentScreenshot) : 'N/A';
    const header = isExternal 
        ? "🚨 <b>EXTERNAL REGISTRATION (VERIFY PAYMENT)</b> 💰" 
        : "✅ <b>New Internal Registration</b>";

    const message = `
${header}
━━━━━━━━━━━━━━━━━━━
<b>📌 Event:</b> ${data.eventName}
<b>🛡️ Team:</b> ${data.teamName}
<b>👤 Captain:</b> ${data.capName}
<b>📞 Phone:</b> ${data.capPhone}
${data.capRoll ? `<b>🎓 Roll:</b> ${data.capRoll}` : ''}
${isExternal ? `<b>📂 File:</b> ${filename}` : ''}
━━━━━━━━━━━━━━━━━━━
`;
    const promises = CHAT_IDS.map(id => 
        axios.post(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
            chat_id: id,
            text: message,
            parse_mode: 'HTML'
        }).catch(err => console.error(`Failed to notify ${id}:`, err.message))
    );

    await Promise.all(promises);
}