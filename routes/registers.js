const express=require('express')
const router=express.Router()
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
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, 'receipt-' + uniqueSuffix + path.extname(file.originalname));
    }
})
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
        const OPEN_EVENTS = ["Photo", "BGMI","IQ Ignition"]; 

        if (!OPEN_EVENTS.includes(eventName)) {
            return res.status(400).json({ 
                error: "Registration for this event is not opened." 
            });
        }
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
        if (participantType === "EXTERNAL") {
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
    // If a file was uploaded, we know it's an External Multipart request
    if (req.file) {
        try {
            const body = req.body;
            
            // 1. Parse nested JSON if it exists (e.g. if frontend sends 'captain' object)
            // If your frontend sends flat fields, this part is optional but good for safety.
            const captainData = body.captain ? JSON.parse(body.captain) : {};

            // 2. Construct the payload
            const payload = {
                ...body,
                ...captainData, // Flatten captain details to top level if necessary
                teamMembers: body.teamMembers ? JSON.parse(body.teamMembers) : [],
                paymentScreenshot: req.file.path,
                participantType: "EXTERNAL"
            };

            // 3. CRITICAL FIX: Convert empty strings to undefined for optional unique fields
            if (payload.capRoll === "") {
                payload.capRoll = undefined;
            }

            return payload;
        } catch (e) {
            throw new Error("Invalid Data Format: Could not parse FormData JSON");
        }
    }
    
    // For Internal Users
    const payload = req.body;
    // Good practice to apply the same fix for internal users just in case
    if (payload.capRoll === "") {
        payload.capRoll = undefined;
    }
    return payload;
}