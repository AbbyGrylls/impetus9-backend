const express=require('express')
const router=express.Router()
const registerModel=require('../models/registerModel')
const verifyCaptcha=require("../middleware/verifyCaptcha")
const rateLimiter=require("../middleware/rateLimiter")
const allowedEvents=require('../allowedEvents')
router.post("/",rateLimiter,async(req,res)=>{
    try{
        const{
            eventName,
            searchField,
            searchValue,
            captchaToken,
        }=req.body;
        if(!allowedEvents.has(eventName)) return res.status(400).json({error:"Invalid Event"});
        if(!searchValue || !searchField || !captchaToken) return res.status(400).json({error: "Missing Required Fields"});
        const captchaValid=await verifyCaptcha(captchaToken);
        if(!captchaValid) return res.status(400).json({error:"Captcha verification failed"});
        const query={eventName:eventName};
        if(searchField=="receiptID") query.receiptId=searchValue.trim();
        else if(searchField==='RollNo'){
            const roll=searchValue.trim().toUpperCase()
            query.$or = [
                { capRoll: roll },
                { "teamMembers.memRoll": roll }
            ];
        }else return res.status(400).json({error: " Invalid Search Field"});
        const registration=await registerModel.findOne(query).lean();
        if(!registration) return res.status(400).json({error:"No registration found with these details."})
        return res.status(200).json({
            success: true,
            data: {
                eventName: registration.eventName,
                teamName: registration.teamName,
                receiptId: registration.receiptId,
                participantType: registration.participantType,
                capName: registration.capName,
                capPhone: registration.capPhone,
                capRoll: registration.capRoll,
                teamMembers: registration.teamMembers
            }
        });
    }   catch(err){
        console.error("check status error:", err);
        return res.status(500).json({error:"Server Error"});
    }
})
module.exports=router