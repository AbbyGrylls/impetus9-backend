const mongoose=require('mongoose');
const crypto = require('crypto');
const Schema=mongoose.Schema
const mobileRegex=/^[6-9]\d{9}$/;
const rollRegex=/^[0-9]{4}[A-Z]{3}[0-9]{3}$/;
const teamMemSchema=new Schema({
    memName:{
        type:String,
        required:true,
        trim:true,
    },
    memPhone:{
        type:String,
        required:true,
        match:[mobileRegex,'Invalid Mobile Number. Hint - Indian mobile numbers only']
    },
    memRoll:{
        type:String,
        required:function(){
            return this.parent().participantType==="INTERNAL";
        },
        uppercase:true,
        trim:true,
        match:[rollRegex,'Invalid Roll Number. Hint - Only Uppercases (eg., 2023MEB025)']
    }
},{_id:false});

const registerSchema=new Schema({
    eventName:{
        type:String,
        required:true,
        trim:true,
        index:true
    },
    teamName:{
        type:String,
        required:true,
        trim:true,
    },
    capName:{
        type:String,
        required:true,
        trim:true
    },
    capPhone:{
        type:String,
        required:true,
        match:[mobileRegex,'Invalid Mobile Number. Hint - Indian mobile numbers only']
    },
    capRoll:{
        type:String,
        required:function(){
            return this.participantType==="INTERNAL";
        },
        uppercase:true,
        trim:true,
        match:[rollRegex,'Invalid Roll Number. Hint - Only Uppercases (eg., 2023MEB025)']
    },
    teamMembers:{
        type:[teamMemSchema],
        default:[]
    },
    participantType:{
        type:String,
        enum:["INTERNAL","EXTERNAL"],
        required:true
    },
    captchaToken:{
        type:String,
        required:true
    },
    deviceFingerprint:{
        type:String,
        index:true,
    },
    receiptId: {
        type:String,
        unique:true,
    }
   
},{timestamps:true});

registerSchema.index({eventName:1,capPhone:1},{unique:true});
registerSchema.index(
    { eventName:1,capRoll:1}, 
    { unique:true,partialFilterExpression:{capRoll:{$exists:true,$type:"string"}}}
);
registerSchema.pre("validate",function(next){
    const phoneSet=new Set();
    const rollSet=new Set();
    if (this.capPhone){
        if (phoneSet.has(this.capPhone)) return next(new Error("Duplicate phone number: Captain phone is repeated"));
        phoneSet.add(this.capPhone);
    }
    if (this.participantType==="INTERNAL" && this.capRoll) {
        if (rollSet.has(this.capRoll)) return next(new Error("Duplicate roll number: Captain roll is repeated"));
        rollSet.add(this.capRoll);
    }
    if (Array.isArray(this.teamMembers) && this.teamMembers.length > 0){
        for (const member of this.teamMembers){
            if (phoneSet.has(member.memPhone)){
                return next(new Error(`Duplicate phone number found for member ${member.memName}`));
            }
            phoneSet.add(member.memPhone);
            if (this.participantType==="INTERNAL" && member.memRoll){
                if (rollSet.has(member.memRoll)){
                    return next(new Error(`Duplicate roll number found for member ${member.memName}`));
                }
                rollSet.add(member.memRoll);
            }
        }
    }
    next();
});
registerSchema.pre("save", async function(next) {
    if (!this.isNew) return next();
    const Registration=mongoose.models["registerModel"] || this.constructor;
    const phonesToCheck=[this.capPhone];
    this.teamMembers.forEach(m=>phonesToCheck.push(m.memPhone));
    const rollsToCheck=[];
    if (this.participantType==="INTERNAL") {
        if (this.capRoll) rollsToCheck.push(this.capRoll);
        this.teamMembers.forEach(m=>{
            if (m.memRoll) rollsToCheck.push(m.memRoll);
        });
    }
    const orConditions=[
        { capPhone:{ $in: phonesToCheck }},
        { "teamMembers.memPhone":{ $in: phonesToCheck }}
    ];
    if (rollsToCheck.length>0){
        orConditions.push({capRoll:{ $in: rollsToCheck } });
        orConditions.push({"teamMembers.memRoll":{ $in: rollsToCheck } });
    }
    const conflict=await Registration.findOne({
        eventName:this.eventName,
        $or:orConditions
    }).select("_id").lean();
    if (conflict){
        return next(
            new Error("A team member or the captain is already registered in this event")
        );
    }
    next();
});
registerSchema.pre("save", async function (next) {
    if (!this.isNew || this.receiptId) return next();
    const cleanName=this.eventName.replace(/[^a-zA-Z0-9]/g, "");
    let prefix=cleanName.substring(0,3).toUpperCase()
    if(prefix.length<3){
        prefix=prefix.padEnd(3,'X');
    }
    let isUnique = false;
    let newId = "";
    while (!isUnique) {
        const randomString = crypto.randomBytes(3).toString('hex').toUpperCase();
        newId = `${prefix}-${randomString}`;
        const existing = await mongoose.models["registerModel"].findOne({ receiptId: newId });
        if (!existing) isUnique = true;
    }
    this.receiptId = newId;
    next();
});
module.exports=mongoose.model('registerModel',registerSchema)