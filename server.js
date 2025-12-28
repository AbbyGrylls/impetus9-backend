require('dotenv').config()
const express=require('express')
const mongoose=require('mongoose')
const cors=require('cors')
const helmet=require('helmet');
const regisRoutes=require('./routes/registers')
const coordinatorRoutes=require('./routes/coordsLog')
const statusRoute=require('./routes/checkStatus')
const app = express()
app.use(helmet());
app.use(cors({
    origin:process.env.FRONTEND_URL || "http://localhost:3000",
    methods:["GET","POST"],
    credentials:true
}));
app.use(express.json())
app.use((req,res,next)=>{
    console.log(req.path,req.method)
    next()
})
app.use('/api/register',regisRoutes)
app.use('/api/coordinator',coordinatorRoutes);
app.use('/api/checkstatus',statusRoute)
mongoose.connect(process.env.MONGO_URI)
    .then(()=>{
        app.listen(process.env.PORT,()=>{
            console.log('connected to db: listeneing on port number',process.env.PORT)
        })
    })
    .catch((error)=>{
        console.log(error)
        process.exit(1)
    })
process.env