let samples={};

// -------------------------
// TABS
// -------------------------

function openTab(event,name){

    document.querySelectorAll(".tab")
    .forEach(tab=>tab.classList.remove("active"));

    document.getElementById(name)
    .classList.add("active");


    document.querySelectorAll(".tab-button")
    .forEach(btn=>btn.classList.remove("active"));

    if(event){
        event.target.classList.add("active");
    }

    // Resize plots after tab becomes visible
    setTimeout(()=>{
        document
        .querySelectorAll(`#${name} .plot`)
        .forEach(plot=>{
            Plotly.Plots.resize(plot);
        });
    },100);

     if(name === "comparison"){
        buildComparisonTable();
        createComparison();
    }
}

openTab(null,"sample1");

document
    .querySelector(".tab-button")
    .classList.add("active");

// -------------------------
// READ EXCEL
// -------------------------
function readExcel(file,callback){
    let reader=new FileReader();
    reader.onload=function(e){
        let bytes=new Uint8Array(e.target.result);
        let workbook=XLSX.read(bytes,{type:"array"});
        let sheet=workbook.Sheets[workbook.SheetNames[1]];
        let data=XLSX.utils.sheet_to_json(sheet,{header:1});
        callback(data);
    }
    reader.readAsArrayBuffer(file);
}


// -------------------------
// PARSE TRIOS DATA
// -------------------------
function parseData(raw,intensity,rxnStart){
    // TRIOS Excel structure
    let headers = raw[1];
    let units = raw[2];
    let rows = raw.slice(3);

    let storageIndex = headers.indexOf("Storage modulus");
    let lossIndex = headers.indexOf("Loss modulus");
    let timeIndex = headers.indexOf("Step time");

    let time=[];
    let storage=[];
    let loss=[];
    let dosage=[];

    rows.forEach(row=>{
        let t = Number(row[timeIndex]);
        let s = Number(row[storageIndex]);
        let l = Number(row[lossIndex]);

        // remove empty rows
        if(!isNaN(t) && !isNaN(s) && !isNaN(l)){
            time.push(t);
            storage.push(s);
            loss.push(l);
            dosage.push(
                Math.max(t*intensity - intensity*rxnStart,0)
            );
        }
    });

    return {
        time,
        storage,
        loss,
        dosage,

        headers,
        units,

        // IMPORTANT
        // only keep valid rows
        rows: rows.filter(row=>
            !isNaN(
                Number(row[timeIndex])
            )
        ),

        rxnStart
    };
}


// -------------------------
// LOAD SAMPLE
// -------------------------
function loadSample(num){
    let file=document.getElementById("file"+num).files[0];
    if(!file){
        alert("Upload Excel file first");
        return;
    }
    let intensity=Number(document.getElementById("intensity"+num).value);
    let rxn=Number(document.getElementById("rxn"+num).value);
    readExcel(
        file,
        raw=>{
            let sample=parseData(raw,intensity,rxn);
            samples[num]=sample;
            samples[num].name = document.getElementById("sampleName"+num).value;


            samples[num].processed=true;
            console.log("Saved sample:",num,samples[num]);

            document.getElementById("dataDetails"+num).style.display = "block";
            document.getElementById("unitsDetails"+num).style.display = "block";
            document.getElementById("exportButton"+num).style.display = "inline-block";


            document.getElementById("plots"+num).innerHTML=
            `<div 
                class="plot-container">
                <div class="plot-card">
                <div
                id="timePlot${num}"
                class="plot">
                </div>
                </div>

                <div class="plot-card">
                <div
                id="dosePlot${num}"
                class="plot">
                </div>
                </div>
            </div>`;
            // `<div 
            //     class="plot-container">

            //     <div
            //     class="plot">PLOT 1
            //     </div>

            //     <div
            //     class="plot">PLOT 2
            //     </div>
            // </div>`;

            plotTime(sample,"timePlot"+num);
            plotDosage(sample,"dosePlot"+num);

            setTimeout(()=>{
                Plotly.Plots.resize(document.getElementById("timePlot"+num));
                Plotly.Plots.resize(document.getElementById("dosePlot"+num));
            },100);
        }
    );
}


// -------------------------
// PLOTS
// -------------------------

function plotTime(sample,id){
    Plotly.newPlot(
        id,
        [
            {
                x:sample.time,
                y:sample.storage,
                mode:"markers",
                name:"G'"
            },
            {
                x:sample.time,
                y:sample.loss,
                mode:"markers",
                name:"G''"
            }
        ],
        {
            autosize:true,
            // height:450,

            yaxis:{
                type:"log",
                title:"Modulus (MPa)",
                exponentformat:"power",
            },
            xaxis:{
                title:"Time (s)"
            },
            legend:{
                orientation:"h",
                y:-0.25,
                x:0.5,
                xanchor:"center"
            }
        },
        {responsive:true}
    );
}



function plotDosage(sample,id){
    Plotly.newPlot(
        id,
        [
            {
                x:sample.dosage,
                y:sample.storage,
                mode:"markers",
                name:"G'"
            },
            {
                x:sample.dosage,
                y:sample.loss,
                mode:"markers",
                name:"G''"
            }
        ],
        {
            autosize:true,
            // height:450,

            yaxis:{
                type:"log",
                title:"Modulus (MPa)",
                exponentformat:"power",
            },
            xaxis:{
                title:"Dosage (mJ/cm²)"
            },
            legend:{
                orientation:"h",
                y:-0.25,
                x:0.5,
                xanchor:"center"
            }
        },
        {responsive:true}
    );
}



//
// Tables
//
function displayUnits(num){
    let s=samples[num];
    let container=document.getElementById("units"+num);
    let html=
        `<table>
            <tr>

            <th>
            Column
            </th>

            <th>
            Units
            </th>

            </tr>
        `;
    s.headers.forEach((h,i)=>{
        let unit = s.units[i] ?? "";
        html+=
            `<tr>

                <td>
                ${h}
                </td>


                <td>
                ${unit}
                </td>


                </tr>

            `;
        });
    html+="</table>";
    container.innerHTML=html;
}


function showTable(num){
    console.log("showTable called",num);
    let s = samples[num];
    console.log("sample data:",s);
    let container = document.getElementById(
        "tableContainer"+num
    );
    console.log("container:",container);
    if(!container){
        alert(
        "Missing tableContainer"+num
        );
        return;
    }
    let html = 
        `<table border="1">
        <thead>
        <tr>
        `;

    s.headers.forEach(header=>{
        html += `<th>${header}</th>`;
    });

    html += 
        `
            <th>
            Dosage (mJ/cm²)
            </th>
            </tr>
            </thead>
            <tbody>
        `;

    for(let i=0;i<s.rows.length;i++){
        html += "<tr>";
        for(let j=0;j<s.headers.length;j++){
            html += `<td>${s.rows[i][j] ?? ""}</td>`;
        }
        html += `<td> ${s.dosage[i].toFixed(3)} </td>`;
        html += "</tr>";
    }
    html += `</tbody></table>`;
    container.innerHTML = html;
    console.log("table inserted");
}


function exportTable(num){

    let s = samples[num];

    if(!s){
        alert("No sample loaded");
        return;
    }


    let csv = [];

     // Add headers row
    csv.push(
        [...s.headers, "Dosage"].join(",")
    );


    // Add units row
    let unitsRow = [...s.units, "mJ/cm²"];

    csv.push(
        unitsRow.join(",")
    );


    // Add data rows
    for(let i=0;i<s.rows.length;i++){

        let row = [
            ...s.rows[i],
            s.dosage[i].toFixed(3)
        ];

        csv.push(row.join(","));
    }


    let blob = new Blob(
        ["\ufeff"+csv.join("\n")],
        {type:"text/csv;charset=utf-8"}
    );


    let link=document.createElement("a");

    link.href=URL.createObjectURL(blob);

    link.download=`sample${num}_processed_data.csv`;

    link.click();
}



// -------------------------
// INTERPOLATION
// -------------------------
function interpolate(x,y){
    return function(value){
        if(value <= x[0]){
            return y[0];
        }
        if(value >= x[x.length-1]){
            return y[y.length-1];
        }
        for(let i=0;i<x.length-1;i++){
            if(value >= x[i] && value <= x[i+1]){
                let fraction = (value-x[i])/ (x[i+1]-x[i]);
                return (y[i] + fraction*(y[i+1]-y[i]));
            }
        }
        return NaN;
    }
}


// -------------------------
// ROOT FINDING
// -------------------------

function findRoot(func,a,b){
    let fa = func(a);
    let fb = func(b);

    if(isNaN(fa) || isNaN(fb)){
        throw "Bounds outside data range";
    }

    if(fa*fb > 0){
        throw "No gel point found in selected bounds";
    }

    for(let i=0;i<100;i++){
        let mid=(a+b)/2;
        let fm=func(mid);
        if(Math.abs(fm)<1e-12){
            return mid;
        }

        if(fa*fm < 0){
            b=mid;
            fb=fm;
        }
        
        else{
            a=mid;
            fa=fm;
        }
    }
    return (a+b)/2;
}


// -------------------------
// GEL POINT
// -------------------------


function linearInterpolation(x,xData,yData){
    for(let i=0;i<xData.length-1;i++){
        if(
            x>=xData[i] &&
            x<=xData[i+1]
        )
        {
            let fraction =
            (x-xData[i])/
            (xData[i+1]-xData[i]);
            return (
                yData[i] +
                fraction *
                (yData[i+1]-yData[i])
            );
        }
    }
    return NaN;
}

function findRoot(func,a,b){
    let fa = func(a);
    let fb = func(b);
    if(fa*fb>0){
        throw "No gel point found in selected bounds";
    }

    for(let i=0;i<100;i++){
        let mid=(a+b)/2;
        let fm=func(mid);
        if(Math.abs(fm)<1e-10){
            return mid;
        }
        if(fa*fm<0){
            b=mid;
            fb=fm;
        }
        else{
            a=mid;
            fa=fm;
        }
    }
    return (a+b)/2;
}


function calculateGel(num){
    let s=samples[num];
    // if(!s){
    //     alert("Process sample first");
    //     return;
    // }
    if(!s || !s.processed){
        alert("Please process sample before calculating gel point");
        return;
    }
    let lower = Number(document.getElementById("lower"+num).value);
    let upper = Number(document.getElementById("upper"+num).value);
    let difference = function(t){
        return (linearInterpolation(t,s.time,s.storage) - linearInterpolation(t,s.time,s.loss));
    };
    let root;
    try{
        root=findRoot(difference,lower,upper);
    }
    catch(err){
        alert(err);
        return;
    }
    let gelTime = root - s.rxnStart;
    let gelDose = linearInterpolation(root,s.time,s.dosage);
    // save result like streamlit session state
    samples[num].gel={time:gelTime,dose:gelDose};
    // document.getElementById("result"+num).innerHTML= 
    // `
    //     <br>
    //     <h3>
    //     Gel Time:
    //     ${gelTime.toFixed(3)} s
    //     </h3>
        
    //     <h3>
    //     Gel Dose:
    //     ${gelDose.toFixed(3)} mJ/cm²
    //     </h3>
    //     <br>
    // `;
    document.getElementById("result"+num).innerHTML = 
    `
    <div class="result-card">

        <h3>Gel Point Results</h3>

        <div class="result-row">
            <span>Gel Time:</span>
            <span>${gelTime.toFixed(3)} </span>
            <span>s</span>
        </div>

        <div class="result-row">
            <span>Gel Dose:</span>
            <span>${gelDose.toFixed(3)} </span>
            <span>mJ/cm²</span>
        </div>

    </div>
    `;


    // create plot area if missing
    document.getElementById("gelPlot"+num).innerHTML="";
    document.getElementById("result"+num).innerHTML +=
    // `
    //     <div class="plot gel-box gel-wide">
    //         <div id="gelPlot${num}"></div>
    //     </div>
    // `;
        `
            
            <div class="gel-box">
            <div id="gelPlot${num}" class="plot"></div>
            </div>
            
        `;
    createGelPointPlot(num,gelTime,gelDose);
    // updateSummary(num, gelPoint, gelDosage);
}


function createGelPointPlot(num,gelTime,gelDose){
    let s=samples[num];
    let gelAbsoluteTime = gelTime+s.rxnStart;
    let interpTime=[];
    let interpStorage=[];
    let interpLoss=[];

    for(let i=0;i<200;i++){
        let t = Math.min(...s.time)+i*((Math.max(...s.time)-Math.min(...s.time))/199);
        interpTime.push(t);
        interpStorage.push(linearInterpolation(t,s.time,s.storage));
        interpLoss.push(linearInterpolation(t,s.time,s.loss));
    }

    let gelModulus = linearInterpolation(gelAbsoluteTime,s.time,s.storage);
    let traces=[
        {x:s.time,y:s.storage,mode:"markers",name:"G'"},
        {x:s.time,y:s.loss,mode:"markers",name:'G"'},
        {x:interpTime,y:interpStorage,mode:"lines",name:"G' interpolation",line:{dash:"dash"}},
        {x:interpTime,y:interpLoss,mode:"lines",name:'G" interpolation',line:{dash:"dash"}},
        {x:[gelAbsoluteTime],y:[gelModulus],mode:"markers",name:"Gel Point",marker:{size:14,symbol:"star",color:"yellow",line:{width:2,color:"black"}}},
        // {x:[s.rxnStart,s.rxnStart],y:[Math.min(...s.storage,...s.loss),Math.max(...s.storage,...s.loss)],
        //     mode:"lines",name:"Reaction Start",line:{dash:"dot"}}
    ];

    let layout={
        // title:`Gel Point = ${gelTime.toFixed(3)} s | ${gelDose.toFixed(3)} mJ/cm²`,
        xaxis:{title:"Time (s)"},
        yaxis:{
            title:"Modulus (MPa)",
            type:"log",
            exponentformat:"power",
        },

        shapes: [{
            type: "rect",
            x0:Math.min(...s.time),
            x1:s.rxnStart,
            y0:0,
            y1:1,
            yref:"paper",
            fillcolor:"rgba(128,128,128,0.2)",
            line: {width:0}
        }],
        
        legend:{
            orientation:"h",
            y:-0.25,
            x:0.5,
            xanchor:"center"
        },
        margin:{
            b:100
        },
        annotations:[{
            x:gelAbsoluteTime,
            y:Math.log10(gelModulus),
            text:`Gel Point:<br>${gelTime.toFixed(3)} s<br>${gelDose.toFixed(3)} mJ/cm²`,
            showarrow:false,
            // arrowhead:2,
            // arrowwidth:1,
            xshift:+75,
            yshift:-50,

            align:"left",
            bgcolor:"white",
            bordercolor:"black",
            borderwidth:1,
            borderpad:5,
        }]
    };
    
    Plotly.newPlot(
        "gelPlot"+num,
        traces,
        layout
    );
}

function updateSampleName(num){

    if(samples[num]){
        samples[num].name =
            document.getElementById("sampleName"+num).value.trim();

        createComparison();
        buildComparisonTable();
    }

}

function resetSample(num){

    console.log("reset called",num);

    if(samples[num]){
        delete samples[num].gel;
    }


    let result = document.getElementById("result"+num);
    if(result){
        result.innerHTML="";
    }


    let plots = document.getElementById("plots"+num);
    if(plots){
        plots.innerHTML="";
    }


    let data = document.getElementById("dataDetails"+num);
    if(data){
        data.style.display="none";
    }


    let units = document.getElementById("unitsDetails"+num);
    if(units){
        units.style.display="none";
    }


    let exportBtn = document.getElementById("exportButton"+num);
    if(exportBtn){
        exportBtn.style.display="none";
    }

    if(samples[num]){
        samples[num].processed=false;
        delete samples[num].gel;
    }

}



// -------------------------
// COMPARISON
// -------------------------


function createComparison(){
    let traces=[];
    // Object.keys(samples).forEach(num=>{
    //     let s=samples[num];
    Object.keys(samples).forEach(num=>{
        let s=samples[num];
        if(!s.processed){
            return;
        }
        traces.push({
            x:s.time,
            y:s.storage,
            mode:"markers",
            name:"G' "+s.name
        });
        traces.push({
            x:s.time,
            y:s.loss,
            mode:"markers",
            name:"G\" "+s.name
        });
    });

    document.getElementById("comparisonContainer").innerHTML =
        `
        
        <div class="plot-container">
        <div class="plot-card">
            <div class="plot">
                <div id="comparisonTimePlot"></div>
            </div>
        </div>
        <div class="plot-card">
            <div class="plot">
                <div id="comparisonDosePlot"></div>
            </div>
        </div>
        </div>
        
        ` ;
    
    Plotly.newPlot(
        "comparisonTimePlot",
        traces,
        {
            autosize:true,
            height:600,
            yaxis:{
                type:"log",
                title:"Modulus (MPa)",
                exponentformat:"power"
            },
            xaxis:{
                title:"Time (s)"
            },
            legend:{
                orientation:"h",
                y:-0.25,
                x:0.5,
                xanchor:"center"
            }
        },
        {responsive:true}
    );

    let doseTraces=[];
    Object.keys(samples).forEach(num=>{
        let s=samples[num];

        if(!s.processed){
            return;
        }
        doseTraces.push({
            x:s.dosage,
            y:s.storage,
            mode:"markers",
            name:"G' "+s.name
        });
        doseTraces.push({
            x:s.dosage,
            y:s.loss,
            mode:"markers",
            name:"G\" "+s.name
        });
    });

    Plotly.newPlot(
        "comparisonDosePlot",
        doseTraces,
        {
            autosize:true,
            height:600,
            yaxis:{
                type:"log",
                title:"Modulus (MPa)",
                exponentformat:"power"
            },

            xaxis:{
                title:"Dosage (mJ/cm²)"
            },
            legend:{
                orientation:"h",
                y:-0.25,
                x:0.5,
                xanchor:"center"
            }
        },
        {responsive:true}
    );

    let totalT=0;
    let totalD=0;
    let count=0;

    Object.values(samples).forEach(s=>{
        if(s.gel){
            totalT+=s.gel.time;
            totalD+=s.gel.dose;
            count++;
        }
    });

    // if(count>0){document.getElementById("average").innerHTML=
    //     `
    //     <h3>
    //     Average Gel Time: ${(totalT/count).toFixed(3)} s
    //     </h3>

    //     <h3>
    //     Average Dose: ${(totalD/count).toFixed(3)} mJ/cm²
    //     </h3>
    //     `;
    // }
}



function buildComparisonTable(){

    console.log("Building comparison table");
    console.log(samples);

    let html = "";

    let gelTimes = [];
    let gelDoses = [];


    // Always create Sample 1-3 rows
    for(let i = 1; i <= 3; i++){

        let gel = samples[i]?.gel;

        html += `
        <tr>
            <td>${samples[i]?.name || "Sample "+i}</td>
            <td>${gel ? gel.time.toFixed(3) : ""}</td>
            <td>${gel ? gel.dose.toFixed(3) : ""}</td>
        </tr>
        `;


        // Only include calculated values in average
        if(gel){
            gelTimes.push(gel.time);
            gelDoses.push(gel.dose);
        }
    }


    // Always show average row
    let avgTime = "";
    let avgDose = "";

    if(gelTimes.length > 0){

        avgTime =
            (gelTimes.reduce((a,b)=>a+b,0)/gelTimes.length)
            .toFixed(3);

        avgDose =
            (gelDoses.reduce((a,b)=>a+b,0)/gelDoses.length)
            .toFixed(3);
    }


    html += `
    <tr>
        <th>Average</th>
        <th>${avgTime}</th>
        <th>${avgDose}</th>
    </tr>
    `;


    document.getElementById("summaryBody").innerHTML = html;

}

function exportSummary(){

    let csv = [];

    // Header row
    csv.push(
        "Sample,Gel Time (s),Gel Dose (mJ/cm²)"
    );


    let totalTime = 0;
    let totalDose = 0;
    let count = 0;


    Object.keys(samples).forEach(num=>{

        let s = samples[num];

        if(s.gel){

            csv.push(
                `Sample ${num},${s.gel.time.toFixed(3)},${s.gel.dose.toFixed(3)}`
            );

            totalTime += s.gel.time;
            totalDose += s.gel.dose;
            count++;
        }

    });


    if(count>0){

        csv.push(
            `Average,${(totalTime/count).toFixed(3)},${(totalDose/count).toFixed(3)}`
        );

    }


    let blob = new Blob(
        [csv.join("\n")],
        {type:"text/csv;charset=utf-8;"}
    );


    let link=document.createElement("a");

    link.href=URL.createObjectURL(blob);

    link.download="gel_summary.csv";

    link.click();
}
