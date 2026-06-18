const DRIVE_FOLDER_ID = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE";

const QUEUE_HEADERS = ["id","status","name","submitted","practice","value","adverse","email","phone","summary","fullText","formUrl","docsJson","specialTag","createdAt","reviewedAt"];
const REVIEW_HEADERS = ["timestamp","caseId","clientName","qualityLead","consultType","minRetainer","docsNeeded","attorneyNotes","referralInMind","referralNotes"];

function doPost(e){
  try{
    const data = JSON.parse(e.postData.contents || "{}");
    if(data.action === "listCases") return json(listCases_());
    if(data.action === "submitReview") return json(submitReview_(data));
    if(data.action === "addCase") return json(addCase_(data));
    return json({ok:false,error:"Unknown action"});
  }catch(err){return json({ok:false,error:String(err)});}
}
function doGet(){return json({ok:true,message:"Heather Review Portal backend is live."});}
function json(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function sheet_(name, headers){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.appendRow(headers);
  return sh;
}
function rowObj_(headers,row){const o={};headers.forEach((h,i)=>o[h]=row[i]);return o;}
function listCases_(){
  const sh=sheet_("Queue", QUEUE_HEADERS);
  const rows=sh.getDataRange().getValues();
  const headers=rows.shift() || [];
  const cases=rows.map(r=>rowObj_(headers,r)).filter(c=>String(c.status||"Pending").toLowerCase()!=="reviewed").map(c=>{
    let docs=[]; try{docs=c.docsJson?JSON.parse(c.docsJson):[]}catch(e){}
    return {id:c.id,status:c.status,name:c.name,submitted:c.submitted,practice:c.practice,value:c.value,adverse:c.adverse,email:c.email,phone:c.phone,summary:c.summary,fullText:c.fullText,formUrl:c.formUrl,docs:docs,specialTag:c.specialTag};
  });
  return {ok:true,cases};
}
function submitReview_(data){
  const r=sheet_("Reviews", REVIEW_HEADERS);
  r.appendRow([new Date(),data.caseId,data.clientName,data.qualityLead,data.consultType,data.minRetainer,data.docsNeeded,data.attorneyNotes,data.referralInMind,data.referralNotes]);
  const q=sheet_("Queue", QUEUE_HEADERS), values=q.getDataRange().getValues(), headers=values[0];
  const idCol=headers.indexOf("id")+1, statusCol=headers.indexOf("status")+1, reviewedAtCol=headers.indexOf("reviewedAt")+1;
  for(let i=1;i<values.length;i++){if(String(values[i][idCol-1])===String(data.caseId)){q.getRange(i+1,statusCol).setValue("Reviewed");q.getRange(i+1,reviewedAtCol).setValue(new Date());break;}}
  return {ok:true};
}
function addCase_(data){
  const q=sheet_("Queue", QUEUE_HEADERS);
  const id=slug_(data.name)+"-"+Utilities.getUuid().slice(0,8);
  const folder=(DRIVE_FOLDER_ID && DRIVE_FOLDER_ID.indexOf("PASTE_")!==0)?DriveApp.getFolderById(DRIVE_FOLDER_ID):DriveApp.getRootFolder();
  let formUrl="";
  if(data.intakeFile && data.intakeFile.data) formUrl=saveFile_(folder,data.intakeFile).getUrl();
  const docs=[];
  (data.docs||[]).forEach(f=>{const file=saveFile_(folder,f);docs.push({name:f.name,url:file.getUrl(),text:""});});
  q.appendRow([id,"Pending",data.name||"",data.submitted||"",data.practice||"",data.value||"",data.adverse||"",data.email||"",data.phone||"",data.summary||"",data.fullText||"",formUrl,JSON.stringify(docs),data.specialTag||"",new Date(),""]);
  return {ok:true,id:id};
}
function saveFile_(folder,fileObj){
  const bytes=Utilities.base64Decode(fileObj.data);
  const blob=Utilities.newBlob(bytes,fileObj.type||MimeType.BINARY,fileObj.name||"upload");
  return folder.createFile(blob);
}
function slug_(s){return String(s||"case").toLowerCase().replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"");}
