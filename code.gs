/**https://examblog64.krooluang.com/ */

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(config[6][1])
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

var ss = SpreadsheetApp.getActive()  
var sht1 = ss.getSheetByName("Data"); //ชื่อชีตข้อมูล1
var sht2 = ss.getSheetByName("setting")
var config = sht2.getDataRange().getDisplayValues(); //ชื่อชีตข้อมูล2
var idJPG = config[1][1]; //ไอดีโฟลเดอร์ JPG เก็บภาพ JPG
var idSlide = config[2][1]; //ไอดี Slide
var pic = config[3][1]; //ไอดีภาพ
var color = config[4][1]; //สี
var dev = config[5][1]; //ผู้พัฒนา
var title = config[6][1]; //ชื่อเรื่อง
var credit = config[9][1]; //ชื่อเรื่อง

/**Set แถวที่ต้องการรัน*/
function setData(obj){
  sht2.getRange(8,2).setValue(obj.start) 
  sht2.getRange(9,2).setValue(obj.end)
  return obj
}

/**สร้างเกียรติบัตร */
function runJPG(){
  var start = config[7][1]
  var end = config[8][1]

  for(var i=start; i<=end; i++){
    var checkrun = sht1.getRange(i,5).getValue()
    if(checkrun != 'สร้างแล้ว'){

var jpgpFolder = DriveApp.getFolderById(idJPG); 

    var data0 = sht1.getRange(i,1).getValue()//เลขที่
    var data1 = sht1.getRange(i,2).getValue()//ชื่อ สกุล
    var data2 = sht1.getRange(i,3).getValue()//ตำแหน่ง

    var file = DriveApp.getFileById(idSlide); //ตัวแปรไฟล์ไอดีสไลด์

var copyFile = file.makeCopy(data1); //สร้างไฟล์สำเนาลงในโฟลเดอ Temp
var copyId = copyFile.getId(); //เก็บไอดีของสไลด์ที่สำเนาใหม่
var copyDoc = SlidesApp.openById(copyId); //เปิดไฟล์สำเนาใหม่
  
copyDoc.replaceAllText("{เลขที่}",data0) //แทนที่เลขที่
copyDoc.replaceAllText("{ชื่อสกุล}",data1) //แทนทีชื่อ สกุล
copyDoc.replaceAllText("{ตำแหน่ง}",data2) //ตำแหน่ง

copyDoc.saveAndClose(); //ปิดไฟล์สำเนา

var PJName = data1; //กำหนดชื่อไฟล์ jpg

//===========ส่วนการสร้างไฟล์ JPG =========================
var newJPGFile="" //กำหนดค่าว่างของชื่อไฟล์
var ucjpg =""
const srcSlides = copyDoc.getSlides();
srcSlides.forEach((s) => {
    const url = Slides.Presentations.Pages.getThumbnail(copyId, s.getObjectId(), {"thumbnailProperties.mimeType": "PNG"}).contentUrl;
    const blob = UrlFetchApp.fetch(url).getAs(MimeType.JPEG);        
    newJPGFile = jpgpFolder.createFile(blob.setName(PJName));
    var nx1=newJPGFile.getId() //ไอดีของไฟล์ภาพ jpg
     ucjpg= "https://lh5.googleusercontent.com/d/"+[nx1]
    
  });
 
   sht1.getRange(i,4).setValue(ucjpg) 
   sht1.getRange(i,5).setValue('สร้างแล้ว')

copyFile.setTrashed(true)
    }
 }
}

/**ดึงไฟล์ */
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

/**แสดงตาราง */
function getData(){
  var data = sht1.getDataRange().getDisplayValues()
  data.shift()
  return data
}

/**Login */
function loginData(obj){
  var ls = ss.getSheetByName('Admin')
  var data = ls.getDataRange().getDisplayValues()
  var output = data.find(r=>r[0]+r[1] == obj.username+obj.password)
  return output
}

/**แสดงรูปภาพ */
function getPic(){
  return data = sht2.getRange('B4').getDisplayValues() //ชื่อชีตข้อมูล2
}




