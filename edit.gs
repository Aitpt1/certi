/**แก้ไขและรันเกียรติบัตรใหม่ */
function editData(obj){
  var data = sht1.getDataRange().getDisplayValues()
  var id = data.map(r=>r[0])
  var index = id.indexOf(obj.numid)

  sht1.getRange(index+1,2).setValue(obj.fname)
  sht1.getRange(index+1,3).setValue(obj.position)
     
var jpgpFolder = DriveApp.getFolderById(idJPG); //สร้างตัวแปรเปิดการเชื่อมกับโฟลเดอ JPG
var file = DriveApp.getFileById(idSlide); //ตัวแปรไฟล์ไอดีสไลด์
var copyFile = file.makeCopy(obj.fname); //สร้างไฟล์สำเนาลงในโฟลเดอ Temp
var copyId = copyFile.getId(); //เก็บไอดีของสไลด์ที่สำเนาใหม่
var copyDoc = SlidesApp.openById(copyId); //เปิดไฟล์สำเนาใหม่
  
copyDoc.replaceAllText("{เลขที่}",obj.numid) //แทนที่เลขที่
copyDoc.replaceAllText("{ชื่อสกุล}",obj.fname) //แทนทีชื่อ สกุล
copyDoc.replaceAllText("{ตำแหน่ง}",obj.position) //ตำแหน่ง

copyDoc.saveAndClose(); //ปิดไฟล์สำเนา

var PJName = obj.fname; //กำหนดชื่อไฟล์ jpg

//===========ส่วนการสร้างไฟล์ JPG =========================
var newJPGFile="" //กำหนดาค่าว่างของชื่อไฟล์
var ucjpg =""
const srcSlides = copyDoc.getSlides();
srcSlides.forEach((s, i) => {
    const url = Slides.Presentations.Pages.getThumbnail(copyId, s.getObjectId(), {"thumbnailProperties.mimeType": "PNG"}).contentUrl;
    const blob = UrlFetchApp.fetch(url).getAs(MimeType.JPEG);        
    newJPGFile = jpgpFolder.createFile(blob.setName(PJName));
    var nx1=newJPGFile.getId() //ไอดีของไฟล์ภาพ jpg
     ucjpg= "https://lh5.googleusercontent.com/d/"+[nx1]
    
  });
 
   sht1.getRange(index+1,4).setValue(ucjpg) 
   sht1.getRange(index+1,5).setValue('สร้างแล้ว')

copyFile.setTrashed(true)

  data = sht1.getRange(index+1,1,1,sht1.getLastColumn()).getDisplayValues()

  return data
}

