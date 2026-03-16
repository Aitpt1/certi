//สร้างdialog ให้form
//https://www.youtube.com/watch?v=uhP_pO0CYfI

function createDialog() {
  const htmlService = HtmlService.createTemplateFromFile('index').evaluate().addMetaTag('viewport','width=device-width , initial-scale=1').setHeight(2000).setWidth(1000)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
   SpreadsheetApp.getUi().showModelessDialog(htmlService,'index')
}

/**สร้างเมนูใน Googlesheet */
function createMenu() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('📕 สร้างเกียรติบัตร')
  menu.addItem("✅ รันเกียรติบัตร","createSidebar")
  menu.addItem("✅ แสดงตาราง","createDialog")
  menu.addToUi() 
}

/**เปิดเมนูใน Googlesheet */
function onOpen(){
    createMenu()
   SpreadsheetApp.getUi().createMenu("✅ เปลี่ยนเลขไทย")
                         .addItem("✅ เปลี่ยนเป็นเลขไทย", "changeToThaiNumber")
                         .addToUi();
}

/**เปิดฟอร์ม */
function createSidebar() {
  const html = HtmlService.createTemplateFromFile('formRun').evaluate();
        html.setTitle('รันเกียรติบัตร');
        SpreadsheetApp.getUi().showSidebar(html)
}




