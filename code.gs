function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SINBON 任務助手')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function processForm(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0]; 
    
    let sinbonPN = "";     
    let customerPN = "";   
    let customerCode = ""; 
    let dataPath = "";     
    let demandDate = "";   
    let costEstDate = "";  
    let bomEstDate = "";   

    if (data.purpose === "請報價") {
      customerCode = data.field1; 
      customerPN = data.field2;   
      dataPath = data.field3;     
      demandDate = data.costReqDate;
    } 
    else if (data.purpose === "建系統") {
      customerCode = data.field1; 
      sinbonPN = data.field2;     
      dataPath = data.field3;     
      demandDate = data.costReqDate;
    }
    else if (data.purpose === "標準成本建制" || data.purpose === "一元成本建制") {
      customerCode = data.field1; 
      sinbonPN = data.field2;     
      dataPath = data.field3;     
      costEstDate = data.costReqDate;
      bomEstDate = data.bomReqDate;
    }
    else if (data.purpose === "文件回收") {
      sinbonPN = data.field2;     
      demandDate = data.costReqDate;
    }

    let rowData = [
      new Date(),           
      data.fromUser,        
      data.purpose,         
      sinbonPN,             
      customerPN,           
      customerCode,         
      dataPath,             
      demandDate,           
      costEstDate,          
      bomEstDate,           
      data.note || ""       
    ];

    sheet.appendRow(rowData);
    return "成功同步至試算表";
  } catch (e) {
    return "寫入失敗: " + e.toString();
  }
}
