function sendRecentEmails2Discord() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const webhookCell = sheet.getRange('B1');
  const lastUpdateAtCell = sheet.getRange('B2');
  const lastUpdatedAt = new Date(lastUpdateAtCell.getValue());
  const lastUpdatedAtUnixTime = Math.floor(lastUpdatedAt.getTime()/1000);

  const webhookUrl = webhookCell.getValue();

  const query = 'after:' + lastUpdatedAtUnixTime;
  const threads = GmailApp.search(query);
  const now = new Date();

  threads.forEach((thread) => {
    const messages = thread.getMessages();
    const payloads = messages.map((message) => {
      const sendBy = message.getFrom();
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const attatchmentNum = message.getAttachments().length;
      const attatchmentMes = attatchmentNum > 0 ? `[添付ファイル:${attatchmentNum}個]` : "";
      const content = `${sendBy}からのメールです！${attatchmentMes}`
      const payload = {
       content: content,
       embeds: [{
         title: subject,
         author: {
           name: sendBy,
         },
         description: body.substring(0,1024),
       }],
     }
     return {
       method: 'post',
       url: webhookUrl,
       contentType: 'application/json',
       payload: JSON.stringify(payload),
     }
    });
    UrlFetchApp.fetchAll(payloads);
    Utilities.sleep(50);
  });
  lastUpdateAtCell.setValue(now);
}
