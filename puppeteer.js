const puppeteer = require('puppeteer');
//const path = require('path')

(async () => {
   
    let launchOptions = { headless: false, args: ['--start-maximized'] };

    const browser = await puppeteer.launch(launchOptions);
    const page = await browser.newPage();

   
    await page.setViewport({width: 1366, height: 768});
   
    const client = await page.target().createCDPSession();
  await client.send("Page.setDownloadBehavior", {
    behavior: "allow",
    downloadPath: "C:\\Users\\pragyas\\Desktop\\downloaded",
  });
  
    await page.goto(
        'https://translate.google.com/?sl=en&tl=hi&op=docs',
    
        {
            waitUntil: "networkidle2",
          }
    );

    await page.setDefaultTimeout(45000);

   const [fileChooser] = await Promise.all([
        page.waitForFileChooser(),
        page.click("label"),
      ]);

    const inputUploadHandle = await page.$('input[type=file]');

   //---------------file upload method 1----------------
    let fileToUpload = 'inputJsonValue.xlsx';

    // Sets the value of the file input to fileToUpload
    inputUploadHandle.uploadFile(fileToUpload);

//-------------------file upload method 2--------------------------
      await fileChooser.accept(["C:\\Users\\pragyas\\Downloads\\flattening json\\inputJsonValue.xlsx"]);
      await page.setDefaultTimeout(5000);

      const button = await page.waitForSelector(
      "#yDmH0d > c-wiz > div > div.WFnNle > c-wiz > div.R5HjH > c-wiz > div.oLbzv > c-wiz > div > div:nth-child(1) > div > div.ld4Jde > div > div > button > span");
      await button.evaluate((b) => b.click());
      
      const button2 = await page.waitForSelector(
        "#yDmH0d > c-wiz > div > div.WFnNle > c-wiz > div.R5HjH > c-wiz > div.oLbzv > c-wiz > div > div:nth-child(1) > div > div.ld4Jde > div > button > span.VfPpkd-vQzf8d",
        {
          visible: true,
          timeout: 0,
        }
      );
        await button2.evaluate((b) => b.click());
       
        
})();