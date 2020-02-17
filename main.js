const electron = require('electron');
const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;


const puppeteer = require('puppeteer');
var concurentPup;
var fs = require('fs');
var xl = require('excel4node');
var wb;
var ws;
var xlStyleSmall, xlStyleBig;
let fileNametxt = "";
let mainWindow;
var inputPhoneNumberArray = [];
var delayInMilliseconds = 10000;
var exPath = '';
var directionToSource = "";
var limitRequest = 15;
var currentMegre = 0;
var noService = "Thuê bao này hiện không sử dụng dịch vụ nào";
var wrongNumber = "Số điện thoại này bị sai";
var cIII = 0;
var startStartIndex = 0;
var crawling = false;
var fileNamexlxs = "";
var threshHoldeCount = 10;
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800, height: 600, webPreferences: {
            nodeIntegration: true
        }
    });

    mainWindow.loadURL(`file://${__dirname}/index.html`);

    // Build menu from template
    const mainMenu = Menu.buildFromTemplate(mainMenuTemplate);
    // Insert menu
    Menu.setApplicationMenu(mainMenu);

    mainWindow.on('closed', function () {
        mainWindow = null;
    })
}

app.on('ready', createWindow);

ipcMain.on('crawl:do', async function (e, item) {
    ////console.log(e, item);
    delayInMilliseconds = item == null ? 10000 : item;
    //console.log("delayInMilliseconds", delayInMilliseconds,"directionToSource",directionToSource);
    if (directionToSource == "" || directionToSource == null) {
        await chooseSource(readFile, prepareExxcel, doCrawl);
        //await readFile();
        //await doCrawl();
    } else {
        //if (crawling == false) {
        //await readFile();
        ///console.log("do crawl");
        prepareExxcel(doCrawl);
        //crawling = true;
        //}
    }

})

ipcMain.on("crawl:openFile", async function (e, item) {
    chooseSource(readFile, prepareExxcel, nothing);
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
})

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow();
    }
})

async function chooseSource(callback1, callback2, callback3) {
    dialog.showOpenDialog({
        title: "Chọn đường dẫn tới file text chứa danh sách số điện thoại",
        properties: ['openFile', 'multiSelections']
    }, function (files) {
        if (files !== undefined) {
            // handle files
        }
    }).then(async (result) => {
        if (!result.filePaths[0].endsWith(".txt")) {
            await mainWindow.webContents.send('crawl:error_choose_not_txt', true);
        } else {
            directionToSource = result.filePaths[0];
            //console.log(result.filePaths);
            await mainWindow.webContents.send('crawl:error_choose_not_txt', false);
            callback1(callback2,callback3);
        }
    }).catch(err => {
        ////console.log(err);
    });
};

function chooseGoogelPath() {
    dialog.showOpenDialog({
        title: "Chọn đường dẫn tới Google Chrome",
        properties: ['openFile', 'multiSelections']
    }, function (files) {
        if (files !== undefined) {
            // handle files
        }
    }).then(async (result) => {
        if (!result.filePaths[0].endsWith("chrome.exe")) {
            await mainWindow.webContents.send('crawl:error_choose_not_chrome', true);
        } else {
            exPath = result.filePaths[0];
            //console.log(result.filePaths);
            await mainWindow.webContents.send('crawl:error_choose_not_chrome', false);
        }
    }).catch(err => {
        ////console.log(err);
    });
};

// Create menu template
const mainMenuTemplate = [
    // Each object is a dropdown
    {
        label: 'Chức năng',
        submenu: [
            {
                label: 'Chọn tệp chứa danh sách điện thoại',
                accelerator: process.platform == 'darwin' ? 'Command+F' : 'Ctrl+F',
                click() {
                    // if (crawling == false) {
                    chooseSource(readFile, prepareExxcel, nothing);
                    // }
                }
            },
            {
                label: 'Chọn đường dẫn tới Google Chrome',
                accelerator: process.platform == 'darwin' ? 'Command+G' : 'Ctrl+G',
                click() {
                    chooseGoogelPath();
                }
            },
            {
                label: 'Thoát',
                accelerator: process.platform == 'darwin' ? 'Command+Q' : 'Ctrl+Q',
                click() {
                    app.quit();
                }
            }
        ]
    }
];

function nothing(){
    
}

function readFile(callback1, callback2) {
    fs.readFile(directionToSource, 'utf-8', async (err, data) => {
        let arraySourceFileName = directionToSource.split("\\");
        fileNametxt = arraySourceFileName[arraySourceFileName.length - 1];
        fileNametxt = fileNametxt.replace('.txt', '');
        //console.log("file name", fileNametxt);
        if (err) {
            ////console.log("An error ocurred reading the file :" + err.message);
            await mainWindow.webContents.send('crawl:read_error', fileNametxt);
            return;
        }
        // Change how to handle the file content
        if (data == '' || data == null) {
            await mainWindow.webContents.send('crawl:read_error_null', fileNametxt);
        } else {
            let tResult = data.split("\n");
            inputPhoneNumberArray = [];
            tResult.forEach(element => {
                inputPhoneNumberArray.push(element);
            });
            cTotal = inputPhoneNumberArray.length;
            //console.log(inputPhoneNumberArray);

            callback1(callback2);
        }
    });
}

async function writeToFileXLSX() {
    //console.log("file name xlsx",fileNamexlxs);
    await wb.write(fileNamexlxs);
}

function writeToXcell(x, y, title) {
    // if (y < 6) {
    //     ws.cell(x, y).string(title).style(xlStyleSmall);
    // } else {
    //console.log("Ghi vao o ", x, y, "gia tri", title);
    if (y > 10) {
        ws.cell(x, y).string(title).style(xlStyleBig);//.comment({height: '50pt'});
    } else {
        ws.cell(x, y).string(title).style(xlStyleSmall);//.comment({height: '50pt'});
    }
    // }
}

function writeToXcellMerge(x1, y1, x2, y2, title) {
    //console.log("merge", x1, y1, "to", x2, y2,title);
    if (title == noService) {
        ws.cell(x1, y1, x2, y2, true).string(title).style(xlStyleBig).comment({ height: '50pt' });
    } else if (title.endsWith(wrongNumber)) {
        ws.cell(x1, y1, x2, y2, true).string(title).style(xlStyleBig).comment({ height: '50pt' });
    } else {
        ws.cell(x1, y1, x2, y2, true).string(title).style(xlStyleSmall).comment({ height: '50pt' });
    }
}

async function prepareExxcel(callback){
        wb = new xl.Workbook();
        ws = wb.addWorksheet('vinaphone');
        ws.column(1).setWidth(5);
        ws.column(2).setWidth(25);
        ws.column(3).setWidth(25);
        ws.column(4).setWidth(25);
        ws.column(5).setWidth(20);
        ws.column(6).setWidth(28);
        ws.column(7).setWidth(20);
        ws.column(8).setWidth(20);
        ws.column(9).setWidth(20);
        ws.column(10).setWidth(20);
        ws.column(11).setWidth(60);
        ws.column(12).setWidth(60);
    
        xlStyleSmall = wb.createStyle({
            alignment: {
                vertical: ['center'],
                horizontal: ['center'],
                wrapText: true,
            },
            font: {
                name: 'Arial',
                color: '#324b73',
                size: 12,
            }
        });
    
        xlStyleBig = wb.createStyle({
            alignment: {
                vertical: ['center'],
                wrapText: true,
            },
            font: {
                name: 'Arial',
                color: '#324b73',
                size: 12,
            }
        });
    
        let cTimee = new Date();
    
        fileNamexlxs = "(" + cTimee.getHours() + " Gio -" + cTimee.getMinutes() + " Phut Ngay " + cTimee.getDate() + " Thang " + cTimee.getMonth() + " Nam " + cTimee.getFullYear() + ")   " + fileNametxt + ".xlsx";
    
        await mainWindow.webContents.send('crawl:read_sucess', fileNametxt+" "+fileNamexlxs);
    
        var header = ["STT", "Họ và tên", "Số điện thoại", "Loại thuê bao", "Tỉnh/Thành phố", "Số tiền trong tài khoản chính", "Dịch vụ đăng ký trên hệ thống VC", "STT", "DỊCH VỤ", "Gói cước", "Giá cước", "Mổ tả chung", "Đối tượng"];
        currentMegre = 0;
    
        for (let i = 0; i < header.length; i++) {
            if (i <= 5) {
                writeToXcellMerge(1, Number.parseInt(i) + 1, 2, Number.parseInt(i) + 1, header[i]);
            }
            else if (i == 6) {
                writeToXcellMerge(1, 7, 1, 12, header[i]);
            }
            else {
                writeToXcell(2, Number.parseInt(i), header[i]);
            }
        }    
        startStartIndex = 0;
        callback();
}

function doCrawl() {
    //console.log("concurentPup", concurentPup != null);
    if (concurentPup != null) {
        concurentPup = null;
        doCrawl();
    } else {
        mainWindow.webContents.send('crawl:hideBTN', true) ;
        concurentPup = puppeteer.launch({ headless: true, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
            const page = await browser.newPage();

            await page.goto('https://daily.vinaphone.com.vn/');
            page.setViewport({ width: 2600, height: 3000 });

            page.on('dialog', async dialog => {
                await mainWindow.webContents.send('crawl:incorrect_number', inputPhoneNumberArray[cIII]);
                //await dialog.dismiss();
                await browser.close();
                concurentPup = null;
                dialog.dismiss();
                startStartIndex = cIII + 1;
                await writeToXcell(cIII + 3 + currentMegre, 1, startStartIndex + "");
                await writeToXcellMerge(cIII + 3 + currentMegre, 2, cIII + 3 + currentMegre, 12, inputPhoneNumberArray[cIII] + " - " + wrongNumber);
                ws.row(cIII + 3 + currentMegre).setHeight(80);
                await doCrawl();
            });

            await page.click('#btn-alert1 .effect-sadie');

            await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #username1', el => el.value = 'dangky41');
            await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #password1', el => el.value = '858382');

            await page.click('#popupAlert1 #report .clearfix #form-login .from-login .form-row .button');

            await page.waitForNavigation({ waitUntil: 'networkidle0' })

            //await page.click('.sidebar .antiScroll .antiscroll-inner .antiscroll-content .sidebar_inner #side_accordion .accordion-group .accordion-heading .accordion-toggle .icon-6');

            //await page.click('.sidebar .antiScroll .antiscroll-inner .antiscroll-content .sidebar_inner #side_accordion .accordion-group #collapseSix');


            await page.goto('https://daily.vinaphone.com.vn/portal/pcm!executeSearchSubscriber');

            const start = async () => {
                await asyncForEach(inputPhoneNumberArray, startStartIndex, async (element, index) => {
                    cIII = index;
                    if (startStartIndex == 0) {
                        if (index == 0) {
                            await page.$eval('.main_content #searchPhone', (el, value) => el.value = value, element);
                        } else {
                            await page.$eval('.main_content #Pcmp050Form_pcmp050Model_phoneNumber', (el, value) => el.value = value, element);
                        }
                    }
                    else {
                        try {
                            await page.$eval('.main_content #searchPhone', (el, value) => el.value = value, element);
                        } catch (err) {
                            try {
                                await page.$eval('.main_content #Pcmp050Form_pcmp050Model_phoneNumber', (el, value) => el.value = value, element);
                            } catch (err) {

                            }
                        }

                    }

                    //await page.waitForFunction("document.querySelector('.marginB30') && document.querySelector('.marginB30').style.display != 'none'");

                    await page.click('.main_content .button');

                    await page.waitForNavigation({ waitUntil: 'networkidle0' })

                    let arrayName = await page.$$('.marginB30 table.table td');
                    let bodyFileTrCountMoreThan6 = 0;
                    let countMegre = 0;
                    let newcolumn = 6;
                    let loopMerge = currentMegre;
                    let itemArray = [];
                    itemArray.push(index + 1);
                    let currentSerrvice = "";
                    //console.log("arrayName", arrayName.length);
                    for (let i = 0; i < arrayName.length; i++) {
                        if (i < 8) {
                            if (i % 2 === 1) {
                                //  writeToXcell(index + 3, i + 1, await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                                itemArray.push(await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                            }
                            if (i == 1) {
                                itemArray.push(inputPhoneNumberArray[index]);
                            }
                        } else {
                            bodyFileTrCountMoreThan6++;
                            newcolumn++;
                            await writeToXcell(index + 3 + loopMerge, newcolumn, await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                            currentSerrvice += "add";
                            if (bodyFileTrCountMoreThan6 + 1 === 7) {
                                newcolumn = 6;
                                countMegre++;
                                loopMerge++;
                                bodyFileTrCountMoreThan6 = 0;
                            }

                        }

                    }


                    if (currentSerrvice == "") {
                        currentSerrvice = noService;
                        await writeToXcellMerge(index + 3 + currentMegre, 7, index + 3 + currentMegre, 12, currentSerrvice);
                        countMegre = 1;
                    }
                    for (let i = 0; i < itemArray.length; i++) {
                        await writeToXcellMerge(index + 3 + currentMegre, Number.parseInt(i + 1), index + 3 + currentMegre + countMegre - 1, Number.parseInt(i + 1), typeof itemArray[i] == "number" ? itemArray[i] + "" : itemArray[i]);
                    }

                    ws.row(index + 3 + currentMegre).setHeight(50);

                    currentMegre += countMegre - 1;

                });

                // let cTimee = new Date();

                // await wb.write("(" + cTimee.getHours() + " Gio -" + cTimee.getMinutes() + " Phut Ngay " + cTimee.getDate() + " Thang " + cTimee.getMonth() + " Nam " + cTimee.getFullYear() + ")   " + fileNametxt + ".xlsx");

                //lần chạy cuối cùng
                await writeToFileXLSX();

                await browser.close();

                await mainWindow.webContents.send('crawl:result', true);

                concurentPup = null;
                //crawling = false;
            }

            start();


        }).catch(async (err) => {
            //console.log("pupperteer error ", err);
            await mainWindow.webContents.send('crawl:network_error', true);
        });
    }
}

function timer(ms) {
    return new Promise(res => setTimeout(res, ms));
}

async function asyncForEach(array, startIndex, callback) {
    let cIndex = 1;
    for (let index = startIndex; index < array.length; index++) {

        await callback(array[index], index);
        //console.log("xong ", cIII + 1, " = " + array[cIII]);
        // mỗi lần ghi đến file thứ limitRequest , ngắt luồng đi delayInMilliseconds
        // if (index == cIndex * limitRequest - 1 && index < array.length - 1) {
        //     cIndex++;
        //     await timer(delayInMilliseconds);
        // }
        await mainWindow.webContents.send('crawl:onrunning', (index + 1) + " " + inputPhoneNumberArray.length);

        if (index % threshHoldeCount === 0 && index > 0) {
            await writeToFileXLSX();
        }

        //crawl xong 1 số -> nghỉ await timer(delayInMilliseconds);
        if (index < array.length - 1) {
            await timer(delayInMilliseconds);
        }
    }
}
