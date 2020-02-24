
const electron = require('electron');
const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;


const puppeteer = require('puppeteer');
var concurentPup, concurentLogin;
var fs = require('fs');
var xl = require('excel4node');
var wb;
var ws;
var xlStyleSmall, xlStyleBig, xlStyleNone;
let fileNametxt = "";
let mainWindow;
var inputPhoneNumberArray = [];
var delayInMilliseconds = 10000;
var exPath = '';
var directionToSource = "";
var limitRequest = 15;
var currentMegre = 0;
var endRange = 11;
var serviceRange = 6;
var headeTitle = "header";
var noService = "Thuê bao này hiện không sử dụng dịch vụ nào";
var wrongNumber = "Số điện thoại này bị sai";
var wrongLogin = "Tên truy nhập hoặc mật khẩu chưa đúng. Vui lòng thử lại";
var cIII = 0;
var startStartIndex = 0;
var crawling = false;
var fileNamexlxs = "";
var threshHoldeCount = 10;
var defaultHeight = 35;
var rowSpacing = 2;
// var USERNAME = ['dangky41'];
// var PASSWORD = ['858382'];
var username = "";
var password = "";
let tResult = "";
var page, pageLogin;

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
        await chooseSource(readFile, specialForOnlyHitButton);
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
    chooseSource(readFile, nothing);
});

ipcMain.on("crawl:login", async function (e, item) {
    username = item.split(" ")[0];
    password = item.split(" ")[1];
    doLogin();
    // USERNAME.map((element, index) => {
    //     if (element == username) {
    //         PASSWORD.map((element, index) => {
    //             if (element == password) {
    //                 mainWindow.webContents.send('crawl:login_success', true);
    //             } else {
    //                 mainWindow.webContents.send('crawl:login_success', false);
    //             }
    //         });
    //     } else {
    //         mainWindow.webContents.send('crawl:login_success', false);
    //     }
    // });
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

async function chooseSource(callback1, callback2) {
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
            await mainWindow.webContents.send('crawl:error_choose_not_txt', false);
            callback1(callback2);
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
                    chooseSource(readFile, nothing);
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

function specialForOnlyHitButton() {
    prepareExxcel(doCrawl);
}

function nothing() {

}


function doLogin() {
    concurentLogin = puppeteer.launch({ headless: true, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
        pageLogin = await browser.newPage();
        await mainWindow.webContents.send('crawl:login_success', 2);
        await pageLogin.goto('https://daily.vinaphone.com.vn/');
        pageLogin.setViewport({ width: 2600, height: 3000 });
        pageLogin.on('dialog', async dialog => {
            //await mainWindow.webContents.send('crawl:incorrect_number', inputPhoneNumberArray[cIII]);
            if (dialog.message() == wrongLogin) {
                await mainWindow.webContents.send('crawl:login_success', 0);
            } else {
                await mainWindow.webContents.send('crawl:login_success', -1);
            }
            dialog.dismiss();
            await browser.close();
            concurentLogin = null;
        });

        await pageLogin.click('#btn-alert1 .effect-sadie');

        await pageLogin.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #username1', (el, value) => el.value = value, username);
        await pageLogin.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #password1', (el, value) => el.value = value, password);

        await pageLogin.click('#popupAlert1 #report .clearfix #form-login .from-login .form-row .button');

        await pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })

        await browser.close();
        concurentLogin = null;
        await mainWindow.webContents.send('crawl:login_success', 1);
    }).catch(err => { console.log("login error", err); });
}

function readFile(callback) {
    fs.readFile(directionToSource, 'utf-8', async (err, data) => {
        let arraySourceFileName = directionToSource.split("\\");
        let isNew = false;
        if (fileNametxt != "") {
            isNew = true;
        }
        fileNametxt = arraySourceFileName[arraySourceFileName.length - 1];
        fileNametxt = fileNametxt.replace('.txt', '');
        console.log("file name", fileNametxt);
        if (err) {
            ////console.log("An error ocurred reading the file :" + err.message);
            await mainWindow.webContents.send('crawl:read_error', fileNametxt);
            return;
        }
        // Change how to handle the file content
        if (data == '' || data == null) {
            await mainWindow.webContents.send('crawl:read_error_null', fileNametxt);
        } else {
            tResult = data.split("\n");
            if (isNew == true) {
                await mainWindow.webContents.send('crawl:read_sucess_new', fileNametxt);
            }
            else {
                await mainWindow.webContents.send('crawl:read_sucess_first_time', fileNametxt);
            }
            callback();
        }
    });
}

async function writeToFileXLSX() {
    //console.log("file name xlsx",fileNamexlxs);
    await wb.write(fileNamexlxs);
}

function writeToXcell(x, y, title) {
    //console.log("Ghi vao o ", x, y, "gia tri", title);

    title += "";

    if (title.startsWith("header")) {
        let ttitle = title.split("-")[1];
        title = JSON.stringify(title);
        //title.replace("\"/g","");
        ws.cell(x, y).string(ttitle).style(xlStyleNone);
    } else {
        if (title == wrongNumber) {
            if (y > 10) {
                ws.cell(x, y).string(title).style(xlStyleNone);//.comment({height: '50pt'});
            } else {
                ws.cell(x, y).string(title).style(xlStyleNone);//.comment({height: '50pt'});
            }
        } else if (title == noService) {
            if (y > 10) {
                ws.cell(x, y).string(title).style(xlStyleNone);//.comment({height: '50pt'});
            } else {
                ws.cell(x, y).string(title).style(xlStyleNone);//.comment({height: '50pt'});
            }
        }
        else {
            if (y > 10) {
                ws.cell(x, y).string(title).style(xlStyleBig);//.comment({height: '50pt'});
            } else {
                ws.cell(x, y).string(title).style(xlStyleSmall);//.comment({height: '50pt'});
            }
        }
    }
    // }
}

async function  writeNumberToCell(x,y,number){
    await ws.cell(x,y).number(number).style(xlStyleSmall);
  }

function writeToXcellMerge(x1, y1, x2, y2, title) {
    //console.log("Ghi vao o ", x1, y1, "den ", x2, y2, "gia tri", title);

    title += "";

    if (title.startsWith("header")) {
        let ttitle = title.split("-")[1];
        title = JSON.stringify(title);
        //title.replace("\"/g","");
        ws.cell(x1, y1, x2, y2, true).string(ttitle).style(xlStyleNone).comment({ height: '50pt' });
    }
    else if (title == noService) {
        ws.cell(x1, y1, x2, y2, true).string(title).style(xlStyleBig).comment({ height: '50pt' });
    } else if (title.endsWith(wrongNumber)) {
        ws.cell(x1, y1, x2, y2, true).string(title).style(xlStyleBig).comment({ height: '50pt' });
    } else {
        ws.cell(x1, y1, x2, y2, true).string(title).style(xlStyleSmall).comment({ height: '50pt' });
    }
}

function convertStringToInteger(str){
    let strArray = str.split(",");
    let temp = strArray.reduce(function(temp1, temp2){return temp1 + temp2}, "");
    return parseInt(temp);
}

async function prepareExxcel(callback) {

    inputPhoneNumberArray = [];
    tResult.forEach(element => {
        inputPhoneNumberArray.push(element);
    });
    cTotal = inputPhoneNumberArray.length;
    console.log(inputPhoneNumberArray);

    wb = new xl.Workbook();
    ws = wb.addWorksheet('vinaphone');
    ws.column(1).setWidth(5);
    ws.column(2).setWidth(25);
    ws.column(3).setWidth(25);
    ws.column(4).setWidth(25);
    ws.column(5).setWidth(20);
    ws.column(6).setWidth(28);
    ws.column(7).setWidth(20);
    ws.column(8).setWidth(47);
    ws.column(9).setWidth(20);
    ws.column(10).setWidth(40);
    ws.column(11).setWidth(85);
    //ws.column(12).setWidth(60);

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
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            name: 'Arial',
            color: '#324b73',
            size: 12,
        }
    });

    xlStyleNone = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            bold: true,
            name: 'Arial',
            color: '#324b73',
            size: 12,
        },
    });

    let cTimee = new Date();

    fileNamexlxs = "(" + cTimee.getHours() + " Gio -" + cTimee.getMinutes() + " Phut Ngay " + cTimee.getDate() + " Thang " + (cTimee.getMonth() + 1) + " Nam " + cTimee.getFullYear() + ")   " + fileNametxt + ".xlsx";

    var header = ["STT", "Họ và tên", "Số điện thoại", "Loại thuê bao", "Tỉnh/Thành phố", "Số tiền trong tài khoản chính", "Dịch vụ đăng ký trên hệ thống VC", "DỊCH VỤ", "Gói cước", "Giá cước", "Mổ tả chung", "Đối tượng"];
    currentMegre = 0;

    for (let i = 0; i < header.length; i++) {
        writeToXcell(1, Number.parseInt(i)+1, headeTitle + "-" + header[i]);
        // if (i <= 5) {
        //     writeToXcellMerge(1, Number.parseInt(i) + 1, 2, Number.parseInt(i) + 1, headeTitle + "-" + header[i]);
        // }
        // else if (i == 6) {
        //     writeToXcellMerge(1, 7, 1, endRange, headeTitle + "-" + header[i]);
        // }
        // else {
        //     writeToXcell(2, Number.parseInt(i), headeTitle + "-" + header[i]);
        // }
    }
    ws.row(1).setHeight(defaultHeight);
    startStartIndex = 0;
    mainWindow.webContents.send('crawl:hideBTN', true);
    callback();
}

function doCrawl() {
    console.log("concurentPup", concurentPup != null);
    if (concurentPup != null) {
        concurentPup = null;
        // browser.close();
        page.close();
        doCrawl();
    } else {
        concurentPup = puppeteer.launch({ headless: true, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
            page = await browser.newPage();

            await page.goto('https://daily.vinaphone.com.vn/');
            page.setViewport({ width: 2600, height: 3000 });
            page.on('dialog', async dialog => {
                await mainWindow.webContents.send('crawl:incorrect_number', inputPhoneNumberArray[cIII]);
                //await dialog.dismiss();
                await browser.close();
                concurentPup = null;
                dialog.dismiss();
                startStartIndex = cIII + 1;
                // await writeToXcell(cIII + rowSpacing + currentMegre, 1, startStartIndex + "");
                // await writeToXcellMerge(cIII + rowSpacing + currentMegre, 2, cIII + rowSpacing + currentMegre, endRange, inputPhoneNumberArray[cIII] + " - " + wrongNumber);
                // ws.row(cIII + rowSpacing + currentMegre).setHeight(80);
                await writeToXcell(cIII + rowSpacing, 1, startStartIndex + "");
                writeToXcell(cIII + rowSpacing, 3, inputPhoneNumberArray[cIII]);
                writeToXcell(cIII + rowSpacing, 4, "");
                writeToXcell(cIII + rowSpacing, 5, "");
                writeToXcell(cIII + rowSpacing, 6, "" + 0);

                await writeToXcell(cIII + rowSpacing, 8, wrongNumber);
                ws.row(cIII + rowSpacing).setHeight(defaultHeight);
                await doCrawl();
            });

            await page.click('#btn-alert1 .effect-sadie');

            await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #username1', (el, value) => el.value = value, username);
            await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #password1', (el, value) => el.value = value, password);

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
                    let userInfoCollumn = 4;//3 cột đầu đã được điền là STT,Họ và tên , Số điện thoại
                    let loopMerge = currentMegre;
                    let itemArray = [];
                    itemArray.push(index + 1);
                    let currentSerrvice = "";
                    //console.log("arrayName", arrayName.length);
                    //wait writeToXcell(index + rowSpacing + loopMerge, 0, index + 1));
                    await writeToXcell(index + rowSpacing, 1, (index + 1) + "");
                    await writeToXcell(index + rowSpacing, 3, inputPhoneNumberArray[index] + "");
                    for (let i = 1; i < arrayName.length; i++) {
                        if (i < 9) {
                            if (i == 8) {
                                continue;
                            }
                            if (i % 2 === 1) {
                                if (i == 7)/*có thể không trả kết quả số tiền*/ {
                                    console.log(await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                                    if (await (await arrayName[i].getProperty('innerHTML')).jsonValue() == null || await (await arrayName[i].getProperty('innerHTML')).jsonValue() == "") {
                                        // writeToXcell(index + rowSpacing, 6, 0 + "");
                                        writeNumberToCell(index + rowSpacing, 6, null);
                                    }
                                    else {
                                        var value = await (await arrayName[i].getProperty('innerHTML')).jsonValue()
                                        writeNumberToCell(index + rowSpacing, 6, convertStringToInteger(value));
                                    }
                                    itemArray.push(0);
                                    userInfoCollumn++;
                                } else if (i == 1) {//họ và tên
                                    itemArray.push(await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                                    writeToXcell(index + rowSpacing, 2, await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                                }
                                else {
                                    writeToXcell(index + rowSpacing, userInfoCollumn, await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                                    itemArray.push(await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                                    userInfoCollumn++;
                                }
                            }
                        } else {
                            bodyFileTrCountMoreThan6++;
                            newcolumn++;
                            //await writeToXcell(index + rowSpacing + loopMerge, newcolumn, await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                            await writeToXcell(index + rowSpacing, newcolumn, await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                            currentSerrvice += "add";
                            if (bodyFileTrCountMoreThan6 + 1 === serviceRange) {
                                break;
                                //newcolumn = 6;
                                //countMegre++;
                                // loopMerge++;
                                //bodyFileTrCountMoreThan6 = 0;
                            }

                        }

                    }


                    if (currentSerrvice == "") {
                        currentSerrvice = noService;
                        //await writeToXcellMerge(index + rowSpacing + currentMegre, 7, index + rowSpacing + currentMegre, 12, currentSerrvice);
                        //await writeToXcellMerge(index + rowSpacing, 7, index + rowSpacing, endRange, currentSerrvice);
                        await writeToXcell(index + rowSpacing, 8, currentSerrvice);
                        //countMegre = 1;
                    }
                    //ghi thoong tin
                    // for (let i = 0; i < itemArray.length; i++) {
                    //await writeToXcellMerge(index + rowSpacing + currentMegre, Number.parseInt(i + 1), index + rowSpacing + currentMegre + countMegre - 1, Number.parseInt(i + 1), typeof itemArray[i] == "number" ? itemArray[i] + "" : itemArray[i]);
                    //}

                    //ws.row(index + rowSpacing + currentMegre).setHeight(50);
                    ws.row(index + rowSpacing).setHeight(defaultHeight);

                    currentMegre += countMegre - 1;

                });

                // let cTimee = new Date();

                // await wb.write("(" + cTimee.getHours() + " Gio -" + cTimee.getMinutes() + " Phut Ngay " + cTimee.getDate() + " Thang " + cTimee.getMonth() + " Nam " + cTimee.getFullYear() + ")   " + fileNametxt + ".xlsx");

                console.log("end");
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
