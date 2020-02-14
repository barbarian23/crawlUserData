const electron = require('electron');
const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;


const puppeteer = require('puppeteer');
var fs = require('fs');
var xl = require('excel4node');
var wb;
var ws;
var xlStyleSmall, xlStyleBig;

let mainWindow;
var inputPhoneNumberArray = [];
var header = ["STT", "Họ và tên", "Số điện thoại", "Loại thuê bao", "Tỉnh/Thành phố", "Dịch vụ đang sử dụng"];
const delayInMilliseconds = 60000;
var exPath = '';
var directionToSource = "";
var limitRequest = 15;
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
    //console.log(e, item);
    if (item) {
        if (directionToSource == "" || directionToSource == null) {
            chooseSource(readFile);
        } else {
            await readFile();
        }
    }
})

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

function chooseSource(callback) {
    dialog.showOpenDialog({
        title :"Chọn đường dẫn tới file text chứa danh sách số điện thoại",
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
            console.log(result.filePaths);
            await mainWindow.webContents.send('crawl:error_choose_not_txt', false);
            callback();
        }
    }).catch(err => {
        //console.log(err);
    });
};

function chooseGoogelPath() {
    dialog.showOpenDialog({
        title :"Chọn đường dẫn tới Google Chrome",
        properties: ['openFile', 'multiSelections']
    }, function (files) {
        if (files !== undefined) {
            // handle files
        }
    }).then( async (result) => {
        if (!result.filePaths[0].endsWith("chrome.exe")) {
            await mainWindow.webContents.send('crawl:error_choose_not_chrome', true);
        } else {
            exPath = result.filePaths[0];
            console.log(result.filePaths);
            await mainWindow.webContents.send('crawl:error_choose_not_chrome', false);
        }
    }).catch(err => {
        //console.log(err);
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
                    chooseSource();
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

function readFile() {
    fs.readFile(directionToSource, 'utf-8', async (err, data) => {
        if (err) {
            //console.log("An error ocurred reading the file :" + err.message);
            await mainWindow.webContents.send('crawl:read_error', true);
            return;
        }
        // Change how to handle the file content
        if (data == '' || data == null) {
            await mainWindow.webContents.send('crawl:read_error', false);
        } else {
            let tResult = data.split(",");
            inputPhoneNumberArray = [];
            tResult.forEach(element => {
                inputPhoneNumberArray.push(element);
            });
            cTotal = inputPhoneNumberArray.length;
            //console.log(inputPhoneNumberArray);

            wb = new xl.Workbook();
            ws = wb.addWorksheet('vinaphone');
            ws.column(1).setWidth(5);
            ws.column(2).setWidth(25);
            ws.column(3).setWidth(25);
            ws.column(4).setWidth(25);
            ws.column(5).setWidth(20);
            ws.column(6).setWidth(148);

            xlStyleSmall = wb.createStyle({
                alignment: {
                    wrapText: true,
                    horizontal: ['center'],
                    vertical: ['center']
                },
                font: {
                    name: 'Arial',
                    color: '#324b73',
                    size: 12,
                }
            });

            xlStyleBig = wb.createStyle({
                alignment: {
                    wrapText: true,
                    vertical: ['center']
                },
                font: {
                    name: 'Arial',
                    color: '#324b73',
                    size: 12,
                }
            });

            await doCrawl();
        }
    });
}

function writeToXcell(x, y, title) {
    if (y < 6) {
        ws.cell(x, y).string(title).style(xlStyleSmall);
    } else {
        ws.cell(x, y).string(title).style(xlStyleBig);
    }
}

var cIII = 0;
var cTotal = 0;

function doCrawl() {
    //console.log("123");
    for (let i = 0; i < header.length; i++) {
        writeToXcell(1, Number.parseInt(i) + 1, header[i]);
    }
    puppeteer.launch({ headless: true, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
        const page = await browser.newPage();
        page.on('dialog', async dialog => {
            await mainWindow.webContents.send('crawl:incorrect_number', inputPhoneNumberArray[cIII]);
        });
        await page.goto('https://daily.vinaphone.com.vn/');
        page.setViewport({ width: 1280, height: 2400 });

        await page.click('#btn-alert1 .effect-sadie');

        await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #username1', el => el.value = 'dangky41');
        await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #password1', el => el.value = '858382');

        await page.click('#popupAlert1 #report .clearfix #form-login .from-login .form-row .button');

        await page.waitForNavigation({ waitUntil: 'networkidle0' })


        //await page.click('.sidebar .antiScroll .antiscroll-inner .antiscroll-content .sidebar_inner #side_accordion .accordion-group .accordion-heading .accordion-toggle .icon-6');

        //await page.click('.sidebar .antiScroll .antiscroll-inner .antiscroll-content .sidebar_inner #side_accordion .accordion-group #collapseSix');


        await page.goto('https://daily.vinaphone.com.vn/portal/pcm!executeSearchSubscriber');

        var bodyFileExxcel = [];
        bodyFileExxcel.push(header);

        //Số điện thoại vào array này
        //var inputPhoneNumberArray = ['0944854975', '0946245467', '0944854975', '0946245467', '0944854975', '0946245467', '0944854975'];

        const start = async () => {
            await asyncForEach(inputPhoneNumberArray, async (element, index) => {
                cIII = index;
                //console.log("index", element, index);
                if (index == 0) {
                    await page.$eval('.main_content #searchPhone', (el, value) => el.value = value, element);
                } else {
                    await page.$eval('.main_content #Pcmp050Form_pcmp050Model_phoneNumber', (el, value) => el.value = value, element);
                }

                await page.click('.main_content .button');

                //await page.waitForFunction("document.querySelector('.marginB30') && document.querySelector('.marginB30').style.display != 'none'");

                await page.waitForNavigation({ waitUntil: 'networkidle0' })

                let arrayName = await page.$$('.marginB30 table.table td');
                let bodyFileTrCountMoreThan6 = 0;
                let itemArray = [];
                itemArray.push(index + 1);
                let currentSerrvice = "";
                for (let i = 0; i < arrayName.length; i++) {
                    if (i < 8) {
                        if (i % 2 === 1 && i != 7) {
                            itemArray.push(await (await arrayName[i].getProperty('innerHTML')).jsonValue());
                        }
                        if (i == 1) {
                            itemArray.push(inputPhoneNumberArray[index]);
                        }
                    } else {
                        bodyFileTrCountMoreThan6++;
                        if (bodyFileTrCountMoreThan6 === 1) {
                            currentSerrvice += "<<< Dịch vụ đang dùng thứ " + await (await arrayName[i].getProperty('innerHTML')).jsonValue() + ">>>\n";
                        } else {

                            currentSerrvice += "- " + await (await arrayName[i].getProperty('innerHTML')).jsonValue() + "\n ";
                            if (bodyFileTrCountMoreThan6 + 1 === 7) {
                                bodyFileTrCountMoreThan6 = 0;
                                currentSerrvice += "===========================================================================\n\n"
                            }
                        }
                    }

                }

                if (currentSerrvice == "") {
                    currentSerrvice = "Thuê bao này hiện không sử dụng dịch vụ nào"
                }
                itemArray.push(currentSerrvice);

                for (let i = 0; i < itemArray.length; i++) {
                    if (typeof itemArray[i] === "number") {
                        itemArray[i] += "";
                    }
                    writeToXcell(index + 2, i + 1, itemArray[i]);
                }

                //console.log("content\n", itemArray);

                bodyFileExxcel.push(itemArray);

                //console.log("excel will be\n", bodyFileExxcel);
               
            });

            //console.log('Đã crawl xong data');

            await wb.write('ketqua.xlsx');

            // var ws = XLSX.utils.aoa_to_sheet(bodyFileExxcel, {cellDates:true})

            // ws['!rows'] = [{hpt:50},{hpt:50}];

            // const bufferExcel = await xlsx.build([{ name: "vinaphone_sheet", data: bodyFileExxcel }],optionsExcel)

            // await fs.writeFile("ketqua.xlsx", bufferExcel, function cb(err) {
            //     if (err) throw err;
            //     //console.log('Đã ghi vào file excel...');
            // });



            // try {
            // await xlsx.readFile("ketqua.xlsx");
            // var worksheet = await workbook.Sheets[0];
            // await XLSX.writeFile('ketqua.xlsx', {cellStyles: true});
            // }
            // catch(e){

            // }



            await browser.close();

            await mainWindow.webContents.send('crawl:result', true);
            //await ipcMain.send('crawl:result', true);            
        }

        start();


    }).catch(async (err) => {
        //console.log("cindererr", err);
        await mainWindow.webContents.send('crawl:network_error', true);
    });
}

function timer(ms) {
    return new Promise(res => setTimeout(res, ms));
}

async function asyncForEach(array, callback) {
    let cIndex = 1;
    for (let index = 0; index < array.length; index++) {
        await callback(array[index], index);
        console.log("xong 1 lan",cIII + " " + cTotal);
        mainWindow.webContents.send('crawl:onrunning', (index+1) + " " + array.length);
        if (index == cIndex * limitRequest - 1 && index < array.length - 1) {
            cIndex++;
            await timer(delayInMilliseconds);
        }
    }
}

