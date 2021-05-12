const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');


// 通常产品就是仍一个excel 啥都往里面写（通常不会按模块分成多个文件）

main()

async function main() {
    const excelPath = path.resolve(__dirname, './src/elinking.xlsx');
    let sheetObj = false;
    try {
        sheetObj = await readExcel(excelPath)
    } catch (e) {
        console.log(e)
    }
    if (sheetObj === false) {
        return
    }

    // 1、 输出 jsonData.json 文件
    const faqOutPutPath = path.resolve(__dirname, './dist/jsonData.json');
    const faqData = JSON.parse(JSON.stringify(sheetObj)); // 避免日后某些方法修改了传参
    writeJsonFile(faqOutPutPath, excelToJson_faq(faqData))


    // 2、 输出 dailyPaymentContent_1.json 文件

    const contentGroup = contentSplice(sheetObj);

    const groupKeys = Object.keys(contentGroup);

    groupKeys.forEach(item => {
        const contentOneOutPutPath = path.resolve(__dirname, `./dist/dailyPaymentContent_${item}.json`);
        writeJsonFile(contentOneOutPutPath, excelToJson_content(contentGroup[item]))
    })


}


function writeJsonFile(outPutPath, data) {
    const excelPath = path.resolve(__dirname, './src/elinking.xlsx');
    readExcel(excelPath).then(result => {
        fs.writeFile(outPutPath, data, (err) => {
            if (err) throw err;
            console.log('文件已被保存');
        });
    })
}

function readExcel(filePath) {
    return new Promise((resolve, reject) => {
        fs.readFile(filePath, (err, data) => {
            if (err) {
                reject(err);
                return
            }
            try {
                const excelWork = xlsx.parse(data)
                resolve(excelWork)
            } catch (e) {
                reject(e)
            }
        });
    })
}

function excelToJson_faq(excelObject) {
    // 将 xlsx  解析的表格数据按需求转化为相应的JSON
    const sheetArr = excelObject[0].data;
    const faqArr = sheetArr.slice(1, sheetArr.length)
    const faqDetail = faqArr.reduce((accumulator, currentValue, index) => {
        const faqObj = {
            issueContent_ch: currentValue[1],
            issueContent_en: currentValue[2],
            issueContent_cht: currentValue[3],
            issueContent_my: currentValue[4],
            issueContent_id: currentValue[5],

            answer_ch: currentValue[6],
            answer_en: currentValue[7],
            answer_cht: currentValue[8],
            answer_my: currentValue[9],
            answer_id: currentValue[10],
        };
        return Object.assign(accumulator, {[index + 1]: faqObj})
    }, {});
    return JSON.stringify(faqDetail)
}

function contentSplice(sheetData) {
    const sheetArr = sheetData[0].data;
    const contentArr = sheetArr.slice(1, sheetArr.length);
    const contentTemp = {};
    let keyCount = 0;
    contentArr.forEach(item => {
        if (item[0]) {
            keyCount++;
            contentTemp[keyCount] = []
        }
        contentTemp[keyCount].push(item)
    });
    return contentTemp
}

function excelToJson_content(excelObject) {
    // 将 xlsx  解析的表格数据按需求转化为相应的JSON
    const contentData = excelObject.reduce((accumulator, currentValue, index) => {
        return accumulator.concat([{
            faqId: index + 1,
            issueContent_ch: currentValue[1],
            issueContent_en: currentValue[2],
            issueContent_cht: currentValue[3],
            issueContent_my: currentValue[4],
            issueContent_id: currentValue[5],
        }])

    }, []);

    return JSON.stringify(contentData)

}
