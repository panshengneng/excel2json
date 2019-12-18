import * as fs from "fs";

import xlsx from "node-xlsx";

let outPath = "./json";
let serverOutPath = outPath + "/server";

let pathArray = [outPath, serverOutPath];

for (let path of pathArray) {
    if (!fs.existsSync(path)) {
        fs.mkdirSync(path);
    }
}

// /**
//  * 字符串格式化
//  * @param format 格式 
//  * @param arg 参数
//  */
// let stringFormat = function(format : string, ...arg : any[]) {  
//     let str = String(format);  
//     for (let i = 0; i < arg.length; i++) {  
//         let re = new RegExp('\\{' + (i) + '\\}', 'gm');  
//         str = str.replace(re, arg[i]);
//     }  
//     return str; 
// }

let replaceAll = function(src, s1, s2){ 
    return src.replace(new RegExp(s1,"gm"),s2); 
}

// let writeFile = function(file : string, json : any) {
//     let jsonString = JSON.stringify(json, null, 4);
//     let path = "./json/" + file;
//     path = replaceAll(path, '.xlsx', '.json');
//     fs.writeFileSync(path, jsonString);
// }

let writeServerFile = function(file : string, json : any) {
    let jsonString = JSON.stringify(json, null, 4);
    let path = serverOutPath + "/" + file;
    // path = replaceAll(path, '.xlsx', '.json');
    fs.writeFileSync(path, jsonString);
    console.log("File Success : " + path);
}

let excelPath : string = "./excel";
fs.readdir(excelPath, (err: NodeJS.ErrnoException, files: string[]) => {
    if (err) {
        console.log("Error");
        console.log("Excel Path :" + excelPath);
        return;
    }
    for (let fileName of files) {
        if (fileName.indexOf("~$") != -1) {
            continue;
        }
        if (fileName.indexOf(".xlsx") == -1 && fileName.indexOf(".xls") == -1) {
            continue;
        }

        parseXlsx(excelPath, fileName);
    }
});

let excelArray = function(xlsx, fileName, sheetName) {
    let jsonFileName = sheetName + ".json";

    let keyArray = [];
    if (xlsx.length <= 1) {
        return [];
    }

    const TypeString = "string";
    const TypeFloat = "float";
    const TypeInt = "int";
    const TypeJson = "json";

    let dateLine = 4;
    let typeLine = xlsx[3];
    let keyLine = xlsx[2];
    let nameDesc = {};

    console.log("typeLine :" + typeLine)
    console.log("keyLine :" + keyLine)

    for (let k in keyLine) {
        let key = keyLine[k];
        // 屏蔽KEY中空格
        key = replaceAll(key, ' ', '');
        keyArray.push(key);
        let typeString : string = typeLine[k];
        let c : boolean = true;
        let s : boolean = true;
        // if (typeString.indexOf("-C") != -1 || typeString.indexOf("-c") != -1) {
        //     c = false;
        // } else if (typeString.indexOf("-S") != -1 || typeString.indexOf("-s") != -1) {
        //     s = false;
        // }

        // let keyTypeString : string = replaceAll(typeString, '-C', '');
        // keyTypeString = replaceAll(keyTypeString, '-c', '');
        // keyTypeString = replaceAll(keyTypeString, '-s', '');
        // keyTypeString = replaceAll(keyTypeString, '-S', '');
        let type = "";
        let keyTypeString = typeString;
        console.log("keyTypeString :" + keyTypeString)
        if (keyTypeString.toUpperCase() == "STRING") {
            type = TypeString;
        } else if (keyTypeString.toUpperCase() == "FLOAT" || keyTypeString.toUpperCase() == "NUMBER") {
            type = TypeFloat;
        } else if (keyTypeString.toUpperCase() == "INT") {
            type = TypeInt;
        } else if (keyTypeString.toUpperCase() == "JSON") {
            type = TypeJson;
        } else {
            throw("Invalid type : " + keyTypeString);
        }
        
        nameDesc[key] = { type : type, c : c, s : s };
    }

    let sdata = [];
    for (let line = dateLine; line < xlsx.length; ++line) {
        let lineData = xlsx[line];

        let da : any = {};
        if (lineData[0] === "" || lineData[0] === undefined) {
            continue;
        }

        for (let k = 0; k < keyArray.length; ++k) {
            let value = lineData[k];
            let key = keyArray[k];
            if (!nameDesc[key].s) {
                continue;
            }
            let type = nameDesc[key].type;
            let typeValue = null;
            try {
                if (type == TypeString) {
                    typeValue = value ? String(value) : "";
                    if (typeValue == null) {
                        typeValue = "";
                    }
                } else if (type == TypeFloat) {
                    typeValue = value ? Number(value) : 0;
                    if (value != null && (typeValue == null || isNaN(typeValue))) {
                        throw("error");
                    }
                    if (typeValue == null) {
                        typeValue = 0;
                    }
                } else if (type == TypeInt) {
                    typeValue = value ? parseInt(value) : 0;
                    if (value != null && (typeValue == null || isNaN(typeValue))) {
                        throw("error");
                    }
                } else if (type == TypeJson) {
                    typeValue = value ? JSON.parse(value) : "";
                    if (value != null && typeValue == null) {
                        throw("error");
                    }
                }
            } catch (e) {
                console.log('Error, Excel :' + sheetName + ', key :\"' + key + '\", Type : ' + type + ', value :\"' + value + '\", line :' + (line + 1));
                return;
            }
            da[key] = typeValue;
        }
        sdata.push(da);
    }
    writeServerFile(jsonFileName, sdata);
}

let parseXlsx = function(excelPath, fileName) {
    let fileFullPath = excelPath + "/" + fileName;
    const workSheetsFromFile = xlsx.parse(fileFullPath);
    // let data = [];
    for (let k in workSheetsFromFile) {
        let sheetName = workSheetsFromFile[k].name;
        let coData = excelArray(workSheetsFromFile[k].data, fileName, sheetName);
        // data.push(coData);
    }
    // return data;
}
