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

let replaceAll = function(src: string, s1: string | RegExp, s2: string){ 
    return src.replace(new RegExp(s1,"gm"),s2); 
}

let writeServerFile = function(file : string, json : any) {
    let jsonString = JSON.stringify(json, null, 4);
    let path = serverOutPath + "/" + file;
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

let excelArray = function(xlsx: string | any[], fileName: any, sheetName: string) {
    let jsonFileName = sheetName + ".json";

    let keyArray = [];
    if (xlsx.length <= 1) {
        return [];
    }

    const TypeString = "string";
    const TypeFloat = "float";
    const TypeInt = "int";
    const TypeJson = "json";
    const FlagBreak = "flag_break"

    let dateLine = 4;
    let typeLine = xlsx[3];
    let keyLine = xlsx[2];
    let nameDesc = {};

    for (let k = 0; k < keyLine.length; ++k) {
        let key = keyLine[k];
        console.log("key: " + key);
        if(key == undefined) {
            keyArray.push(FlagBreak);
            continue;
        } else {
            // 屏蔽KEY中空格
            key = replaceAll(key, ' ', '');
        }
        
        keyArray.push(key);
        let typeString : string = typeLine[k];
        let type = "";
        if (typeString.toUpperCase() == "STRING") {
            type = TypeString;
        } else if (typeString.toUpperCase() == "FLOAT" || typeString.toUpperCase() == "NUMBER") {
            type = TypeFloat;
        } else if (typeString.toUpperCase() == "INT") {
            type = TypeInt;
        } else if (typeString.toUpperCase() == "JSON") {
            type = TypeJson;
        } else {
            throw("Invalid type : " + typeString);
        }
        
        nameDesc[key] = { type : type};
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
            if(key == FlagBreak) {
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

let parseXlsx = function(excelPath: string, fileName: string) {
    let fileFullPath = excelPath + "/" + fileName;
    const workSheetsFromFile = xlsx.parse(fileFullPath);
    for (let k in workSheetsFromFile) {
        let sheetName = workSheetsFromFile[k].name;
        excelArray(workSheetsFromFile[k].data, fileName, sheetName);
    }
}
