"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var path_1 = __importDefault(require("path"));
var xlsx = __importStar(require("xlsx"));
var INPUT_FILE = 'file.xlsx';
var OUTPUT_FILE = 'result.xlsx';
var notAllowedStartsWith = [
    '-----',
    '____',
    '    ',
];
var loadFileToJson = function () {
    var workbook = xlsx.readFile(path_1.default.join(__dirname, "../input/".concat(INPUT_FILE)));
    var sheet = workbook.Sheets[workbook.SheetNames[0]];
    var json = xlsx.utils.sheet_to_json(sheet);
    json.splice(0, 5);
    return json;
};
var simplifyData = function (json) {
    var data = json.map(function (row) {
        return row['__EMPTY'];
    });
    return data;
};
var groupFirstAndSecondLines = function (arr) {
    var result = [];
    for (var i = 0; i < arr.length; i += 2) {
        result.push([arr[i], arr[i + 1]]);
    }
    return result;
};
var getNameAndValueFromArray = function (arr) {
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        var name = arr[i][0];
        var value = arr[i][1];
        result.push({ name: name, value: value });
    }
    return result;
};
var extractOnlyNameFromText = function (text) {
    var name = '';
    var regexNumber = /\d/;
    for (var i = 0; i < text.length; i++) {
        if (regexNumber.test(text[i])) {
            name = text.slice(0, i);
            break;
        }
    }
    return name.trim();
};
var clearNameAndValue = function (arr) {
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        var name = extractOnlyNameFromText(arr[i].name);
        var valueArr = arr[i].value.split(' ');
        var value = valueArr[valueArr.length - 1];
        result.push({ name: name, value: value });
    }
    return result;
};
var removeRows = function (arr) {
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        var isAllowed = true;
        for (var j = 0; j < notAllowedStartsWith.length; j++) {
            if (arr[i].startsWith(notAllowedStartsWith[j])) {
                isAllowed = false;
                break;
            }
        }
        if (isAllowed)
            result.push(arr[i]);
    }
    return result;
};
var saveResult = function (arr) {
    var wb = xlsx.utils.book_new();
    var ws = xlsx.utils.json_to_sheet(arr);
    xlsx.utils.book_append_sheet(wb, ws, 'result');
    xlsx.writeFile(wb, path_1.default.join(__dirname, "../output/".concat(OUTPUT_FILE)));
};
var json = loadFileToJson();
var simplifiedData = simplifyData(json);
var removeTrash = removeRows(simplifiedData);
var gropedData = groupFirstAndSecondLines(removeTrash); // First line is name, second is value
var nameAndValue = getNameAndValueFromArray(gropedData);
var clearedNameAndValue = clearNameAndValue(nameAndValue);
saveResult(clearedNameAndValue);
console.log(clearedNameAndValue[3]);
