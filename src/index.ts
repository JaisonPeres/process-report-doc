import path from 'path';
import * as xlsx from 'xlsx';

const INPUT_FILE = 'file.xlsx';
const OUTPUT_FILE = 'result.xlsx';

const notAllowedStartsWith = [
  '-----',
  '____',
  '    ',
]

interface IRow {
  [key: string]: string;
}

interface InterestData {
  name: string;
  value: string;
}

const loadFileToJson = () => {
  const workbook = xlsx.readFile(path.join(__dirname, `../input/${INPUT_FILE}`));
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = xlsx.utils.sheet_to_json<IRow>(sheet);
  json.splice(0, 5);
  return json;
}

const simplifyData = (json: any) => {
  const data = json.map((row: IRow) => {
    return row['__EMPTY']
  });
  return data;
}

const groupFirstAndSecondLines = (arr: any[]) => {
  const result = [];
  for (let i = 0; i < arr.length; i += 2) {
    result.push([arr[i], arr[i + 1]]);
  }
  return result;
}

const getNameAndValueFromArray = (arr: string[][]) => {
  const result = [];
  for (let i = 0; i < arr.length; i++) {
    const name = arr[i][0];
    const value = arr[i][1];
    result.push({ name, value });
  }
  return result;
}

const clearNameAndValue = (arr: InterestData[]) => {
  const result = [];
  for (let i = 0; i < arr.length; i++) {
    const name = arr[i].name.split('0')[0].trim();
    const valueArr = arr[i].value.split(' ');
    const value = valueArr[valueArr.length - 1];
    result.push({ name, value });
  }
  return result;
};

const removeRows = (arr: string[]) => {
  const result = [];
  for (let i = 0; i < arr.length; i++) {
    let isAllowed = true;
    for (let j = 0; j < notAllowedStartsWith.length; j++) {
      if (arr[i].startsWith(notAllowedStartsWith[j])) {
        isAllowed = false;
        break;
      }
    }
    if (isAllowed) result.push(arr[i]);
  }
  return result;
}

const saveResult = (arr: InterestData[]) => {
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(arr);
  xlsx.utils.book_append_sheet(wb, ws, 'result');
  xlsx.writeFile(wb, path.join(__dirname, `../output/${OUTPUT_FILE}`));
}


const json = loadFileToJson();
const simplifiedData = simplifyData(json);
const removeTrash = removeRows(simplifiedData);
const gropedData = groupFirstAndSecondLines(removeTrash); // First line is name, second is value
const nameAndValue = getNameAndValueFromArray(gropedData);
const clearedNameAndValue = clearNameAndValue(nameAndValue);
saveResult(clearedNameAndValue);

console.log(clearedNameAndValue[3]);