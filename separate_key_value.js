const reader = require('xlsx');

const data = {
  one: 1,
  two: {
    three: 3,
  },
  four: {
    five: 5,
    six: {
      seven: 7,
    },
    eight: 8,
  },
  nine: {
    h1: "hello",
    h2: "world",
  },
  a: {
    k1: "pragya",
    k2: {
      k3: "singh",
      k4: "from",
      k5: "jamshedpur",
    },
  },
};
//---------------flat json--------------------
const flattenJSON = (data = {}, res = {}, extraKey = "") => {
  for (key in data) {
    if (typeof data[key] !== "object") {
      res[extraKey + key] = data[key];
    } else {
      flattenJSON(data[key], res, `${extraKey}${key}.`);
    }
  }
  return res;
};
const flat_json = flattenJSON(data);
console.log(flat_json);

//-----------------separating json keys-----------------------
const getNestedKeys = (flat_json, keys) => {
  if (!(flat_json instanceof Array) && typeof flat_json == "object") {
    Object.keys(flat_json).forEach((key) => {
      keys.push(key);
      const value = flat_json[key];
      if (typeof value === "object" && !(value instanceof Array)) {
        getNestedKeys(value, keys);
      }
    });
  }
  return keys;
};
const json_key = (getNestedKeys(flat_json, []));
console.log(json_key);

//---------------------separating json values-----------------------
const getNestedValues = (data, values) => {
  if (!(data instanceof Array) && typeof data == "object") {
    Object.values(data).forEach((value) => {
      if (typeof value === "object" && !(value instanceof Array)) {
        getNestedValues(value, values);
      } else {
        values.push(value);
      }
    });
  }
  return values;
};
const json_value = getNestedValues(data, [])
console.log(json_value);
//------------------json keys to excel file----------------

let convertJsonKeyToExcel = [json_key]

const jsonKeyToExcel=()=>{
  
    const ws = reader.utils.json_to_sheet(convertJsonKeyToExcel)
    const wb = reader.utils.book_new()
    reader.utils.book_append_sheet(wb, ws, 'json_keys')
    reader.writeFile(wb, 'inputJsonKey.xlsx')
}
jsonKeyToExcel();
console.log("JSON keys shifted to excel file!!");

//---------------------json value to excel---------------------------  

let convertJsonValueToExcel  = [json_value]

const jsonValueToExcel=()=>{

  const ws = reader.utils.json_to_sheet(convertJsonValueToExcel)
  const wb = reader.utils.book_new()
  reader.utils.book_append_sheet(wb, ws, 'json_values')
  reader.writeFile(wb, 'inputJsonValue.xlsx')
}
jsonValueToExcel();
console.log("json values shifted to excel file!!");

//---------------------merging two excel files-------------------

let mergingTwoExcelFiles  = [json_key, json_value]

const translatedKeyValueFile=()=>{

  const ws = reader.utils.json_to_sheet(mergingTwoExcelFiles)
  const wb = reader.utils.book_new()
  reader.utils.book_append_sheet(wb, ws, 'json_values')
  reader.writeFile(wb, 'translatedKeyValue.xlsx')
}
translatedKeyValueFile();
console.log("translated key value excel file!!");