const fs = require("fs");
const JSZip = require("jzip");
const path = require("path");
const Docxtemplater = require("docxtemplater");
const compressing = require("compressing");

function buildWord(params, fileName, i) {
  let type = "";
  if (i == 1) {
    if (params.hjd == "" && params.wzfbf == 0) {
      type += "1";
    } else if (params.hjd == "" && params.wzfbf !== 0) {
      type += "2";
    } else if (params.hjd !== "" && params.wzfbf == 0) {
      type += "3";
    } else if (params.hjd !== "" && params.wzfbf !== 0) {
      type += "4";
    }
  }
  if (i == 2) {
    params.dlr2Name = "   ";
    if (params.dlr2[1].name) {
      params.dlr2Name = params.dlr2[1].name;
    }
  }
  if (i == 3) {
    if (params.hjd == "") {
      type += "1";
    } else {
      type += "2";
    }
  }

  var content = fs.readFileSync(
    path.join(__dirname, `../template/起诉状/${fileName}${type}.docx`),
    "binary"
  );
  var zip = new JSZip(content);
  var doc = new Docxtemplater();
  doc.loadZip(zip);
  doc.setData(params);
  try {
    doc.render();
  } catch (error) {
    var err = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties
    };
    console.log(JSON.stringify(err));
    throw error;
  }
  var buf = doc.getZip().generate({ type: "nodebuffer" });
  fs.writeFileSync(
    path.join(__dirname, `../output/起诉状/${fileName}.docx`),
    buf
  );
}

async function getQszDocs(params) {
  let files = [
    "1、法定代表人证明书",
    "2、民事起诉状",
    "3、授权委托书",
    "4、保全申请书",
    "担保函",
    "证据目录"
  ];
  params.finalDlr = params.dlr2[0].value;
  params.finalDlr2 = params.dlr2[1].value;
  params.yg2 = params.yg2.trim() + "                         ";
  files.forEach((file, i) => {
    buildWord(params, file, i);
  });
  await zipFolder(
    path.join(__dirname, `../output/起诉状`),
    path.join(__dirname, `../output/result/result.zip`)
  );
}

function zipFolder(inputFolder, outPutZip) {
  return new Promise((resolve, reject) => {
    compressing.zip
      .compressDir(inputFolder, outPutZip)
      .then(res => {
        resolve();
      })
      .catch(err => {
        reject(err);
      });
  });
}

module.exports = {
  getQszDocs
};
