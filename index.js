const express = require('express')
const fs = require('fs')
const PizZip = require("pizzip");
const DocxTemplater = require("docxtemplater")
const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);

const app = express()
app.use(express.json())
app.use(express.urlencoded({ extended: true }))

/**
 * 获取模版文件二进制内容
 * @param {*} templateFile 模版文件
 * @returns Promise<FileReader>
 */
function getFileBinaryString(templateFile) {
  return new Promise((resolve, reject) => {
    const content = fs.readFileSync(templateFile,"binary")
    resolve(content)
  });
}

/**
 * 根据 docx 模版文件重新生成新的 docx 文件
 * @param {*} templatePath 模版文件路径
 * @param {*} varData 传入模版处理的变量
 * @returns 
 */
async function generateDocxFile(templatePath, varData, outputPath) {
  return new Promise((resolve, reject) => {
    getFileBinaryString(templatePath)
      .then(templateContent => {
        const zip = new PizZip(templateContent);
        const doc = new DocxTemplater()
          .loadZip(zip)
          .render(varData); // varData是我们需要定义好，传给docxtempale的数据。
        
        // Get the zip document and generate it as a nodebuffer
        const outBuffer = doc.getZip().generate({
          type: "nodebuffer",
          // compression: DEFLATE adds a compression step.
          // For a 50MB output document, expect 500ms additional CPU time
          compression: "DEFLATE",
        });

        // buf is a nodejs Buffer, you can either write it to a
        // file or res.send it with express for example.
        fs.writeFileSync(outputPath, outBuffer);
        resolve(outBuffer);
      })
      .catch(reject);
  });
}

/**
 * 将 docx 文档转换为 pdf 文档
 * @param {*} inputDocxPath 
 * @param {*} outputPdfPath 
 */
async function docxToPdf(inputDocxPath, outputPdfPath) {
  const ext = '.pdf'
  const docxBuf = fs.readFileSync(inputDocxPath);

  // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
  let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);
  
  // Here in done you have pdf file which you can save or transfer in another stream
  return fs.writeFileSync(outputPdfPath, pdfBuf);
}

// 1. 读取模板并生成新的 docx 文档
app.get('/generate/docx', async (req, res) => {
    const varData = req.query
    if (!fs.existsSync('data/cache')) {
      fs.mkdirSync('data/cache')
    }

    const outputPath = `data/cache/${varData.template}_${Date.now()}.docx`
    const templatePath = `data/templates/${varData.template}.docx`
    try {

      await generateDocxFile(templatePath, varData, outputPath)

        // 下载文件
        res.download(outputPath, (err) => {
          if (err) {
            console.log('Error downloading file:', err);
          } else {
            console.log('File downloaded successfully.');
          }
        })
    } finally {
      // 30s after remove cache generate file
      setTimeout(() => {
        fs.unlinkSync(outputPath);
      }, 60000)
    }
})
    

const port = process.env.PORT || 11000
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`)
})
