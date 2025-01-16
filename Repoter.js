// Load our library that generates the document
const Docxtemplater = require("docxtemplater");
// Load PizZip library to load the docx/pptx/xlsx file in memory
const PizZip = require("pizzip");
const ImageModule = require("docxtemplater-image-module-free");

// Builtin file system utilities
const fs = require("fs");
const path = require("path");

class Repoter {
    // 构造函数
    constructor({ templet = "", tittleLevel = 0 } = {}) {
        this.keys = ["paragraph", "pic", "pic_desc", "table_name", "table"]
        this.tittles = []

        let content = ""

        if (templet == "") {
            // Load the docx file as binary content
            content = fs.readFileSync(
                path.resolve(__dirname, "InputFile.docx"),
                "binary"
            );

            if (tittleLevel == 0) {
                this.keys = ["tittle1", "tittle2", "tittle3", "tittle4", "paragraph", "pic", "pic_desc", "table_name", "table"]
                this.tittles = [1, 2, 3, 4]
            }
        } else {
            content = fs.readFileSync(
                templet,
                "binary"
            );

        }

        // Unzip the content of the file
        this.zip = new PizZip(content);

        if (tittleLevel != 0) {
            for (let i = 0; i < tittleLevel; i++) {
                let level = i + 1
                let tittle = `tittle${level}`

                this.tittles.push(level)
                this.keys.push(tittle)
            }
        }

        // console.log("keys, ", this.keys)


        // 组内默认属性
        this.tittleNow = 1          // 此时无标题
        this.hasParagraph = false   // 此时无段落
        this.hasPic = false         // 此时无图片
        this.hasTable = false       // 此时无表格

        this.mainpages = []

        this.imgDataDict = []
        this.picCounter = 0
        this.renderFlag = false
        this.doc = new Docxtemplater(this.zip, {
            paragraphLoop: true,
            linebreaks: true,
            modules: [
                new ImageModule({
                    getImage: (value, key) => {
                        this.picCounter += 1
                        // console.log('this.picCounter, ',this.picCounter, this.imgDataDict[this.picCounter]["pic"])
                        return this.imgDataDict[this.picCounter]["pic"];
                    },
                    getSize: (afterValue, value, key) => {
                        // console.log('this.picCounter, ',this.picCounter)
                        return this.imgDataDict[this.picCounter]["size"];
                        // return [400, 400];
                    },
                })
            ],
        });
    }

    // Render data to files
    render() {
        if (this.renderFlag) {
            console.warn("You Could only render Once! All work about the content below will be abandon!")
        }

        // 长度+1 此时最后一张图片才能正常显示
        if (this.imgDataDict.length != 0) {
            this.imgDataDict.push(this.imgDataDict[0])
        }

        this.doc.render({ mainpages: this.mainpages })
        this.renderFlag = true
        console.log("Render file success!")
    }

    // Save file
    saveFile() {
        if (!this.renderFlag) {
            this.render()
        }

        try {

            // 获取文档的 zip 包内容
            let zipData = this.doc.getZip();

            // 获取并修改 word/document.xml 内容
            let documentXml = zipData.file("word/document.xml").asText();

            // 删除空段落（检测 <w:p> 标签内没有文本的段落）
            // 这里的正则保证了只删除空的段落，不会破坏其他内容
            let cleanedDocumentXml = documentXml.replace(
                /((<w:p\s[^><]*?>)|(<w:p>))(<[^><]*?>)+?<\/w:p>((<w:p\s[^><]*?>)|(<w:p>))(<[^><]*?>)+?<\/w:p>/g, '');
            // /((<w:p\s[^><]*?>)|(<w:p>))(<[^><]*?>)+?<\/w:p>/g // 一个回车

            // 更新 XML 文件内容
            zipData.file("word/document.xml", cleanedDocumentXml);

            // 重新生成并保存为 docx 文件
            let outputPath = path.resolve(__dirname, 'OutputFile.docx');
            let newDocxData = zipData.generate({ type: 'nodebuffer' });
            fs.writeFileSync(outputPath, newDocxData);

            console.log("Document generated and saved to", outputPath);
        } catch (error) {
            console.error("Error rendering document:", error);
        }
    }

    // 添加一级标题
    // 参数：
    //      level： 标题等级 —— int
    //      tittle：标题文本 —— str
    addTittle(level, tittle) {

        if (this.renderFlag) {
            console.warn("You have rendered the file, action addTittle was abandoned!")
            return false
        }

        let tittleKey = ""
        if (this.tittles.includes(level)) {
            tittleKey = `tittle${level}`
        } else {
            console.warn(`Unexcepted tittle level ${level}! Please check the arg level!`)
            return false
        }

        if (this.mainpages.length == 0 | (this.mainpages.length > 0 & this.tittleNow > level) | this.hasParagraph | this.hasPic | this.hasTable) {
            // 没有现成组存在
            // 或 当前组内标题等级高于新插入的级别
            // 或 当前组内已有段落
            // 或 当前组内已有图片
            // 或 当前组内已有表格
            let entity = {}
            entity[tittleKey] = tittle
            this.mainpages.push(entity)

            // 组属性归零
            this.tittleNow = 1
            this.hasParagraph = false
            this.hasPic = false
            this.hasTable = false
        } else {
            // 继续上个对象增加
            this.mainpages[this.mainpages.length - 1][tittleKey] = tittle
            this.tittleNow = level
        }

        return true
    }

    // 添加系列标题
    // 参数：
    //      startLevel  标题开始等级 —— int
    //      tittles     标题文本 —— list od strs
    addTittles(startLevel, tittles) {
        if (this.renderFlag) {
            console.warn("You have rendered the file, action addTittle was abandoned!")
            return false
        }

        if (startLevel > this.tittles.length) {
            // 开始的标题等级就过高
            console.warn(`StartLevel ${startLevel} is more than the supported e! Return without any action! Please check the arg startLevel!`)
            return false
        } else if (startLevel + tittles.length > this.tittles.length) {
            // 开始的等级能满足，但是给的标题太多了
            console.warn(`Target tittle level ${startLevel + tittles.length} is more than the supported! Return without any action! Please check the arg startLevel and tittles!`)
            return false
        } else {
            for (let i = 0; i < tittles.length; i++) {
                let level = startLevel + i
                let tittle = tittles[i]

                this.addTittle(level, tittle)
            }

            return true
        }
    }

    // 添加一段文本
    // 参数：
    //      paragraph  段落文本 —— str
    addParagraph(paragraph) {
        if (this.renderFlag) {
            console.warn("You have rendered the file, action addTittle was abandoned!")
            return false
        }

        if (this.mainpages.length == 0 | this.hasPic | this.hasTable) {
            // 没有现成组存在
            // 或 当前组内已有图片
            // 或 当前组内已有表格
            let entity = {}
            entity["paragraph"] = [paragraph]
            this.mainpages.push(entity)

        } else {
            // 继续上个对象增加
            // console.log("this.mainpages: ", this.mainpages)

            if (!("paragraph" in this.mainpages[this.mainpages.length - 1])) {
                this.mainpages[this.mainpages.length - 1]["paragraph"] = []
            }
            this.mainpages[this.mainpages.length - 1]["paragraph"].push(paragraph)
        }

        this.hasParagraph = true

        return true
    }

    // 添加多段文本
    // 参数：
    //      paragraphs  段落文本 —— list of strs
    addParagraphs(paragraphs) {
        if (this.renderFlag) {
            console.warn("You have rendered the file, action addTittle was abandoned!")
            return false
        }

        if (this.mainpages.length == 0 | this.hasParagraph | this.hasPic | this.hasTable) {
            // 没有现成组存在
            // 或 当前组内已有段落
            // 或 当前组内已有图片
            // 或 当前组内已有表格
            let entity = {}
            entity["paragraph"] = paragraphs
            this.mainpages.push(entity)

        } else {
            // 继续上个对象增加
            for (let i = 0; i < paragraphs.length; i++) {
                let p = paragraphs[i]

                this.addParagraph(p)
            }
        }

        this.hasParagraph = true

        return true
    }

    // 添加图片
    // 参数：
    //      picBase64  图片的base64编码 —— str
    //      desc       图片的图注 —— str
    //      size       图片尺寸 —— list, default [400, 400]
    // return:
    //      插入图片的总序号
    addPicWithDesc(picBase64, desc = "示意图", size = [400, 400]) {
        this.imgDataDict.push({
            "pic": this.base64DataURLToArrayBuffer(picBase64),
            "size": size
        })

        // console.log("this.imgDataDict, ", this.imgDataDict)

        if (this.renderFlag) {
            console.warn("You have rendered the file, action addTittle was abandoned!")
            return 0
        }

        if (this.mainpages.length == 0 | this.hasPic | this.hasTable) {
            // 没有现成组存在
            // 或 当前组内已有图片
            // 或 当前组内已有表格
            let entity = {}
            entity["pic"] = 1
            entity["pic_desc"] = desc
            this.mainpages.push(entity)

        } else {
            // 继续上个对象增加
            this.mainpages[this.mainpages.length - 1]["pic"] = 1
            this.mainpages[this.mainpages.length - 1]["pic_desc"] = desc
        }
        this.hasPic = true

        return this.imgDataDict.length
    }

    base64DataURLToArrayBuffer = (dataURL) => {
        const base64Regex = /^data:image\/(png|jpg|jpeg|svg|svg\+xml);base64,/;
        if (!base64Regex.test(dataURL)) {
            return false;
        }
        const stringBase64 = dataURL.replace(base64Regex, "");
        let binaryString;
        if (typeof window !== "undefined") {
            binaryString = window.atob(stringBase64);
        } else {
            binaryString = new Buffer(stringBase64, "base64").toString("binary");
        }
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            const ascii = binaryString.charCodeAt(i);
            bytes[i] = ascii;
        }
        return bytes.buffer;
    }

    // 添加表格
    // 参数：
    //      headers     表头 —— list of strs
    //      rows        表格内容 —— list of list
    //      desc        表题 —— str
    //      corrective  纠正模式 —— bool, default True
    addTable(headers = [], rows, desc = "", corrective = true) {

        if (this.renderFlag) {
            console.warn("You have rendered the file, action addTittle was abandoned!")
            return false
        }

        if (corrective) {
            let maxColNum = 0;
            for (let i = 0; i < rows.length; i++) {
                let row = rows[i]

                let colNum = row.length

                if (colNum > maxColNum) {
                    maxColNum = colNum
                }
            }

            // 纠正标题行
            if (headers.length != 0 & headers.length < maxColNum) {
                for (let i = 0; i < maxColNum - headers.length; i++) {
                    headers.push("")
                }
            }

            // 纠正数据组
            for (let i = 0; i < rows.length; i++) {
                let row = rows[i]

                let colNum = row.length

                if (colNum < maxColNum) {
                    for (let i = 0; i < maxColNum - colNum; i++) {
                        row.push("")
                    }
                }
            }
        }

        let tableXML = this.generateTableXML(headers, rows);

        if (this.mainpages.length == 0 | this.hasTable) {
            // 没有现成组存在
            // 或 当前组内已有表格
            let entity = {}
            entity["table"] = tableXML
            this.mainpages.push(entity)

        } else {
            // 继续上个对象增加
            this.mainpages[this.mainpages.length - 1]["table"] = tableXML
        }

        if (desc != "") {
            this.mainpages[this.mainpages.length - 1]["table_name"] = desc
        }

        this.hasTable = true

        return true

    }

    generateTableXML(headers, rows) {
        let tableXML = `
            <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:tblPr>
                    <w:tblBorders>
                        <w:top w:val="single" w:sz="4" w:color="000000" />
                        <w:left w:val="single" w:sz="4" w:color="000000" />
                        <w:bottom w:val="single" w:sz="4" w:color="000000" />
                        <w:right w:val="single" w:sz="4" w:color="000000" />
                        <w:insideH w:val="single" w:sz="4" w:color="000000" />
                        <w:insideV w:val="single" w:sz="4" w:color="000000" />
                    </w:tblBorders>
                    <w:jc w:val="center"/>
                    <w:ind w:left="0"/> 
    
                    <w:tblW w:w="4500" w:type="pct" />
                </w:tblPr>
        `;

        // 动态生成表头（如果有表头数据）
        if (headers && headers.length > 0) {
            tableXML += `<w:tr>`;
            headers.forEach(header => {
                tableXML += `
                    <w:tc>
                        <w:tcPr>
                            <!-- 不设置固定宽度，让其根据内容调整 -->
                            <w:tcW w:type="auto" />
                        </w:tcPr>
                        <w:p>
                            <w:pPr>
                                <w:jc w:val="center"/>
                                <w:ind w:left="0"/> 
                            </w:pPr>
                            <w:r>
                                <w:rPr>
                                    <w:b/>
                                </w:rPr>
                                <w:t>${header}</w:t>
                            </w:r>
                        </w:p>
                    </w:tc>
                `;
            });
            tableXML += `</w:tr>`;
        }

        // 动态生成数据行
        rows.forEach(row => {
            tableXML += `<w:tr>`;
            row.forEach(cell => {
                tableXML += `
                    <w:tc>
                        <w:tcPr>
                            <!-- 不设置固定宽度，让其根据内容调整 -->
                            <w:tcW w:type="auto" />
                        </w:tcPr>
                        <w:p>
                            <w:pPr>
                                <w:jc w:val="center"/>
                                <w:ind w:left="0"/> 
                            </w:pPr>
                            <w:r>
                                <w:t>${cell}</w:t>
                            </w:r>
                        </w:p>
                    </w:tc>
                `;
            });
            tableXML += `</w:tr>`;
        });

        tableXML += `</w:tbl>`;

        return tableXML;
    }
}


module.exports = Repoter;