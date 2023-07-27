const asyncHandler = require("express-async-handler");
const mammoth = require("mammoth");
const fs = require("fs");
const upload = require("express-fileupload");
const multiparty = require("multiparty");
const util = require("util");
const reader = require('any-text');
const { UploadFileRequest, WordsApi, SaveAsRequest, HtmlSaveOptionsData, DownloadFileRequest, ConvertDocumentRequest } = require("asposewordscloud");
const XLSX = require("xlsx");



exports.index_get = asyncHandler(async (req, res, next) => {
    // var form = new formidable.IncomingForm();
    // form.parse(req, function (err, fields, files) {
    //     res.write("File uploaded!");
    //     res.end();
    // })

    // res.sendFile(__dirname + "/index.html");

    res.render("index", { title: "FAR/IP Data Extraction Tool" });
});

exports.index_post = asyncHandler(async (req, res, next) => {
    if (req.url === "/import" && req.method === "POST") {

        var form = new multiparty.Form();
        form.parse(req, function (err, fields, files) {

            const clientID = "01814e31-a76b-43c9-8908-5558ba9e39f4";
            const secret = "730c9a7f0f088a80d38fc12885618d25";
            const wordsAPI = new WordsApi(clientID, secret);


            var filePath = files.document[0].path;
            console.log(files.document);
            // var fileName = filePath.substring(filePath.lastIndexOf("/") + 1);

            var request = new ConvertDocumentRequest();
            request.document = fs.createReadStream(filePath);
            request.format = "html";

            wordsAPI.convertDocument(request)
                .then((res) => {
                    result = res.body.toString();
                    // console.log(result);
                    text1 = result.substring(result.indexOf(/Program/g), result.indexOf(/20\d\d<\/span>.+Previous/g) + 4);
                    // console.log(text1);


                    // var text2 = result.substring(result.indexOf("Current Review: Findings and Recommendations"), result.lastIndexOf("</span></li>") + + 12);
                    var text2 = result.substring(result.indexOf("Current Review: Findings and Recommendations"));
                    console.log(text2);
                    var temp2 = text2;

                    if (filePath.match(/\.pdf\b/g)) {
                        // var arr = new Array();
                        // while (temp2.indexOf("<") >= 0) {
                        //     if (temp2.indexOf("<ol") == 0) {
                        //         var arr_line = new Array();
                        //         var text_line = temp2.substring(temp2.indexOf("<span"));
                        //         text_line = text_line.substring(text_line.indexOf(">") + 1, text_line.indexOf("</span>"));
                        //         arr_line.push(text_line);
                        //         arr.push(arr_line);
                        //         // temp2 = temp2.substring(temp2.indexOf("</ol>"));
                        //     }
                        //     else if (temp2.indexOf("<p") == 0) {
                        //         var text_temp = temp2.substring(temp2.indexOf("<span"), temp2.indexOf("</p>"));
                        //         var text_line = "";
                        //         while (text_temp.indexOf("<span") >= 0) {
                        //             if (text_temp.indexOf("<span") == 0) {
                        //                 var text_part = text_temp.substring(text_temp.indexOf("<span"));
                        //                 text_part = text_part.substring(text_part.indexOf(">") + 1, text_part.indexOf("</span>"));
                        //                 if (text_part.match(/^[\sa-zA-Z]/gm)) {
                        //                     console.log("text_line = " + text_line);
                        //                     console.log("text_part = " + text_part);
                        //                     text_line += text_part;
                        //                 }
                        //             }
                        //             text_temp = text_temp.substring(1);
                        //             text_temp = text_temp.substring(text_temp.indexOf("<"));
                        //         }

                        //         text_line = text_line.replaceAll("&#xa0;", "");
                        //         text_line = text_line.replaceAll("<span>•", "");
                        //         text_line = text_line.trim();

                        //         if (text_line.includes("strength")) {
                        //             text_line = "Strengths";
                        //         } else if (text_line.includes("areas of concern")) {
                        //             text_line = "Areas of Concern";
                        //         } else if (text_line.includes("recommendation")) {
                        //             text_line = "Recommendations";
                        //         }

                        //         arr[arr.length - 1].push(text_line);
                        //     }
                        //     else if (temp2.indexOf("<ul") == 0) {
                        //         while (temp2.indexOf("<li") >= 0) {
                        //             if (temp2.indexOf("<li") == 0) {
                        //                 var text_temp = temp2.substring(temp2.indexOf("<span"), temp2.indexOf("</p>"));
                        //                 var text_line = "";
                        //                 while (text_temp.indexOf("<span") >= 0) {
                        //                     if (text_temp.indexOf("<span") == 0) {
                        //                         var text_part = text_temp.substring(text_temp.indexOf("<span"));
                        //                         text_part = text_part.substring(text_part.indexOf(">") + 1, text_part.indexOf("</span>"));
                        //                         if (text_part.match(/^[\sa-zA-Z]/gm)) {
                        //                             console.log("text_line = " + text_line);
                        //                             console.log("text_part = " + text_part);
                        //                             text_line += text_part;
                        //                         }
                        //                     }
                        //                     text_temp = text_temp.substring(1);
                        //                     text_temp = text_temp.substring(text_temp.indexOf("<"));
                        //                 }

                        //                 text_line = text_line.replaceAll("&#xa0;", "");
                        //                 text_line = text_line.replaceAll("<span>•", "");
                        //                 text_line = text_line.trim();
                        //             }
                        //         }
                        //         if (arr[arr.length - 1].length != 0) {
                        //             arr[arr.length - 1].push(text_line);
                        //             arr.push(new Array());
                        //         } else {
                        //             for(var i = 0; i < arr[arr.length - 2].length - 1; i++) {
                        //                 arr[arr.length - 1].push(arr[arr.length - 2][i]);
                        //                 arr[arr.length - 1].push(text_line);
                        //                 arr.push(new Array());
                        //             }
                        //         }
                        //         temp2 = temp2.substring(temp2.indexOf("</ul>") + 1);
                        //         temp2 = temp2.substring(temp2.indexOf("<"));
                        //     }
                        //     temp2 = temp2.substring(1);
                        //     temp2 = temp2.substring(temp2.indexOf("<"));
                        // }
                        // console.log(arr);

                    }
                    else {
                        var max_n_col = 5;
                        var n_rows = (text2.match(/<\/span><\/li>/gm) || []).length - 2;
                        // var n_rows = 56;
                        console.log("n_rows = " + n_rows);
                        var check_next_p = false;
                        var check_next_li = false;
                        var arr = new Array(n_rows);
                        for (var i = 0; i < arr.length; i++) {
                            arr[i] = new Array(5);
                        }
                        var loop = 0;

                        var i = -1;
                        var j = 0;

                        while (temp2.indexOf("<") >= 0) {

                            if (temp2.indexOf("<ol") == 0) {
                                i += 1;
                                j = 0;
                                var str = temp2.substring(temp2.indexOf("<span"));
                                str = str.substring(str.indexOf(">") + 1, str.indexOf("</span>"));
                                arr[i][j] = str;

                                temp2 = temp2.substring(temp2.indexOf("</ol>") + 1);
                                check_next_p = false;
                            } else if (temp2.indexOf("<p") == 0) {
                                var line = temp2.substring(0, temp2.indexOf("</p>"));
                                if (line.includes("strengths")) {
                                    if (check_next_p) {
                                        i += 1;
                                        arr[i][j - 1] = arr[i - 1][j - 1];
                                    }
                                    check_next_p = true;
                                    j = 1;
                                    arr[i][j] = "Strengths";
                                }
                                else if (line.includes("areas of concern")) {
                                    if (check_next_p) {
                                        i += 1;
                                        arr[i][j - 1] = arr[i - 1][j - 1];
                                    }
                                    j = 1;
                                    arr[i][j] = "Areas of concern";
                                }
                                else if (line.includes("recommendations")) {
                                    if (check_next_p) {
                                        i += 1;
                                        arr[i][j - 1] = arr[i - 1][j - 1];
                                    }
                                    j = 1;
                                    arr[i][j] = "Recommendations";
                                }
                                temp2 = temp2.substring(temp2.indexOf("</p>") + 1);
                            } else if (temp2.indexOf("<ul") == 0) {
                                check_next_li = false;
                                temp2 = temp2.substring(1);
                            } else if (temp2.indexOf("<li") == 0) {
                                if (check_next_li) {
                                    i += 1;
                                    for (var temp_j = 0; temp_j < j; temp_j++) {
                                        arr[i][temp_j] = arr[i - 1][temp_j];
                                    }
                                } else {
                                    j += 1;
                                }

                                var str = temp2.substring(temp2.indexOf("<span"), temp2.indexOf("</span>"));
                                str = str.substring(str.indexOf(">") + 1);
                                arr[i][j] = str;

                                var end = temp2.substring(temp2.indexOf("</span"));
                                check_next_li = end.indexOf("</span><ul>") != 0;

                                temp2 = temp2.substring(temp2.indexOf("</span>") + 1);
                            } else if (temp2.indexOf("</ul>") == 0) {
                                j -= 1;
                                temp2 = temp2.substring(1);
                            } else {
                                temp2 = temp2.substring(1);
                            }
                            temp2 = temp2.substring(temp2.indexOf("<"));
                        }
                        console.log(arr);

                        var workbook = XLSX.utils.book_new();
                        var worksheet = XLSX.utils.aoa_to_sheet(arr);
                        workbook.SheetNames.push("First");
                        workbook.Sheets["First"] = worksheet;
                        XLSX.writeFile(workbook, "demo.xlsx");
                    }





                    // let a = document.createElement("A");
                    // a.setAttribute("href", excelData);
                    // a.setAttribute("download", "file.xlsx");
                    // document.body.appendChild(a);
                    // a.click();
                })
                .catch((err) => {
                    console.error(err);
                });
            res.redirect("/download");

            // const uploadRequest = new UploadFileRequest();
            // uploadRequest.path = files.document[0].path;
            // uploadRequest.fileContent = fs.createReadStream(files.document[0].path);

            // wordsAPI.uploadFile(uploadRequest)
            //     .then((_uploadResult) => {
            //         var request = new SaveAsRequest({
            //             name: fileName,
            //             saveOptionsData: new HtmlSaveOptionsData(
            //                 {
            //                     fileName: "destination.html"
            //                 })
            //         });

            //         wordsAPI.saveAs(request)
            //             .then((_result) => {
            //                 // console.log(result);
            //             })

            //         // var request = new DownloadFileRequest({
            //         //     path: "destination.html"
            //         // });

            //         // wordsAPI.downloadFile(request);
            //     })

            // const downloadRequest = new DownloadFileRequest();
            // downloadRequest.path = "destination.html";

            // wordsAPI.downloadFile(downloadRequest)
            //     .then((result) => {
            //         console.log(result);
            //     })

            // res.writeHead(200, { 'content-type': 'text/plain' });
            // res.write('received upload:\n\n');
            // res.end(util.inspect({ fields: fields, files: files }))
        });
    }
});

exports.download_get = asyncHandler(async (req, res, next) => {
    res.download("demo.xlsx");
})

function formatStrInput(input) {

} 