const asyncHandler = require("express-async-handler");
const fs = require("fs");
const multiparty = require("multiparty");
const { WordsApi, ConvertDocumentRequest } = require("asposewordscloud");
const xlsx = require("xlsx");

var outputFileName;

exports.index_post = asyncHandler(async (req, res, next) => {
    if (req.url === "/import" && req.method === "POST") {

        var form = new multiparty.Form();
        form.parse(req, function (err, fields, files) {

            const clientID = "59611537-80f3-4efd-85e2-30be22ac2e9d";
            const secret = "184c09e6e580ae01ea5ae83e7fa3bd08";
            const wordsAPI = new WordsApi(clientID, secret);

            console.log(files.document[0]);
            var fileName = files.document[0].originalFilename;
            var filePath = files.document[0].path;

            var request = new ConvertDocumentRequest();
            request.document = fs.createReadStream(filePath);
            request.format = "html";

            wordsAPI.convertDocument(request)
                .then((response) => {
                    result = response.body.toString();

                    if (filePath.match(/\.pdf\b/g)) {
                        console.log(result.search(/Current Review: Findings and Recommendation/gs));
                        console.log(result.search(/Response &amp; Implementation/g));
                        // console.log(result);
                        var text2_start = result.search(/Current Review: Findings and Recommendation/gs);
                        var text2 = result.substring(text2_start);
                        if (text2_start == -1) {
                            console.error("Cannot trim the text");
                        }
                        var text2_end_1 = text2.search(/Response &amp; Implementation/g);
                        var text2_end_2 = text2.indexOf("<img");
                        if (text2_end_1 == -1 & text2_end_2 == -1) {
                            console.error("Cannot trim the text");
                        } else if (text2_end_1 == -1) {
                            text2 = text2.substring(0, text2_end_2);
                        } else if (text2_end_2 == -1) {
                            text2 = text2.substring(0, text2_end_1);
                        } else {
                            text2 = text2.substring(0, Math.min(text2_end_1, text2_end_2));
                        }
                        console.log(text2);
                        var temp2 = text2;
                        var arr = new Array();
                        var i = -1;
                        var j = 0;

                        while (temp2.indexOf("<") >= 0) {
                            if (temp2.indexOf("<ol") == 0) {
                                i += 1;
                                j = 0;
                                var raw_text = temp2.substring(0, temp2.indexOf("</ol>"));
                                var val = processText(raw_text);
                                arr.push(new Array(5));
                                arr[i][j] = val;
                                temp2 = temp2.substring(temp2.indexOf("</ol>"));
                            }
                            else if (temp2.indexOf("<h") == 0
                                & temp2.indexOf(`<span style="font-family:'Lucida Bright'; font-size:13.5pt`) == temp2.indexOf(">") + 1) {
                                i += 1;
                                j = 0;
                                var raw_text = temp2.substring(0, temp2.indexOf("</h"));
                                var val = processText(raw_text);
                                arr.push(new Array(5));
                                arr[i][j] = val.replaceAll(/[0-9.]/g,"").trim();
                                temp2 = temp2.substring(temp2.indexOf("</h"));
                            }
                            else if (temp2.indexOf("<p") == 0) {
                                if (temp2.indexOf(`<span style="font-family:'Lucida Bright'; font-size:13.5pt`) == temp2.indexOf(">") + 1) {
                                    i += 1;
                                    j = 0;
                                    var raw_text = temp2.substring(0, temp2.indexOf("</p>"));
                                    var val = processText(raw_text);
                                    arr.push(new Array(5));
                                    arr[i][j] = val.replaceAll(/[0-9.]/g,"").trim();
                                    temp2 = temp2.substring(temp2.indexOf("</p>"));
                                    continue;
                                }
                                var raw_text = temp2.substring(0, temp2.indexOf("</p>"));
                                var val = processText(raw_text);
                                var is_sar = (val.search(/following.+?strengths/g) != -1) | (val.search(/following.+?areas of concern/g) != -1) | (val.search(/following.+?recommendation/g) != -1);
                                if (is_sar == 1) {
                                    if (j >= 1) {
                                        i += 1;
                                        arr.push(new Array(5));
                                        arr = copyFromAbove(arr, i, 1);
                                    }
                                    j = 1;
                                    if (val.includes("strength")) {
                                        arr[i][j] = "Strengths";
                                    }
                                    else if (val.includes("areas of concern")) {
                                        arr[i][j] = "Areas of Concern";
                                    }
                                    else if (val.includes("recommendation")) {
                                        arr[i][j] = "Recommendations";
                                    }
                                }
                                else if (val[0] == "•") {
                                    if (j >= 2) {
                                        i += 1;
                                        arr.push(new Array(5));
                                        arr = copyFromAbove(arr, i, 2);
                                    }
                                    j = 2;
                                    arr[i][j] = val.substring(1).trim();
                                }
                                else if (val[0] == "" | val[0] == "o") {
                                    if (j >= 3) {
                                        i += 1;
                                        arr.push(new Array(5));
                                        arr = copyFromAbove(arr, i, 3);
                                    }
                                    j = 3;
                                    arr[i][j] = val.substring(1).trim();
                                }
                                else {
                                    console.error(val);
                                }
                                temp2 = temp2.substring(temp2.indexOf("</p>"));
                            }
                            else if (temp2.indexOf('<ul type="square"') == 0) {
                                var raw_text = temp2.substring(0, temp2.indexOf("</li>"));
                                var val = processText(raw_text);
                                if (j >= 4) {
                                    i += 1;
                                    arr.push(new Array(5));
                                    arr = copyFromAbove(arr, i, 4);
                                }
                                j = 4;
                                arr[i][j] = val;
                                if (temp2.indexOf("<li", temp2.indexOf("</li>")) != -1) {
                                    temp2 = temp2.substring(Math.min(temp2.indexOf("</ul>"), temp2.indexOf("<li", temp2.indexOf("</li>")) - 1));
                                } else {
                                    temp2 = temp2.substring(temp2.indexOf("</ul>"));
                                }
                            }
                            else if (temp2.indexOf('<ul type="circle"') == 0) {
                                var raw_text = temp2.substring(0, temp2.indexOf("</li>"));
                                var val = processText(raw_text);
                                if (j >= 3) {
                                    i += 1;
                                    arr.push(new Array(5));
                                    arr = copyFromAbove(arr, i, 3);
                                }
                                j = 3;
                                arr[i][j] = val;
                                if (temp2.indexOf('<ul type="square') != -1) {
                                    temp2 = temp2.substring(Math.min(temp2.indexOf('<ul type="square') - 1, temp2.indexOf("</li>")));
                                } else {
                                    temp2 = temp2.substring(temp2.indexOf("</li>"));
                                }
                            }
                            else if (temp2.indexOf("<li") == 0) {
                                var raw_text = temp2.substring(0, temp2.indexOf("</li>"));
                                var val = processText(raw_text);
                                i += 1;
                                arr.push(new Array(5));
                                arr = copyFromAbove(arr, i, j);
                                arr[i][j] = val;
                                temp2 = temp2.substring(temp2.indexOf("</li>"));
                            }
                            temp2 = temp2.substring(1);
                            temp2 = temp2.substring(temp2.indexOf("<"));
                        }
                        console.log(arr);

                        // convert arr to xlsx
                        var workbook = xlsx.utils.book_new();
                        var worksheet = xlsx.utils.aoa_to_sheet(arr);
                        workbook.SheetNames.push("Main");
                        workbook.Sheets["Main"] = worksheet;
                        outputFileName = fileName.substring(0, fileName.length - 3) + "xlsx";
                        xlsx.writeFile(workbook, outputFileName);

                        res.redirect("/download");
                    }
                    else if (filePath.match(/\.docx\b/g)) {
                        var text2;
                        var text2_start1 = result.search(/Findings &amp; Recommendations/gs);
                        var text2_start2 = result.search(/FINDINGS AND RECOMMENDATION/gs);
                        if (text2_start1 != -1 & text2_start2 == -1) {
                            text2 = result.substring(text2_start1);
                        } else if (text2_start2 != -1 & text2_start1 == -1) {
                            text2 = result.substring(text2_start2);
                        } else {
                            console.error("Ambiguous start point");
                        }
                        var temp2 = text2;
                        console.log(result);
                        var n_rows = (text2.match(/<\/span><\/li>/gm) || []).length - 2;
                        var check_next_p = false;
                        var check_next_li = false;
                        var arr = new Array(n_rows);
                        for (var i = 0; i < arr.length; i++) {
                            arr[i] = new Array(5);
                        }

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
                })
                .catch((err) => {
                    console.error(err);
                })
            // res.redirect("/download");
        })
    }
});

exports.download_get = asyncHandler(async (req, res, next) => {
    res.download(outputFileName);
});

function processText(raw_text) {
    var result = "";
    while (raw_text.indexOf("<span") >= 0) {
        var str_start = raw_text.search(/>[^<]/g) + 1;
        var str_end = raw_text.indexOf("</span>", str_start)
        var str = raw_text.substring(str_start, str_end);
        result += str;
        raw_text = raw_text.substring(str_end);
    }
    result = result.replaceAll("&#xa0;", "").replaceAll("&amp;", "&").trim();
    return result;
}

function copyFromAbove(arr, i, j) {
    if (!(arr[i][j - 1] == null) | i - 1 < 0) {
        console.error("Tried copyFromAbove but already filled or no above");
        return;
    }
    for (var temp_j = 0; temp_j < j; temp_j++) {
        arr[i][temp_j] = arr[i - 1][temp_j];
    }
    return arr;
}