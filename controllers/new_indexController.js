const asyncHandler = require("express-async-handler");
const fs = require("fs");
const path = require("path");
const multiparty = require("multiparty");
const { WordsApi, ConvertDocumentRequest } = require("asposewordscloud");
const xlsx = require("xlsx");

var outputFileName;

exports.index_get = asyncHandler(async (req, res, next) => {
    res.render("index", { title: "FAR/IP Data Extraction Tool" });
});

exports.index_post = asyncHandler(async (req, res, next) => {
    if (req.url === "/import" && req.method === "POST") {

        var form = new multiparty.Form();
        form.parse(req, function (err, fields, files) {

            const clientID = "ae9136a2-f5a9-47e8-b2ed-682e16cfaf3a";
            const secret = "335c530ecfac6965943ba418b4643313";
            const wordsAPI = new WordsApi(clientID, secret);

            console.log(files.document[0]);
            var fileName = files.document[0].originalFilename;
            var filePath = files.document[0].path;

            if (!filePath.match(/\.pdf\b/g) & !filePath.match(/\.docx\b/g)) {
                console.log(filePath);
                // alert("File format not supported.");
                res.redirect("/");
            }

            var request = new ConvertDocumentRequest();
            request.document = fs.createReadStream(filePath);
            request.format = "html";

            wordsAPI.convertDocument(request)
                .then((response) => {
                    result = response.body.toString();

                    if (filePath.match(/\.pdf\b/g)) {
                        var text2_start = result.search(/Current Review: Findings and Recommendation/gs);
                        var text2 = result.substring(text2_start);
                        if (text2_start == -1) {
                            console.error("Cannot trim the text");
                            return res.render("error",
                                {
                                    title: "Extraction Failed",
                                    message1: "Error: the submitted file does not include the required headings",
                                    message2: "Please go back and submit another file"
                                });
                        }
                        var text2_end_1 = text2.search(/Response &amp; Implementation/g);
                        var text2_end_2 = text2.indexOf("<img");
                        if (text2_end_1 == -1 & text2_end_2 == -1) {
                            console.error("Cannot trim the text");
                            return res.render("error",
                                {
                                    title: "Extraction Failed",
                                    message1: "Error: the submitted file does not include the required headings",
                                    message2: "Please go back and submit another file"
                                });
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
                                arr[i][j] = val.replaceAll(/[0-9.]/g, "").trim();
                                temp2 = temp2.substring(temp2.indexOf("</h"));
                            }
                            else if (temp2.indexOf("<p") == 0) {
                                if (temp2.indexOf(`<span style="font-family:'Lucida Bright'; font-size:13.5pt`) == temp2.indexOf(">") + 1) {
                                    i += 1;
                                    j = 0;
                                    var raw_text = temp2.substring(0, temp2.indexOf("</p>"));
                                    var val = processText(raw_text);
                                    arr.push(new Array(5));
                                    arr[i][j] = val.replaceAll(/[0-9.]/g, "").trim();
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
                                    return res.render("error",
                                        {
                                            title: "Extraction Failed",
                                            message1: "Error: characters cannot be identified",
                                            message2: "This file may contain illegal characters. Please go back and submit the file in another acceptable format"
                                        });
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

                        return res.redirect("/download");
                    }
                    else if (filePath.match(/\.docx\b/g)) {
                        var text2;
                        var text2_start1 = result.search(/Current Review: Findings &amp; Recommendations/gs);
                        var text2_start2 = result.search(/Current Review: Findings and Recommendations/gs);
                        var text2_start3 = result.search(/FINDINGS AND RECOMMENDATION/gs);
                        if (text2_start1 != -1) {
                            text2 = result.substring(text2_start1);
                        } else if (text2_start2 != -1) {
                            text2 = result.substring(text2_start2);
                        } else if (text2_start3 != -1) {
                            text2 = result.substring(text2_start3);
                        } else {
                            console.error("Ambiguous start point");
                            return res.render("error",
                                {
                                    title: "Extraction Failed",
                                    message1: "Error: the submitted file does not include the required headings",
                                    message2: "Please go back and submit another file"
                                });
                        }
                        var temp2 = text2;
                        var arr = new Array();
                        var i = -1;
                        var j = 0;
                        console.log(text2_start1)
                        console.log(text2_start2)
                        console.log(text2_start3)
                        console.log(temp2);

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
                            else if (temp2.indexOf("<h") == 0) {
                                var raw_text = temp2.substring(0, temp2.indexOf("</h"));
                                if (!raw_text.includes(`<span style="font-family:'Lucida Bright'`)) {
                                    temp2 = temp2.indexOf("</h");
                                } else {
                                    i += 1;
                                    j = 0;
                                    arr.push(new Array(5));
                                    var val = processText(raw_text);
                                    arr[i][j] = val.replaceAll(/[0-9.]/g, "").trim();
                                    temp2 = temp2.substring(temp2.indexOf("</h"));
                                }
                            }
                            else if (temp2.indexOf("<p") == 0) {
                                console.log(temp2.substring(0, 100));
                                var raw_text = temp2.substring(0, temp2.indexOf("</p>"));
                                console.log("raw = " + raw_text);
                                var val = processText(raw_text);
                                var is_sar = (val.search(/following.+?strengths/g) != -1) | (val.search(/following.+?areas of concern/g) != -1) | (val.search(/following.+?recommendation/g) != -1);
                                var is_review_item = (val.includes("1. Undergraduate Program") | val.includes("2. Graduate Program") | val.includes("3. Faculty/Research") | val.includes("4. Administration"));
                                console.log("val = " + val);
                                if (val == "Administrative response—appended" | val == "ADMINISTRATIVE RESPONSE – Appended" | typeof (temp2) != "string" | temp2.indexOf("<div") == 0) {
                                    console.log("breakkkkk!")
                                    break;
                                }
                                if (is_review_item) {
                                    i += 1;
                                    j = 0;
                                    arr.push(new Array(5));
                                    if (val.includes("(")) {
                                        arr[i][j] = val.substring(0, val.indexOf("(")).replaceAll(/[0-9.]/g, "").trim();
                                    } else {
                                        arr[i][j] = val.replaceAll(/[0-9.]/g, "").trim();
                                    }

                                }
                                else if (is_sar == 1) {
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
                                else if (val[0] == "") {
                                    if (j >= 2) {
                                        i += 1;
                                        arr.push(new Array(5));
                                        arr = copyFromAbove(arr, i, 2);
                                    }
                                    j = 2;
                                    arr[i][j] = val.substring(1).trim();
                                }
                                else if (val[0] == "o") {
                                    if (j >= 3) {
                                        i += 1;
                                        arr.push(new Array(5));
                                        arr = copyFromAbove(arr, i, 3);
                                    }
                                    j = 3;
                                    arr[i][j] = val.substring(1).trim();
                                }
                                else if (val[0] == "") {
                                    if (j >= 4) {
                                        i += 1;
                                        arr.push(new Array(5));
                                        arr = copyFromAbove(arr, i, 4);
                                    }
                                    j = 4;
                                    arr[i][j] = val.substring(1).trim();
                                }
                                else {
                                    temp2 = temp2.substring(1);
                                    next;
                                }
                                temp2 = temp2.substring(temp2.indexOf("</p>"));
                                if (val == "") next;
                            }
                            else if (temp2.indexOf('<ul type="circle') == 0) {
                                var raw_text = temp2.substring(0, temp2.indexOf("</li>"));
                                var val = processText(raw_text);
                                if (j >= 3) {
                                    i += 1;
                                    arr.push(new Array(5));
                                    arr = copyFromAbove(arr, i, 3);
                                }
                                j = 3;
                                arr[i][j] = val;
                                if (temp2.indexOf("<li", temp2.indexOf("</li>")) != -1) {
                                    temp2 = temp2.substring(Math.min(temp2.indexOf("</ul>"), temp2.indexOf("<li", temp2.indexOf("</li>")) - 1));
                                } else {
                                    temp2 = temp2.substring(temp2.indexOf("</ul>"));
                                }
                            }
                            else if (temp2.indexOf('<ul type="disc') == 0) {
                                console.log("detected disc");
                                var raw_text;
                                if (temp2.indexOf('<ul type="circle') != -1) {
                                    raw_text = temp2.substring(0, Math.min(temp2.indexOf('<ul type="circle'), temp2.indexOf("</li>")));
                                } else {
                                    raw_text = temp2.substring(0, temp2.indexOf("</li>"));
                                }
                                console.log("raw = " + raw_text);
                                var val = processText(raw_text);
                                console.log(val);
                                if (j >= 2) {
                                    i += 1;
                                    arr.push(new Array(5));
                                    arr = copyFromAbove(arr, i, 2);
                                }
                                j = 2;
                                arr[i][j] = val;
                                if (temp2.indexOf('<ul type="circle') != -1) {
                                    temp2 = temp2.substring(Math.min(temp2.indexOf('<ul type="circle') - 1, temp2.indexOf("</li>")));
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

                            if (val == "Administrative response—appended" | val == "ADMINISTRATIVE RESPONSE – Appended" | typeof (temp2) != "string" | temp2.indexOf("<div") == 0) {
                                console.log("breakkkkk!")
                                break;
                            }
                            temp2 = temp2.substring(1);
                            temp2 = temp2.substring(temp2.indexOf("<"));
                            console.log(temp2.substring(0, 100));
                        }
                        console.log(arr);

                        var workbook = xlsx.utils.book_new();
                        var worksheet = xlsx.utils.aoa_to_sheet(arr);
                        workbook.SheetNames.push("Main");
                        workbook.Sheets["Main"] = worksheet;
                        outputFileName = fileName.substring(0, fileName.length - 4) + "xlsx";
                        xlsx.writeFile(workbook, outputFileName);

                        return res.redirect("/download");
                    }
                })
                .catch((err) => {
                    console.error(err);
                    return res.render("error",
                        {
                            title: "Extraction Failed",
                            message1: "Error: an unidentified error occurs during the conversion",
                            message2: "Please go back and submit another file"
                        });
                })
            // res.redirect("/download");
        })
    }
});

exports.download_get = asyncHandler(async (req, res, next) => {
    res.download(outputFileName, (err) => {
        if (err) {
            console.error(err);
            return res.render("error",
                {
                    title: "Extraction Failed",
                    message1: "Error: an unidentified error occurs during the conversion",
                    message2: "Please go back and submit another file"
                });
        }
        else {
            fs.unlink(outputFileName, (error) => {
                console.error(error);
                return res.render("error",
                    {
                        title: "Extraction Failed",
                        message1: "Error: an unidentified error occurs during the conversion",
                        message2: "Please go back and submit another file"
                    });
            })
        }
    });
});

function processText(raw_text) {
    var result = "";
    while (raw_text.search(/<span.*?>[^<]/g) >= 0) {
        var str_start = raw_text.search(/>[^<]/g) + 1;
        var str_end = raw_text.indexOf("</span>", str_start);
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
        return res.render("error",
            {
                title: "Extraction Failed",
                message1: "Error: an unidentified error occurs during the conversion",
                message2: "Please go back and submit another file"
            });
    }
    for (var temp_j = 0; temp_j < j; temp_j++) {
        arr[i][temp_j] = arr[i - 1][temp_j];
    }
    return arr;
}