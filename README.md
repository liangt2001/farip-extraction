# FAR/IP Data Extraction Tool

## Introduction

This application is for the use of the Office of Academic Programs, University of Toronto. 

The Data Extraction Tool scrapes off the Final Assessment Report and Implementation Plan documents (FAR/IPs) and outputs the review items in a .xlsx file for analysis purposes. 

## Usage

### Online

A deployed application can be accessed at https://farip-extraction.up.railway.app.

### Local

Clone the repository, install dependencies, and start the server locally.

```
git clone https://github.com/liangt2001/farip-extraction.git
cd farip-extraction
npm i
npm start
```
Open `http://localhost:3000` in the browser.

## Limitations

1. Only files in **.pdf/.docx** format are supported.
2. The interpreted FAR/IP documents must include the following session headings:
   - **Current Review: Findings and Recommendations**
   - **2. Administrative Response & Implementation Plan**
3. Sub-list items are sometimes missed.

## Author

[Tingyu Liang](https://github.com/liangt2001)
