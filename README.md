# FAR/IP Data Extraction Tool

## Introduction

This application is for the use of the Office of Academic Programs, University of Toronto. 

The Data Extraction Tool scrapes the Final Assessment Report and Implementation Plan documents (FAR/IPs) and outputs the review items in .xlsx files for analysis purposes. 

## Usage

### Online

A deployed application can be accessed at https://farip-extraction.up.railway.app

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

- Only files in **.pdf/.docx** format are supported.
- The interpreted FAR/IP documents must include the following session headings:
   - **Current Review: Findings and Recommendations**
   - **2. Administrative Response & Implementation Plan** (for full-length pdf documents)
- Up to **150** file conversions (including bad requests) per month can be made due to limited usage of third-party API.
- Files with unusual formats (i.e. triangular or square sub-list points) may fail to extract data or miss sub-list items. Try converting the input document into another acceptable format before the data extraction.
- Unsuccessful data extraction will direct to an error page. Follow the indicated instructions and proceed.

## Milestones

1. Plan and design the application
2. Implement the foundational functionality: scrape the .docx FAR/IPs
3. Develop to support the .pdf documents
4. Create error pages for unsuccessful extractions
5. Deploy the application online with a user-friendly UI
6. Develop documentation and instructions
7. Debug to ensure robustness

## Author

[Tingyu Liang](https://github.com/liangt2001)
