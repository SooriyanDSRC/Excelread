import React, { useState } from 'react'
import { Container, Row, Col, Form, Alert } from 'react-bootstrap';
// import { OutTable, ExcelRenderer } from 'react-excel-renderer';
import './FileImport.css';
import '@grapecity/spread-sheets-react';
import Spreadsheet from "react-spreadsheet";
/* eslint-disable */
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import { SpreadSheets, Worksheet, Column } from '@grapecity/spread-sheets-react';
import {IO} from "@grapecity/spread-excelio";
import { saveAs } from 'file-saver';
// import * as XLSX from 'xlsx';
import XLSX from 'xlsx';
const FileImport = () => {
    const [file, setFile] = useState({
        fileObj: '',
        fileName: 'Upload Excel',
    });
    const config = {
        sheetName: 'Sales Data',
        hostClass: ' spreadsheet',
        autoGenerateColumns: false,
        width: 200,
        visible: true,
        resizable: true,
        priceFormatter: '$ #.00',
        chartKey: 1
    }
    const [overallSheetData, setOverallSheetData] = useState(null);
    const [_spread, setSpread] = useState({});
    const [alert, setAlert] = useState('');
    const { fileObj, fileName } = file;
    function make_cols(refstr) {
        var o = [],
            C = XLSX.utils.decode_range(refstr).e.c + 1;
        for (var i = 0; i < C; ++i) {
            o[i] = { name: XLSX.utils.encode_col(i), key: i };
        }
        return o;
    }
    function workbookInit(spread) {
        setSpread(spread)
    }
    const onFileChangeHandler = (e) => {
        let excelFile = e.target.files[0];
        setOverallSheetData(null);
        if (e.target.files[0]) {
            if (!excelFile.name.match(/\.(xlsx|xls|csv|xlsm)$/)) {
                setAlert(<Alert variant="danger" className="alert_component_css">Please Upload Excel File</Alert>);
                setTimeout(() => setAlert(''), 3000)
            } else {
                setFile({
                    fileObj: excelFile,
                    fileName: excelFile.name
                });
                const data = new Promise(function (resolve, reject) {
                    var reader = new FileReader();
                    var rABS = !!reader.readAsBinaryString;
                    reader.onload = function (e) {
                        var bstr = e.target.result;
                        var wb = XLSX.read(bstr, { type: rABS ? "binary" : "array" });
                        const allsheets = wb.SheetNames.map((current_sheet) => {
                            var ws = wb.Sheets[current_sheet];
                            var json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                            var cols = make_cols(ws["!ref"]);
                            var data = { rows: json, cols: cols, sheetName: current_sheet };
                            return data;
                        });
                        resolve(allsheets);
                    };
                    if (file && rABS) reader.readAsBinaryString(excelFile);
                    else reader.readAsArrayBuffer(excelFile);
                });
                data.then((overall_data) => {
                    if (overall_data) {
                        var allsheets = [];
                        overall_data.forEach(element => {
                            let current_sheet_rows = element.rows;
                            var allrows = [];
                            var allcols = element.cols;
                            var sheetname = element.sheetName;
                            current_sheet_rows.forEach(current_sheet_row_element => {
                                var row_obj = {};
                                current_sheet_row_element.forEach(most_nested_row_element => {
                                    let current_index = element.cols.filter(obj => {
                                        return obj.key === current_sheet_row_element.indexOf(most_nested_row_element);
                                    });
                                    let index_to_be_allotted = current_index[0].name;
                                    let current_mnrow_obj = {};
                                    current_mnrow_obj[index_to_be_allotted] = most_nested_row_element;
                                    row_obj[index_to_be_allotted] = most_nested_row_element;
                                });
                                allrows.push(row_obj);
                            });
                            var row_col_data = { allrows, allcols, sheetname };
                            allsheets.push(row_col_data);
                        });
                        setOverallSheetData(allsheets);
                        console.log(JSON.stringify(allsheets));
                    }
                });
            }
        }

    }
    const exportSheet = () => {
        const spread = _spread;
        const fileName = "Sample3.xlsx";
        const sheet = spread.getSheet(0);
        const excelIO = new IO();
        const json = JSON.stringify(spread.toJSON({ 
            includeBindingSource: true,
            
        }));
        excelIO.save(json, (blob) => {
            saveAs(blob, fileName);
        }, function (e) {  
            alert(e);  
        });     
    }
    return (
        <Container>
            <Row>
                <Col md={2}></Col>
                <Col md={8}>
                    {alert}
                    <Form.Group controlId='formBasicTest'>
                        <Form.Label>
                            <b>Upload Excel</b>
                        </Form.Label>
                        <Form.File name='excel_details' onChange={(e) => onFileChangeHandler(e)} label={fileName} />
                    </Form.Group>
                </Col>
                <Col md={2}></Col>
            </Row>
            <Row>
                <Col md={12}>
                    {overallSheetData && <SpreadSheets hostClass={config.hostClass} workbookInitialized={workbookInit}>
                        {overallSheetData.map((current_sheet) => {
                            const { allcols, allrows, sheetname } = current_sheet;
                            return (<Worksheet key={sheetname} name={sheetname} dataSource={allrows} autoGenerateColumns={config.autoGenerateColumns}>
                                {allcols.map((current_col) => (
                                    <Column width={100}  dataField={current_col.name} headerText={current_col.name}></Column>
                                ))}
                            </Worksheet>)
                        })}
                    </SpreadSheets>}
                </Col>
            </Row>
            {overallSheetData && <Row>
                <Col className="text-center m-10">
                    <div className="dashboardRow">
                        <button className="btn btn-primary dashboardButton"
                            onClick={exportSheet}>Export to Excel</button>
                    </div>
                </Col>
            </Row>}
        </Container>
    )
}

export default FileImport
