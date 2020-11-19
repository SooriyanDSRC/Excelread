import React, { useState } from 'react'
import { Container, Row, Col, Form, Alert, Tabs, Tab, ListGroup } from 'react-bootstrap';
import { OutTable } from 'react-excel-renderer';
import DataGrid from 'react-data-grid';
import 'react-data-grid/dist/react-data-grid.css';
import './FileImport.css';
import LineTo, { SteppedLineTo } from 'react-lineto';
import '@grapecity/spread-sheets-react';
// import ReactDataSheet from 'react-datasheet';
// Be sure to include styles at some point, probably during your bootstrapping
// import 'react-datasheet/lib/react-datasheet.css';
import Spreadsheet from "react-spreadsheet";
import * as SpreadsheetComponent from 'react-spreadsheet-component';
/* eslint-disable */
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import { SpreadSheets, Worksheet, Column } from '@grapecity/spread-sheets-react';
import { IO } from "@grapecity/spread-excelio";
import { saveAs } from 'file-saver';
// import * as XLSX from 'xlsx';
import ExcelJS from "exceljs";
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
    const [overallSheetData_1, setOverallSheetDatatbl] = useState(null);
    const [sheet_columns, setSheetColumns] = useState(null);
    const [excel_renderer, setExcelRenderer] = useState(null);
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
    const datagrid_columns = [
        { key: 'id', name: '' },
        { key: 'id', name: 'IDdata' },
        { key: 'title', name: 'Title' }
    ];

    const datagrid_rows = [
        { id: 0, title: 'Example' },
        { id: 1, title: 'Demo' }
    ];
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
                console.log(excelFile);
                // if (excelFile.name.match(/\.(xlsm)$/)) {
                //     const wb = new ExcelJS.Workbook();
                //     const reader = new FileReader();
                //     reader.readAsArrayBuffer(excelFile);
                //     reader.onload = () => {
                //         const buffer = reader.result;
                //         var workbook_data = new Array();
                //         wb.xlsx.load(buffer).then(workbook => {
                //             console.log(workbook, "workbook instance");
                //             workbook.eachSheet((sheet, id) => {
                //                 workbook_data[id - 1] = new Array();
                //                 console.log(id - 1);
                //                 sheet.eachRow((row, rowIndex) => {
                //                     // console.log(rowIndex);
                //                     workbook_data[id - 1][rowIndex - 1] = row.values;
                //                 });
                //             });
                //             console.log(JSON.stringify(workbook_data));
                //         });
                //     };
                // } else {
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
                        setExcelRenderer(allsheets);
                        resolve(allsheets);
                        // console.log(allsheets);
                    };
                    if (file && rABS) reader.readAsBinaryString(excelFile);
                    else reader.readAsArrayBuffer(excelFile);
                });
                data.then((overall_data) => {
                    if (overall_data) {
                        var allsheets = [];
                        // var all_sheets_1 = [];
                        overall_data.forEach(element => {
                            let current_sheet_rows = element.rows;
                            var allrows = [];
                            // var all_rows_1 = [];
                            var allcols = element.cols;
                            var sheetname = element.sheetName;
                            current_sheet_rows.forEach(current_sheet_row_element => {
                                var row_obj = [];
                                var row_obj_1 = {};
                                allcols.map((col_element) => {
                                    if (current_sheet_row_element[col_element.key]) {
                                        row_obj_1[col_element.name] = current_sheet_row_element[col_element.key];
                                        row_obj.push({ value: current_sheet_row_element[col_element.key] });
                                    } else {
                                        row_obj_1[col_element.name] = null;
                                        row_obj.push({ value: null });
                                    }
                                });
                                all_rows_1.push(row_obj);
                                allrows.push(row_obj_1);
                            });
                            var row_col_data = { allrows, allcols, sheetname };
                            var row_col_data_1 = { all_rows_1, allcols, sheetname };
                            all_sheets_1.push(row_col_data_1);
                            // console.log(all_sheets_1);
                            allsheets.push(row_col_data);
                        });
                        setOverallSheetData(allsheets);
                        var allsheet_cols = [];
                        allsheets.map((sheet) => {
                            var current_sheet = {};
                            current_sheet["sheetname"] = sheet.sheetname;
                            current_sheet["cols"] = [];
                            const isNotEmpty = sheet.allrows.map((current_row) => {
                                if (Object.values(current_row).every(x => (x !== null && x !== ''))) {
                                    Object.values(current_row).every(x => (current_sheet["cols"].push(x)));
                                }
                            })
                            // console.log(isNotEmpty);
                            allsheet_cols.push(current_sheet);
                        })
                        setSheetColumns(allsheet_cols);
                        // setOverallSheetDatatbl(all_sheets_1);
                        // console.log(all_sheets_1);
                    }
                });
                // }
            }
        }

    }
    // const exportSheet = () => {
    //     const spread = _spread;
    //     const fileName = "Sample3.xlsx";
    //     const sheet = spread.getSheet(0);
    //     const excelIO = new IO();
    //     const json = JSON.stringify(spread.toJSON({
    //         includeBindingSource: true,

    //     }));
    //     excelIO.save(json, (blob) => {
    //         saveAs(blob, fileName);
    //     }, function (e) {
    //         alert(e);
    //     });
    // }
    // console.log(JSON.stringify(overallSheetData_1));
    // console.log(JSON.stringify(overallSheetData));
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
                <Col></Col>
                <Col><DataGrid
                    columns={datagrid_columns}
                    rows={datagrid_rows}
                    className="rdg-light"
                /></Col>
                <Col></Col>
            </Row>
            <Row>
                <Col md={12}>
                    {overallSheetData && <SpreadSheets hostClass={config.hostClass} workbookInitialized={workbookInit}>
                        {overallSheetData.map((current_sheet) => {
                            const { allcols, allrows, sheetname } = current_sheet;
                            return (<Worksheet key={sheetname} name={sheetname} dataSource={allrows} autoGenerateColumns={config.autoGenerateColumns}>
                                {allcols.map((current_col) => (
                                    <Column width={100} dataField={current_col.name} headerText={current_col.name}></Column>
                                ))}
                            </Worksheet>)
                        })}
                    </SpreadSheets>}
                </Col>
            </Row>
            {/* {overallSheetData && <Row>
                <Col className="text-center m-10">
                    <div className="dashboardRow">
                        <button className="btn btn-primary dashboardButton"
                            onClick={exportSheet}>Export to Excel</button>
                    </div>
                </Col>
            </Row>} */}
            {
                overallSheetData_1 && <Row >
                    <Col>
                        <Tabs defaultActiveKey={overallSheetData_1[0].sheetname} id="uncontrolled-tab-example" >
                            {overallSheetData_1.map((current_sheet) => (
                                <Tab eventKey={current_sheet.sheetname} title={current_sheet.sheetname} className="excel_component">
                                    <Spreadsheet data={current_sheet.all_rows_1} />
                                </Tab>
                            ))}
                        </Tabs>

                    </Col>
                </Row>
            }
            {
                sheet_columns && <Row>
                    <Col md={5}>
                        <Tabs defaultActiveKey={sheet_columns[0].sheetname} id="uncontrolled-tab-example" >
                            {sheet_columns.map((current_sheet) => (
                                <Tab eventKey={current_sheet.sheetname} key={current_sheet.sheetname} title={current_sheet.sheetname} className="">
                                    <ListGroup>
                                        {current_sheet.cols.map((current_col) => (
                                            <ListGroup.Item classname="A">{current_col}</ListGroup.Item>
                                        ))}
                                    </ListGroup>
                                </Tab>
                            ))}
                        </Tabs>
                    </Col>
                    <Col md={7}>
                        <ListGroup>
                            <ListGroup.Item className="B">Joint Number</ListGroup.Item>
                            <ListGroup.Item>Feature Number</ListGroup.Item>
                            <ListGroup.Item>Cluster Number</ListGroup.Item>
                            <ListGroup.Item>Feature Type</ListGroup.Item>
                        </ListGroup>
                    </Col>
                    <SteppedLineTo from="A" to="B" orientation="v" />
                </Row>
            }
            <div style={{ width: 50, height: 50, marginRight: 40, background: 'white' }} className="c1" />
            <div style={{ width: 50, height: 50, background: 'white' }} className="c2" />
            <LineTo from="c1" to="c2" />
        </Container>
    )
}

export default FileImport
