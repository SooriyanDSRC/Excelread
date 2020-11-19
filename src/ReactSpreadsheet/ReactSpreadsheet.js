import React, { useState } from 'react'
import { Container, Row, Col, Form, Alert, Tabs, Tab } from 'react-bootstrap';
import Spreadsheet from "react-spreadsheet";
import XLSX from 'xlsx';

const ReactSpreadsheet = () => {
    const [alert, setAlert] = useState('');
    const [overallSheetData, setOverallSheetData] = useState(null);
    const [overallSheetData_1, setOverallSheetDatatbl] = useState(null);
    const [file, setFile] = useState({
        fileObj: '',
        fileName: 'Upload Excel',
    });
    const { fileObj, fileName } = file;
    function make_cols(refstr) {
        var o = [],
            C = XLSX.utils.decode_range(refstr).e.c + 1;
        for (var i = 0; i < C; ++i) {
            o[i] = { name: XLSX.utils.encode_col(i), key: i };
        }
        return o;
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
                        // console.log(allsheets);
                    };
                    if (file && rABS) reader.readAsBinaryString(excelFile);
                    else reader.readAsArrayBuffer(excelFile);
                });
                data.then((overall_data) => {
                    if (overall_data) {
                        var allsheets = [];
                        var all_sheets_1 = [];
                        overall_data.forEach(element => {
                            let current_sheet_rows = element.rows;
                            var all_rows_1 = [];
                            var allcols = element.cols;
                            var sheetname = element.sheetName;
                            current_sheet_rows.forEach(current_sheet_row_element => {
                                var row_obj = [];
                                allcols.map((col_element) => {
                                    if (current_sheet_row_element[col_element.key]) {
                                        row_obj.push({ value: current_sheet_row_element[col_element.key] });
                                    } else {
                                        row_obj.push({ value: null });
                                    }
                                });
                                all_rows_1.push(row_obj);
                            });
                            var row_col_data_1 = { all_rows_1, allcols, sheetname };
                            all_sheets_1.push(row_col_data_1);
                        });
                        setOverallSheetDatatbl(all_sheets_1);
                    }
                });
            }
        }

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
        </Container>

    )
}

export default ReactSpreadsheet
