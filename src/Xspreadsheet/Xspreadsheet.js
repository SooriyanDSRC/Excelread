import React, { useState } from 'react';
import Spreadsheet from "x-data-spreadsheet";
import zhCN from 'x-data-spreadsheet/dist/locale/zh-cn';
import { Container, Row, Col, Form, Alert, Tabs, Tab } from 'react-bootstrap';
import XLSX from 'xlsx';

const Xspreadsheet = () => {
    const [alert, setAlert] = useState('');
    const [overallSheetData, setOverallSheetData] = useState(null);
    const [overallSheetData_1, setOverallSheetDatatbl] = useState(null);
    const [file, setFile] = useState({
        fileObj: '',
        fileName: 'Upload Excel',
    });
    const options = {
        mode: 'edit', // edit | read
        showToolbar: true,
        showGrid: true,
        showContextmenu: true,
        view: {
            height: () => document.documentElement.clientHeight,
            width: () => document.documentElement.clientWidth,
        },
        row: {
            len: 100,
            height: 25,
        },
        col: {
            len: 26,
            width: 100,
            indexWidth: 60,
            minWidth: 60,
        },
        style: {
            bgcolor: '#ffffff',
            align: 'left',
            valign: 'middle',
            textwrap: false,
            strike: false,
            underline: false,
            color: '#0a0a0a',
            font: {
                name: 'Helvetica',
                size: 10,
                bold: false,
                italic: false,
            },
        }
    };
    var grid = new Spreadsheet("#x-spreadsheet-demo",options);
    function stox(wb) {
        var out = [];
        wb.SheetNames.forEach(function (name) {
            var o = { name: name, rows: {} };
            var ws = wb.Sheets[name];
            var aoa = XLSX.utils.sheet_to_json(ws, { raw: false, header: 1 });
            aoa.forEach(function (r, i) {
                var cells = {};
                r.forEach(function (c, j) { cells[j] = ({ text: c }); });
                o.rows[i] = { cells: cells };
            })
            out.push(o);
        });
        return out;
    }
    const onFileChangeHandler = (e) => {
        let excelFile = e.target.files[0];
        setOverallSheetData(null);
        if (e.target.files[0]) {
            if (!excelFile.name.match(/\.(xlsx|xls|csv|xlsm)$/)) {
                setAlert(<Alert variant="danger" className="alert_component_css">Please Upload Excel File</Alert>);
                setTimeout(() => setAlert(''), 3000)
            } else {
                
                const data = new Promise(function (resolve, reject) {
                    var reader = new FileReader();
                    var rABS = !!reader.readAsBinaryString;
                    reader.onload = function (e) {
                        var bstr = e.target.result;
                        var wb = XLSX.read(bstr, { type: rABS ? "binary" : "array" });
                        resolve(wb);
                    };
                    if (file && rABS) reader.readAsBinaryString(excelFile);
                    else reader.readAsArrayBuffer(excelFile);
                });
                data.then((exceldata) => {
                    grid.loadData(stox(exceldata));
                });
            }
        }
    }
    /* load data */
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
                        <Form.File name='excel_details' onChange={(e) => onFileChangeHandler(e)} />
                    </Form.Group>
                </Col>
                <Col md={2}></Col>
            </Row>
            <div id="x-spreadsheet-demo">
            </div>
            <div id="gridctr">
            </div>
        </Container>
    )
}

export default Xspreadsheet
