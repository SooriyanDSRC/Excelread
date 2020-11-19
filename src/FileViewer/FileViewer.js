import React, { useState } from 'react'
import { Container, Row, Col, Form, Alert, Tabs, Tab, ListGroup } from 'react-bootstrap';

const FileViewer = () => {
    const [file, setFile] = useState({
        fileObj: '',
        fileName: 'Upload Excel',
        fileType: ''
    });
    const onFileChangeHandler = (e) => {
        let excelFile = e.target.files[0];
        var file_type = (/[.]/.exec(excelFile.name)) ? /[^.]+$/.exec(excelFile.name) : undefined;

    }

    return (
        <Container>
            {/* <Row>
                <Col>
                <Form.Label>
                            <b>Upload Excel</b>
                        </Form.Label>
                        <Form.File name='excel_details' onChange={(e) => onFileChangeHandler(e)} label={fileName} />

                </Col>
            </Row> */}
            <Row>
                <Col md={2}></Col>
                <Col md={8}>
                    <FileViewer
                        fileType='xlsx'
                        filePath='C:\Users\sooriyan.p\Documents\Sample ILI Excel\Sample 1 ILI Tethered Report.xlsx'
                        errorComponent={<div>Error</div>}
                    />
                </Col>
                Hii
            </Row>
        </Container>
    )
}

export default FileViewer
