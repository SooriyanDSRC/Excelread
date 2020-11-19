import './App.css';
import React, { Fragment } from 'react';
import { BrowserRouter as Router, Route, Switch } from 'react-router-dom';
import FileImport from './FileImport/FileImport';
import ReactSpreadsheet from './ReactSpreadsheet/ReactSpreadsheet';
import ReactDataGrid from './ReactDataGrid/ReactDataGrid';
import Xspreadsheet from './Xspreadsheet/Xspreadsheet';
import FieldMappingSample from './FieldMappingSample/FieldMappingSample';
import FileViewer from './FileViewer/FileViewer';
function App() {
  return (
    <Fragment>
      <Router>
        <Switch>
          <Route exact path='/' component={FileImport}/>
          <Route exact path='/sprdsheet' component={ReactSpreadsheet}/>
          <Route exact path='/datagrid' component={ReactDataGrid}/>
          <Route exact path='/xspreadsheet' component={Xspreadsheet}/>
          <Route exact path="/fieldmappingsample" component={FieldMappingSample}/>
          <Route exact path='/excelview' component={FileViewer}/>
        </Switch>
      </Router>
    </Fragment>
  );
}
export default App;
