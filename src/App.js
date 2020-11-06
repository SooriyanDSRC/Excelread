import './App.css';
import React, { Fragment } from 'react';
import { BrowserRouter as Router, Route, Switch } from 'react-router-dom';
import FileImport from './FileImport/FileImport';
function App() {
  return (
    <Fragment>
      <Router>
        <Switch>
          <Route exact path='/' component={FileImport}/>
        </Switch>
      </Router>
    </Fragment>
  );
}

export default App;
