import React from 'react';
import loader from '@ibsheet/loader';
import { BrowserRouter, Route, Link } from 'react-router-dom';
import { SheetSample } from './components/SheetSample';
import { Home } from './components/Home';

import './App.scss';
import { FastLoadSheet } from './components/FastLoad';
import { LazyLoadSheet } from './components/LazyLoad';

loader.config({
  registry: [{
    name: 'ibsheet',
    // baseUrl: 'https://www.ibsheet.com/ibsheet8/customer-sample/assets/ibsheet/',
    baseUrl: '/ibsheet',
    locales: ['ko', 'en'],
    plugins: [
      'excel',
      'common'
    ]
  }]
});

function App() {
  return (
    <div className="App">
      <BrowserRouter className="main">
        <div className="top-nav">
          <span className="logo-text">IBSheet 8</span>
          <ul>
            <li><Link to="/">Home</Link></li>
            <li><Link to="/sample-sheet">Sample</Link></li>
            <li><Link to="/fast-load">Fast load</Link></li>
            <li><Link to="/lazy-load">Lazy load</Link></li>
          </ul>
        </div>
        <div className="content">
          <Route exact path="/" component={Home} />
          <Route exact path="/sample-sheet" component={SheetSample} />
          <Route exact path="/fast-load" component={FastLoadSheet} />
          <Route exact path="/lazy-load" component={LazyLoadSheet} />
        </div>
      </BrowserRouter>
    </div>
  );
}

export default App;
