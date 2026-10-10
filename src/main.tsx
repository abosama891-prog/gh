import React from 'react';
import ReactDOM from 'react-dom/client';
import CarsApp from './CarsApp';
import { initializeServiceWorker } from './services/serviceWorker';
import './styles/globals.css';

initializeServiceWorker();

const root = ReactDOM.createRoot(document.getElementById('root') as HTMLElement);
root.render(
  <React.StrictMode>
    <CarsApp />
  </React.StrictMode>
);