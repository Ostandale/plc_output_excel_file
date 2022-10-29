import React from 'react';
import logo from './logo.svg';
import './App.css';

import { open } from '@tauri-apps/api/dialog';
import { invoke } from '@tauri-apps/api/tauri';


let error_mes_from_rust;
let is_error_from_rust: any;

function openDialog() {
  open({
    directory: true
  }).then(
    value => invoke('select_csv_dir', { path: value })
      .then(() => {
        is_error_from_rust = false;
      })
      .catch((err) => {
        console.log("error :" + err);
        is_error_from_rust = true;
        error_mes_from_rust = err;
      })
  );
}

function App() {
  return (
    <div className="App">
      <header className="App-header">

        <h1>FETリード修正機のCSVデータを<br></br>エクセル用に解析するプログラム</h1>
        <button onClick={openDialog}>CSVログのディレクトリを指定して下さい</button>
      </header>
    </div>
  );
}

export default App;
