import React from 'react';
import './App.css';
import Calendar from './components/Calendar';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>Calendar App</h1>
      </header>
      <main className="App-main">
        <Calendar />
      </main>
    </div>
  );
}

export default App;
