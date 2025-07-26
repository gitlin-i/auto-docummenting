import React, { useState } from 'react';
import './App.css';
import Calendar from './components/Calendar';
import ManagerList from './components/ManagerList';

function App() {
  const [selectedManager, setSelectedManager] = useState(null);
  const [managers, setManagers] = useState([]);

  const handleManagerSelect = (manager) => {
    setSelectedManager(manager);
  };

  const handleManagersUpdate = (newManagers) => {
    setManagers(newManagers);
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>청년이룸 출근부</h1>
      </header>
      <main className="App-main">
        <div className="app-container">
          <ManagerList 
            onManagerSelect={handleManagerSelect}
            selectedManager={selectedManager}
            onManagersUpdate={handleManagersUpdate}
          />
          <Calendar 
            selectedManager={selectedManager}
            managers={managers}
          />
        </div>
      </main>
    </div>
  );
}

export default App;
