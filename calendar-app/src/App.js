import React, { useState, useEffect } from 'react';
import './App.css';
import Calendar from './components/Calendar';
import ManagerList from './components/ManagerList';
import HwpButtons from './components/HwpButtons';
import FileStorage from './utils/fileStorage';

function App() {
  const [selectedManager, setSelectedManager] = useState(null);
  const [managers, setManagers] = useState([]);
  const [fileStorage] = useState(new FileStorage());
  const [isFileInitialized, setIsFileInitialized] = useState(false);
  const [savedCalendarData, setSavedCalendarData] = useState({
    overtimeSchedules: {},
    substituteHolidays: {},
    vacationSchedules: {}
  });

  // 앱 시작 시 File System Access API 지원 확인
  useEffect(() => {
    if (!('showOpenFilePicker' in window)) {
      alert('이 브라우저는 File System Access API를 지원하지 않습니다. Chrome 86+ 버전을 사용해주세요.');
    }
  }, []);

  const handleManagerSelect = (manager) => {
    setSelectedManager(manager);
  };

  const handleManagersUpdate = (newManagers) => {
    setManagers(newManagers);
    // 파일에 매니저 데이터 저장
    if (isFileInitialized) {
      fileStorage.updateManagers(newManagers);
      fileStorage.saveData();
    }
  };

  const handleCalendarDataUpdate = (overtimeSchedules, substituteHolidays, vacationSchedules) => {
    // 파일에 달력 데이터 저장
    if (isFileInitialized) {
      fileStorage.updateCalendarData(overtimeSchedules, substituteHolidays, vacationSchedules);
      fileStorage.saveData();
    }
  };

  const handleOpenFile = async () => {
    const fileOpened = await fileStorage.initializeFile();
    if (fileOpened) {
      // 파일에서 데이터 로드
      const loadSuccess = await fileStorage.loadData();
      if (loadSuccess) {
        const data = fileStorage.getData();
        setManagers(data.managers || []);
        setSavedCalendarData({
          overtimeSchedules: data.overtimeSchedules || {},
          substituteHolidays: data.substituteHolidays || {},
          vacationSchedules: data.vacationSchedules || {}
        });
        setIsFileInitialized(true);
        alert('파일이 성공적으로 로드되었습니다.');
      } else {
        alert('파일 로드에 실패했습니다.');
      }
    }
  };

  const handleCreateNewFile = async () => {
    const created = await fileStorage.createNewFile();
    if (created) {
      setIsFileInitialized(true);
      // 새 파일 생성 시 기존 데이터 초기화
      setManagers([]);
      setSavedCalendarData({
        overtimeSchedules: {},
        substituteHolidays: {},
        vacationSchedules: {}
      });
      alert('새 파일이 생성되었습니다.');
    }
  };

  const handleSaveFile = async () => {
    if (!isFileInitialized) {
      alert('먼저 파일을 열거나 새 파일을 생성해주세요.');
      return;
    }
    
    const saved = await fileStorage.saveData();
    if (saved) {
      alert('데이터가 성공적으로 저장되었습니다.');
    } else {
      alert('데이터 저장에 실패했습니다.');
    }
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>청년이룸 출근부</h1>
        <div className="file-controls">
          <button 
            className="file-btn"
            onClick={handleOpenFile}
            title="기존 파일 열기"
          >
            📂 열기
          </button>
          <button 
            className="file-btn"
            onClick={handleCreateNewFile}
            title="새 파일 생성"
          >
            📄 새로 만들기
          </button>
          <button 
            className="file-btn"
            onClick={handleSaveFile}
            title="현재 데이터 저장"
            disabled={!isFileInitialized}
          >
            💾 저장
          </button>
        </div>
      </header>
      <main className="App-main">
        <div className="app-container">
          <ManagerList 
            onManagerSelect={handleManagerSelect}
            selectedManager={selectedManager}
            onManagersUpdate={handleManagersUpdate}
            savedManagers={managers}
          />
          <Calendar 
            selectedManager={selectedManager}
            managers={managers}
            onDataUpdate={handleCalendarDataUpdate}
            savedData={savedCalendarData}
          />
          <HwpButtons 
            managers={managers}
            savedData={savedCalendarData}
            currentDate={new Date()}
          />
        </div>
      </main>
    </div>
  );
}

export default App;
