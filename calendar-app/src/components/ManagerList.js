import React, { useState, useEffect, useCallback, useRef } from 'react';
import './ManagerList.css';

const ManagerList = ({ onManagerSelect, selectedManager, onManagersUpdate, savedManagers = [] }) => {
  const [managers, setManagers] = useState(savedManagers);
  const [showAddForm, setShowAddForm] = useState(false);
  const [newManagerName, setNewManagerName] = useState('');
  const isInitialMount = useRef(true);
  const lastSavedManagers = useRef(savedManagers);

  const colorPalette = [
    '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4',
    '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F'
  ];

  // savedManagers가 변경될 때 상태 업데이트 (초기 마운트 시에는 제외)
  useEffect(() => {
    if (!isInitialMount.current && JSON.stringify(savedManagers) !== JSON.stringify(lastSavedManagers.current)) {
      setManagers(savedManagers);
      lastSavedManagers.current = savedManagers;
    } else if (isInitialMount.current) {
      isInitialMount.current = false;
      lastSavedManagers.current = savedManagers;
    }
  }, [savedManagers]);

  // 매니저 목록이 변경될 때마다 부모 컴포넌트에 전달 (실제 변경이 있을 때만)
  const handleManagersUpdate = useCallback((newManagers) => {
    if (JSON.stringify(newManagers) !== JSON.stringify(lastSavedManagers.current)) {
      lastSavedManagers.current = newManagers;
      if (onManagersUpdate) {
        onManagersUpdate(newManagers);
      }
    }
  }, [onManagersUpdate]);

  useEffect(() => {
    if (!isInitialMount.current) {
      handleManagersUpdate(managers);
    }
  }, [managers, handleManagersUpdate]);

  const addManager = () => {
    if (managers.length >= 4) {
      alert('매니저는 최대 4명까지 추가할 수 있습니다.');
      return;
    }
    setShowAddForm(true);
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    if (newManagerName.trim()) {
      const newManager = {
        id: Date.now(),
        name: newManagerName.trim(),
        color: colorPalette[managers.length % colorPalette.length]
      };
      setManagers([...managers, newManager]);
      setNewManagerName('');
      setShowAddForm(false);
    }
  };

  const removeManager = (id) => {
    setManagers(managers.filter(manager => manager.id !== id));
    // 선택된 매니저가 삭제되면 선택 해제
    if (selectedManager && selectedManager.id === id) {
      onManagerSelect(null);
    }
  };

  const handleManagerClick = (manager) => {
    onManagerSelect(manager);
  };

  return (
    <div className="manager-list-container">
      <h2 className="manager-title">매니저 목록</h2>
      
      <div className="managers-grid">
        {managers.map((manager) => (
          <div
            key={manager.id}
            className={`manager-card ${selectedManager && selectedManager.id === manager.id ? 'selected' : ''}`}
            style={{ backgroundColor: manager.color }}
            onClick={() => handleManagerClick(manager)}
          >
            <div className="manager-info">
              <span className="manager-name">{manager.name}</span>
            </div>
            <button
              className="remove-manager-btn"
              onClick={(e) => {
                e.stopPropagation();
                removeManager(manager.id);
              }}
            >
              ×
            </button>
          </div>
        ))}
        
        {managers.length < 4 && (
          <div className="add-manager-card" onClick={addManager}>
            <div className="add-icon">+</div>
            <span className="add-text">매니저 추가</span>
          </div>
        )}
      </div>

      {showAddForm && (
        <div className="modal-overlay">
          <div className="modal">
            <h3>새 매니저 추가</h3>
            <form onSubmit={handleSubmit}>
              <input
                type="text"
                value={newManagerName}
                onChange={(e) => setNewManagerName(e.target.value)}
                placeholder="매니저 이름을 입력하세요"
                className="manager-name-input"
                autoFocus
              />
              <div className="modal-buttons">
                <button type="submit" className="submit-btn">
                  추가
                </button>
                <button 
                  type="button" 
                  className="cancel-btn"
                  onClick={() => {
                    setShowAddForm(false);
                    setNewManagerName('');
                  }}
                >
                  취소
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default ManagerList; 