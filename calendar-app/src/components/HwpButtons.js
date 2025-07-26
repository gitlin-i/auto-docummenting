import React from 'react';
import './HwpButtons.css';

const HwpButtons = () => {
  const handleSinglePrint = () => {
    console.log('한글파일 단일 출력 버튼 클릭됨');
    // TODO: 한글파일 단일 출력 로직 구현
  };

  const handlePrintAll = () => {
    console.log('모두 출력 버튼 클릭됨');
    // TODO: 모두 출력 로직 구현
  };

  return (
    <div className="hwp-buttons-container">
      <h2 className="hwp-title">한글 파일 관리</h2>
      
      <div className="hwp-buttons">
        <button 
          className="hwp-button single-print-button"
          onClick={handleSinglePrint}
        >
          <div className="button-icon">📄</div>
          <span className="button-text">한글파일 단일 출력</span>
        </button>
        
        <button 
          className="hwp-button print-all-button"
          onClick={handlePrintAll}
        >
          <div className="button-icon">🖨️</div>
          <span className="button-text">모두 출력</span>
        </button>
      </div>
    </div>
  );
};

export default HwpButtons; 