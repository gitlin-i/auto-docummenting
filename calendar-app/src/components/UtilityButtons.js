import React from 'react';
import './UtilityButtons.css';

const UtilityButtons = ({ 
  onClearMonth, 
  onNextMonth, 
  onPrevMonth,
  currentDate 
}) => {
  const handleClearMonth = () => {
    if (window.confirm('현재 달의 모든 데이터를 정리하시겠습니까?')) {
      onClearMonth();
    }
  };

  const handleNextMonth = () => {
    if (window.confirm('다음 달로 이동하고 데이터를 한 달 뒤로 이동하시겠습니까?')) {
      onNextMonth();
    }
  };

  const handlePrevMonth = () => {
    if (window.confirm('이전 달로 이동하고 데이터를 한 달 앞으로 이동하시겠습니까?')) {
      onPrevMonth();
    }
  };

  const getMonthYearText = () => {
    const monthNames = [
      '1월', '2월', '3월', '4월', '5월', '6월',
      '7월', '8월', '9월', '10월', '11월', '12월'
    ];
    return `${currentDate.getFullYear()}년 ${monthNames[currentDate.getMonth()]}월`;
  };

  return (
    <div className="utility-buttons-container">
      <h2 className="utility-title">편의 기능</h2>
      
      <div className="utility-buttons">
        <button
          className="utility-btn clear-btn"
          onClick={handleClearMonth}
          title="현재 달의 모든 데이터 정리"
        >
          <div className="button-icon">🗑️</div>
          <span className="button-text">클리어</span>
        </button>

        <button
          className="utility-btn prev-month-btn"
          onClick={handlePrevMonth}
          title="이전 달로 이동 (데이터 한 달 앞으로)"
        >
          <div className="button-icon">⬅️</div>
          <span className="button-text">이전달 이동</span>
        </button>

        <button
          className="utility-btn next-month-btn"
          onClick={handleNextMonth}
          title="다음 달로 이동 (데이터 한 달 뒤로)"
        >
          <div className="button-icon">➡️</div>
          <span className="button-text">다음달 이동</span>
        </button>
      </div>

      <div className="current-month-info">
        현재: {getMonthYearText()}
      </div>
    </div>
  );
};

export default UtilityButtons; 