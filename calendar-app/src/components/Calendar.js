import React, { useState } from 'react';
import './Calendar.css';

const Calendar = ({ selectedManager, managers }) => {
  const [currentDate, setCurrentDate] = useState(new Date());
  const [overtimeSchedules, setOvertimeSchedules] = useState({}); // 토요일 추가근무 스케줄
  const [substituteHolidays, setSubstituteHolidays] = useState({}); // 대체 휴무일
  const [vacationSchedules, setVacationSchedules] = useState({}); // 연가 스케줄
  const [selectedLegend, setSelectedLegend] = useState('overtime'); // 선택된 범례

  const daysInMonth = (date) => {
    return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  };

  const firstDayOfMonth = (date) => {
    return new Date(date.getFullYear(), date.getMonth(), 1).getDay();
  };

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  const prevMonth = () => {
    setCurrentDate(new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1));
  };

  const nextMonth = () => {
    setCurrentDate(new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1));
  };

  const isSaturday = (day) => {
    const date = new Date(currentDate.getFullYear(), currentDate.getMonth(), day);
    return date.getDay() === 6;
  };

  const isWeekday = (day) => {
    const date = new Date(currentDate.getFullYear(), currentDate.getMonth(), day);
    const dayOfWeek = date.getDay();
    return dayOfWeek >= 1 && dayOfWeek <= 5; // 월요일(1) ~ 금요일(5)
  };

  const handleLegendClick = (legendType) => {
    setSelectedLegend(legendType);
  };

  const handleDateClick = (day) => {
    if (!selectedManager) {
      alert('매니저를 먼저 선택해주세요.');
      return;
    }

    const dateKey = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    
    if (selectedLegend === 'overtime' && isSaturday(day)) {
      // 토요일 추가근무 처리
      setOvertimeSchedules(prev => {
        const newSchedules = { ...prev };
        if (!newSchedules[dateKey]) {
          newSchedules[dateKey] = [];
        }
        
        const managerExists = newSchedules[dateKey].find(m => m.id === selectedManager.id);
        
        if (managerExists) {
          newSchedules[dateKey] = newSchedules[dateKey].filter(m => m.id !== selectedManager.id);
          if (newSchedules[dateKey].length === 0) {
            delete newSchedules[dateKey];
          }
        } else {
          newSchedules[dateKey] = [...newSchedules[dateKey], selectedManager];
        }
        
        return newSchedules;
      });
    } else if (selectedLegend === 'holiday' && isWeekday(day)) {
      // 평일 대체 휴무일 처리
      setSubstituteHolidays(prev => {
        const newHolidays = { ...prev };
        if (!newHolidays[dateKey]) {
          newHolidays[dateKey] = [];
        }
        
        const managerExists = newHolidays[dateKey].find(m => m.id === selectedManager.id);
        
        if (managerExists) {
          newHolidays[dateKey] = newHolidays[dateKey].filter(m => m.id !== selectedManager.id);
          if (newHolidays[dateKey].length === 0) {
            delete newHolidays[dateKey];
          }
        } else {
          newHolidays[dateKey] = [...newHolidays[dateKey], selectedManager];
        }
        
        return newHolidays;
      });
    } else if (selectedLegend === 'vacation') {
      // 연가 처리
      setVacationSchedules(prev => {
        const newVacations = { ...prev };
        if (!newVacations[dateKey]) {
          newVacations[dateKey] = [];
        }
        
        const managerExists = newVacations[dateKey].find(m => m.id === selectedManager.id);
        
        if (managerExists) {
          newVacations[dateKey] = newVacations[dateKey].filter(m => m.id !== selectedManager.id);
          if (newVacations[dateKey].length === 0) {
            delete newVacations[dateKey];
          }
        } else {
          newVacations[dateKey] = [...newVacations[dateKey], selectedManager];
        }
        
        return newVacations;
      });
    }
  };

  const renderCalendar = () => {
    const days = [];
    const totalDays = daysInMonth(currentDate);
    const firstDay = firstDayOfMonth(currentDate);

    // Add empty cells for days before the first day of the month
    for (let i = 0; i < firstDay; i++) {
      days.push(<div key={`empty-${i}`} className="calendar-day empty"></div>);
    }

    // Add days of the month
    for (let day = 1; day <= totalDays; day++) {
      const isToday = new Date().toDateString() === new Date(currentDate.getFullYear(), currentDate.getMonth(), day).toDateString();
      const dateKey = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const overtimeManagers = overtimeSchedules[dateKey] || [];
      const holidayManagers = substituteHolidays[dateKey] || [];
      const vacationManagers = vacationSchedules[dateKey] || [];
      
      const isSaturdayDay = isSaturday(day);
      const isWeekdayDay = isWeekday(day);
      
      days.push(
        <div
          key={day}
          className={`calendar-day ${isToday ? 'today' : ''} ${isSaturdayDay ? 'saturday' : ''} ${isWeekdayDay ? 'weekday' : ''}`}
          onClick={() => handleDateClick(day)}
        >
          <span className="day-number">{day}</span>
          {overtimeManagers.length > 0 && (
            <div className="manager-indicators">
              {overtimeManagers.map((manager) => (
                <div
                  key={manager.id}
                  className="manager-indicator overtime"
                  style={{ backgroundColor: manager.color }}
                  title={`${manager.name} - 토요일 근무`}
                />
              ))}
            </div>
          )}
          {holidayManagers.length > 0 && (
            <div className="manager-indicators">
              {holidayManagers.map((manager) => (
                <div
                  key={manager.id}
                  className="manager-indicator holiday"
                  style={{ backgroundColor: manager.color }}
                  title={`${manager.name} - 대체 휴무일`}
                />
              ))}
            </div>
          )}
          {vacationManagers.length > 0 && (
            <div className="manager-indicators">
              {vacationManagers.map((manager) => (
                <div
                  key={manager.id}
                  className="manager-indicator vacation"
                  style={{ backgroundColor: manager.color }}
                  title={`${manager.name} - 연가`}
                />
              ))}
            </div>
          )}
        </div>
      );
    }

    return days;
  };

  return (
    <div className="calendar-container">
      <div className="calendar-header">
        <button className="calendar-nav-btn" onClick={prevMonth}>
          ‹
        </button>
        <h2 className="calendar-title">
          {monthNames[currentDate.getMonth()]} {currentDate.getFullYear()}
        </h2>
        <button className="calendar-nav-btn" onClick={nextMonth}>
          ›
        </button>
      </div>
      
      <div className="calendar-legend">
        <div 
          className={`legend-item ${selectedLegend === 'overtime' ? 'selected' : ''}`}
          onClick={() => handleLegendClick('overtime')}
        >
          <div className="legend-indicator overtime"></div>
          <span>토요일 근무</span>
        </div>
        <div 
          className={`legend-item ${selectedLegend === 'holiday' ? 'selected' : ''}`}
          onClick={() => handleLegendClick('holiday')}
        >
          <div className="legend-indicator holiday"></div>
          <span>대체 휴무일</span>
        </div>
        <div 
          className={`legend-item ${selectedLegend === 'vacation' ? 'selected' : ''}`}
          onClick={() => handleLegendClick('vacation')}
        >
          <div className="legend-indicator vacation"></div>
          <span>연가</span>
        </div>
      </div>
      
      <div className="calendar-grid">
        <div className="calendar-weekdays">
          {dayNames.map(day => (
            <div key={day} className="weekday">{day}</div>
          ))}
        </div>
        <div className="calendar-days">
          {renderCalendar()}
        </div>
      </div>
      
      {selectedManager && (
        <div className="selected-manager-info">
          선택된 매니저: <span style={{ color: selectedManager.color, fontWeight: 'bold' }}>{selectedManager.name}</span>
          <br />
          선택된 범례: <span style={{ fontWeight: 'bold' }}>
            {selectedLegend === 'overtime' && '토요일 근무'}
            {selectedLegend === 'holiday' && '대체 휴무일'}
            {selectedLegend === 'vacation' && '연가'}
          </span>
        </div>
      )}
    </div>
  );
};

export default Calendar; 