import React, { useState, useEffect } from 'react';
import './Calendar.css';

const Calendar = ({ selectedManager, managers, onDataUpdate, savedData }) => {
  const [currentDate, setCurrentDate] = useState(new Date());
  const [overtimeSchedules, setOvertimeSchedules] = useState(savedData?.overtimeSchedules || {}); // 토요일 추가근무 스케줄
  const [substituteHolidays, setSubstituteHolidays] = useState(savedData?.substituteHolidays || {}); // 대체 휴무일
  const [vacationSchedules, setVacationSchedules] = useState(savedData?.vacationSchedules || {}); // 연가 스케줄
  const [selectedLegend, setSelectedLegend] = useState('overtime'); // 선택된 범례
  const [targetDate, setTargetDate] = useState(savedData?.targetDate || `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`);

  // savedData가 변경될 때 상태 업데이트
  useEffect(() => {
    if (savedData) {
      setOvertimeSchedules(savedData.overtimeSchedules || {});
      setSubstituteHolidays(savedData.substituteHolidays || {});
      setVacationSchedules(savedData.vacationSchedules || {});
    }
  }, [savedData]);

  // 데이터 변경 시 부모 컴포넌트에 알림
  useEffect(() => {
    if (onDataUpdate) {
      console.log('Calendar에서 데이터 업데이트 호출:', {
        overtimeSchedules,
        substituteHolidays,
        vacationSchedules,
        targetDate
      });
      onDataUpdate(overtimeSchedules, substituteHolidays, vacationSchedules, targetDate);
    }
  }, [overtimeSchedules, substituteHolidays, vacationSchedules, targetDate, onDataUpdate]);

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

  const koreanMonthNames = [
    '1월', '2월', '3월', '4월', '5월', '6월',
    '7월', '8월', '9월', '10월', '11월', '12월'
  ];

  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  const prevMonth = () => {
    const newDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
    setCurrentDate(newDate);
    updateTargetDate(newDate);
  };

  const nextMonth = () => {
    const newDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
    setCurrentDate(newDate);
    updateTargetDate(newDate);
  };

  const updateTargetDate = (date) => {
    const newTargetDate = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    setTargetDate(newTargetDate);
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

  // 편의 기능: 현재 달 데이터 클리어
  const clearCurrentMonth = () => {
    const currentMonthKey = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
    
    setOvertimeSchedules(prev => {
      const newSchedules = { ...prev };
      Object.keys(newSchedules).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          delete newSchedules[key];
        }
      });
      return newSchedules;
    });

    setSubstituteHolidays(prev => {
      const newHolidays = { ...prev };
      Object.keys(newHolidays).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          delete newHolidays[key];
        }
      });
      return newHolidays;
    });

    setVacationSchedules(prev => {
      const newVacations = { ...prev };
      Object.keys(newVacations).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          delete newVacations[key];
        }
      });
      return newVacations;
    });
    
    // 파일 저장 알림
    if (onDataUpdate) {
      onDataUpdate(overtimeSchedules, substituteHolidays, vacationSchedules);
    }
  };

  // 편의 기능: 다음 달로 이동 (데이터 4주 뒤로)
  const moveToNextMonth = () => {
    const nextMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 1);
    const currentMonthKey = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
    const nextMonthKey = `${nextMonth.getFullYear()}-${String(nextMonth.getMonth() + 1).padStart(2, '0')}`;

    // 토요일 근무와 대체 휴무일을 4주 뒤로 이동
    setOvertimeSchedules(prev => {
      const newSchedules = { ...prev };
      Object.keys(newSchedules).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          const day = key.split('-')[2];
          const currentDate = new Date(key.split('-')[0], key.split('-')[1] - 1, day);
          const nextDate = new Date(currentDate.getTime() + (4 * 7 * 24 * 60 * 60 * 1000)); // 4주 뒤
          const newKey = `${nextDate.getFullYear()}-${String(nextDate.getMonth() + 1).padStart(2, '0')}-${String(nextDate.getDate()).padStart(2, '0')}`;
          newSchedules[newKey] = newSchedules[key];
          delete newSchedules[key];
        }
      });
      return newSchedules;
    });

    setSubstituteHolidays(prev => {
      const newHolidays = { ...prev };
      Object.keys(newHolidays).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          const day = key.split('-')[2];
          const currentDate = new Date(key.split('-')[0], key.split('-')[1] - 1, day);
          const nextDate = new Date(currentDate.getTime() + (4 * 7 * 24 * 60 * 60 * 1000)); // 4주 뒤
          const newKey = `${nextDate.getFullYear()}-${String(nextDate.getMonth() + 1).padStart(2, '0')}-${String(nextDate.getDate()).padStart(2, '0')}`;
          newHolidays[newKey] = newHolidays[key];
          delete newHolidays[key];
        }
      });
      return newHolidays;
    });

    // 연가는 클리어
    setVacationSchedules(prev => {
      const newVacations = { ...prev };
      Object.keys(newVacations).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          delete newVacations[key];
        }
      });
      return newVacations;
    });

    setCurrentDate(nextMonth);
    
    // 파일 저장 알림
    if (onDataUpdate) {
      onDataUpdate(overtimeSchedules, substituteHolidays, vacationSchedules);
    }
  };

  // 편의 기능: 이전 달로 이동 (데이터 4주 앞으로)
  const moveToPrevMonth = () => {
    const prevMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
    const currentMonthKey = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
    const prevMonthKey = `${prevMonth.getFullYear()}-${String(prevMonth.getMonth() + 1).padStart(2, '0')}`;

    // 토요일 근무와 대체 휴무일을 4주 앞으로 이동
    setOvertimeSchedules(prev => {
      const newSchedules = { ...prev };
      Object.keys(newSchedules).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          const day = key.split('-')[2];
          const currentDate = new Date(key.split('-')[0], key.split('-')[1] - 1, day);
          const prevDate = new Date(currentDate.getTime() - (4 * 7 * 24 * 60 * 60 * 1000)); // 4주 앞
          const newKey = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, '0')}-${String(prevDate.getDate()).padStart(2, '0')}`;
          newSchedules[newKey] = newSchedules[key];
          delete newSchedules[key];
        }
      });
      return newSchedules;
    });

    setSubstituteHolidays(prev => {
      const newHolidays = { ...prev };
      Object.keys(newHolidays).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          const day = key.split('-')[2];
          const currentDate = new Date(key.split('-')[0], key.split('-')[1] - 1, day);
          const prevDate = new Date(currentDate.getTime() - (4 * 7 * 24 * 60 * 60 * 1000)); // 4주 앞
          const newKey = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, '0')}-${String(prevDate.getDate()).padStart(2, '0')}`;
          newHolidays[newKey] = newHolidays[key];
          delete newHolidays[key];
        }
      });
      return newHolidays;
    });

    // 연가는 클리어
    setVacationSchedules(prev => {
      const newVacations = { ...prev };
      Object.keys(newVacations).forEach(key => {
        if (key.startsWith(currentMonthKey)) {
          delete newVacations[key];
        }
      });
      return newVacations;
    });

    setCurrentDate(prevMonth);
    
    // 파일 저장 알림
    if (onDataUpdate) {
      onDataUpdate(overtimeSchedules, substituteHolidays, vacationSchedules);
    }
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
        
        const managerExists = newSchedules[dateKey].find(m => m.name === selectedManager.name);
        
        if (managerExists) {
          newSchedules[dateKey] = newSchedules[dateKey].filter(m => m.name !== selectedManager.name);
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
        
        const managerExists = newHolidays[dateKey].find(m => m.name === selectedManager.name);
        
        if (managerExists) {
          newHolidays[dateKey] = newHolidays[dateKey].filter(m => m.name !== selectedManager.name);
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
        
        const managerExists = newVacations[dateKey].find(m => m.name === selectedManager.name);
        
        if (managerExists) {
          newVacations[dateKey] = newVacations[dateKey].filter(m => m.name !== selectedManager.name);
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
          {currentDate.getFullYear()}년 {koreanMonthNames[currentDate.getMonth()]} 출근부
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

      {/* 편의 기능 버튼들 */}
      <div className="calendar-utilities">
        <button className="utility-btn clear-month-btn" onClick={clearCurrentMonth}>
          🗑️ 현재 달 클리어
        </button>
        <div className="month-navigation">
          <button className="utility-btn prev-month-btn" onClick={moveToPrevMonth}>
            ⬅️ 4주 앞 이동
          </button>
          <button className="utility-btn next-month-btn" onClick={moveToNextMonth}>
            ➡️ 4주 뒤 이동
          </button>
        </div>
      </div>
    </div>
  );
};

export default Calendar; 