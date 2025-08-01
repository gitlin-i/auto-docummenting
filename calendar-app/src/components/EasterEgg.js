import React, { useState } from 'react';
import './EasterEgg.css';

const EasterEgg = () => {
  const [showTooltip, setShowTooltip] = useState(false);

  return (
    <div 
      className="easter-egg"
      onMouseEnter={() => setShowTooltip(true)}
      onMouseLeave={() => setShowTooltip(false)}
    >
      <span className="easter-egg-emoji">😊</span>
      {showTooltip && (
        <div className="easter-egg-tooltip">
          만든이: 박석진
        </div>
      )}
    </div>
  );
};

export default EasterEgg; 