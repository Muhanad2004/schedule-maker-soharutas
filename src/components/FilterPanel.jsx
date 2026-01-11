import { useState } from 'react';

export default function FilterPanel({ blockedSlots, onToggleSlot, onToggleDay, onReset, t }) {
  const [expanded, setExpanded] = useState(false);
  const days = t?.days || ['Sun', 'Mon', 'Tue', 'Wed', 'Thu'];
  const hours = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  const isBlocked = (day, hour) => blockedSlots.includes(`${day}-${hour}`);

  return (
    <div className="filter-panel">
      <button className="filter-toggle" onClick={() => setExpanded(!expanded)}>
        <span>🕐 {t.filterTitle}</span>
        <span>{expanded ? '▼' : '▲'}</span>
      </button>

      {expanded && (
        <div className="filter-content">
          <div className="filter-header">
            <span className="filter-hint">{t.filterHint}</span>
            <button className="reset-btn" onClick={onReset}>
              {t.resetFilters}
            </button>
          </div>

          <div className="time-grid">
            <div className="grid-corner"></div>
            {days.map((day, idx) => (
              <div
                key={idx}
                className="grid-header clickable"
                onClick={() => onToggleDay(idx)}
                title={`Block/unblock all ${day}`}
              >
                {day}
              </div>
            ))}

            {hours.map(hour => (
              <>
                <div key={`label-${hour}`} className="grid-time-label">
                  {hour % 12 || 12} {hour < 12 ? 'AM' : 'PM'}
                </div>
                {days.map((_, dayIdx) => (
                  <div
                    key={`${dayIdx}-${hour}`}
                    className={`grid-cell ${isBlocked(dayIdx, hour) ? 'blocked' : ''}`}
                    onClick={() => onToggleSlot(dayIdx, hour)}
                  >
                    {isBlocked(dayIdx, hour) && '✕'}
                  </div>
                ))}
              </>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
