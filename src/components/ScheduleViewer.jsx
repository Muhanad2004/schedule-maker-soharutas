import { useMemo, useRef, useCallback, useState, memo } from 'react';
import { formatTime } from '../utils/timeUtils';

const START_HOUR = 8;
const END_HOUR = 18; // 6 PM
const HOUR_HEIGHT = 90;
const TOTAL_HEIGHT = (END_HOUR - START_HOUR) * HOUR_HEIGHT;

const COLORS = [
  { bg: '#3b82f6', border: '#2563eb' },
  { bg: '#10b981', border: '#059669' },
  { bg: '#f59e0b', border: '#d97706' },
  { bg: '#ef4444', border: '#dc2626' },
  { bg: '#8b5cf6', border: '#7c3aed' },
  { bg: '#ec4899', border: '#db2777' },
  { bg: '#06b6d4', border: '#0891b2' },
  { bg: '#84cc16', border: '#65a30d' },
];

const formatHour = (h) => `${h % 12 || 12} ${h < 12 ? 'AM' : 'PM'}`;

const ScheduleViewer = memo(function ScheduleViewer({
  schedule,
  scheduleIndex,
  totalSchedules,
  onNext,
  onPrev,
  t
}) {
  const scheduleRef = useRef(null);
  const [examsExpanded, setExamsExpanded] = useState(false);
  // Ensure we always have days, even if translation fails
  const days = (t && t.days && t.days.length === 5) ? t.days : ['Sun', 'Mon', 'Tue', 'Wed', 'Thu'];
  const hours = Array.from({ length: END_HOUR - START_HOUR }, (_, i) => START_HOUR + i);

  const handleDownload = useCallback(async () => {
    if (!scheduleRef.current) return;

    // Add export class for clean B&W look
    scheduleRef.current.classList.add('export-mode');

    try {
      const html2canvas = (await import('html2canvas')).default;
      const canvas = await html2canvas(scheduleRef.current, {
        backgroundColor: '#ffffff',
        scale: 2 // Improve quality
      });
      const link = document.createElement('a');
      link.download = `schedule-${scheduleIndex + 1}.png`;
      link.href = canvas.toDataURL();
      link.click();
    } catch (error) {
      console.error('Download failed:', error);
    } finally {
      // Remove export class
      scheduleRef.current.classList.remove('export-mode');
    }
  }, [scheduleIndex]);

  // Process all blocks from all sections
  const blocksByDay = useMemo(() => {
    if (!schedule) return {};

    const result = { 0: [], 1: [], 2: [], 3: [], 4: [] };

    schedule.forEach((section, sectionIdx) => {
      const slots = section.parsedSlots || [];
      slots.forEach(slot => {
        if (slot.day >= 0 && slot.day < 5) {
          result[slot.day].push({
            ...section,
            ...slot,
            colorIndex: sectionIdx
          });
        }
      });
    });

    return result;
  }, [schedule]);

  // Calculate position for a block
  const getBlockStyle = (block) => {
    const top = ((block.start - START_HOUR * 60) / 60) * HOUR_HEIGHT;
    const height = ((block.end - block.start) / 60) * HOUR_HEIGHT;
    const color = COLORS[block.colorIndex % COLORS.length];

    return {
      top: `${top}px`,
      height: `${Math.max(height, 30)}px`,
      backgroundColor: color.bg,
      borderLeftColor: color.border
    };
  };

  if (!schedule) {
    return (
      <div className="schedule-viewer">
        <div className="schedule-empty">
          <span className="empty-icon">📅</span>
          <h3>{t.schedule}</h3>
          <p>{t.noSchedules}</p>
        </div>
      </div>
    );
  }

  // Determine direction for arrows
  const isRtl = document.body.classList.contains('rtl');
  const prevIcon = isRtl ? '→' : '←';
  const nextIcon = isRtl ? '←' : '→';

  return (
    <div className="schedule-viewer">
      {/* Header */}
      <div className="schedule-header">
        <h2>
          {t.schedule} {scheduleIndex + 1}
          <span className="text-dim"> {t.of} {totalSchedules}</span>
        </h2>
        <div className="schedule-controls">
          <button className="nav-btn" onClick={onPrev} disabled={scheduleIndex === 0}>{prevIcon}</button>
          <span className="schedule-counter">{scheduleIndex + 1} / {totalSchedules}</span>
          <button className="nav-btn" onClick={onNext} disabled={scheduleIndex === totalSchedules - 1}>{nextIcon}</button>
          <button className="icon-btn" onClick={handleDownload} title={t.saveImage}>📷</button>
        </div>
      </div>

      {/* Calendar */}
      <div className="schedule-calendar">
        <div ref={scheduleRef} className="calendar-container">
          {/* Day Headers */}
          <div className="calendar-header-row">
            <div className="time-gutter"></div>
            {days.map(day => (
              <div key={day} className="day-header">{day}</div>
            ))}
          </div>

          {/* Grid Body */}
          <div className="calendar-body">
            {/* Time Gutter */}
            <div className="time-gutter">
              {hours.map(h => (
                <div key={h} className="time-slot" style={{ height: HOUR_HEIGHT }}>
                  {formatHour(h)}
                </div>
              ))}
            </div>

            {/* Day Columns */}
            {days.map((_, dayIdx) => (
              <div key={dayIdx} className="day-column" style={{ height: TOTAL_HEIGHT }}>
                {/* Hour grid lines */}
                {hours.map((h, i) => (
                  <div
                    key={h}
                    className="hour-line"
                    style={{ top: i * HOUR_HEIGHT }}
                  />
                ))}

                {/* Course blocks for this day */}
                {(blocksByDay[dayIdx] || []).map((block, i) => (
                  <div
                    key={`${block.code}-${block.section}-${i}`}
                    className="course-block"
                    style={getBlockStyle(block)}
                    title={`${block.code} - ${block.instructor}`}
                  >
                    <div className="block-code">{block.code}/{block.section}</div>
                    <div className="block-time">
                      {formatTime(block.start)} - {formatTime(block.end)}
                    </div>
                    <div className="block-instructor">{block.instructor}</div>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Exam Footer */}
      <div className="schedule-footer">
        <div className="footer-toggle" onClick={() => setExamsExpanded(!examsExpanded)}>
          <span>📋 {t.examFooter}</span>
          <span>{examsExpanded ? '▼' : '▲'}</span>
        </div>

        {examsExpanded && (
          <div className="exam-list">
            {schedule.map((section, idx) => (
              <div
                key={idx}
                className="exam-item"
                style={{ borderLeftColor: COLORS[idx % COLORS.length].bg }}
              >
                <span className="exam-code">{section.code}</span>
                <span className="exam-date">{section.exam || 'TBA'}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
});

export default ScheduleViewer;
