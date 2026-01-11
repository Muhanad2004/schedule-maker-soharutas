/**
 * Time utilities
 */

import { parseTimeSlots } from './scheduler';

/**
 * Format minutes to "HH:MM AM/PM"
 */
export function formatTime(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  const period = h >= 12 ? 'PM' : 'AM';
  const displayH = h % 12 || 12;
  return `${displayH}:${m.toString().padStart(2, '0')} ${period}`;
}

/**
 * Get summary of instructor's teaching times
 */
export function getInstructorScheduleSummary(sections) {
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const daySlots = {};

  for (const section of sections) {
    const slots = parseTimeSlots(section.time);
    for (const slot of slots) {
      const dayName = dayNames[slot.day];
      if (!daySlots[dayName]) daySlots[dayName] = [];

      const timeStr = `${formatTime(slot.start)}-${formatTime(slot.end)}`;
      if (!daySlots[dayName].includes(timeStr)) {
        daySlots[dayName].push(timeStr);
      }
    }
  }

  return Object.entries(daySlots)
    .map(([day, times]) => `${day}: ${times.join(', ')}`)
    .join(' | ');
}

// Re-export for convenience
export { parseTimeSlots };
