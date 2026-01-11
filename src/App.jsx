import { useState, useEffect, useCallback } from 'react';
import { loadCourses } from './utils/dataLoader';
import { generateSchedules } from './utils/scheduler';
import { parseTimeSlots } from './utils/timeUtils';
import { ThemeProvider } from './components/ThemeContext';
import { LanguageProvider, useLanguage } from './components/LanguageContext';
import ThemeToggle from './components/ThemeToggle';
import LanguageToggle from './components/LanguageToggle';
import LoadingScreen from './components/LoadingScreen';
import CourseSelector from './components/CourseSelector';
import FilterPanel from './components/FilterPanel';
import ScheduleViewer from './components/ScheduleViewer';
import DataSourceBadge from './components/DataSourceBadge';

// LocalStorage helpers
const loadState = (key, fallback) => {
  try {
    const saved = localStorage.getItem(key);
    return saved ? JSON.parse(saved) : fallback;
  } catch {
    return fallback;
  }
};

const saveState = (key, value) => {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore
  }
};

function AppContent() {
  const { t } = useLanguage();

  // Data
  const [allCourses, setAllCourses] = useState([]);
  const [loading, setLoading] = useState(true);

  // User selections
  const [selectedCourses, setSelectedCourses] = useState(() => loadState('utas_selectedCourses_v1', []));
  const [instructorFilters, setInstructorFilters] = useState(() => loadState('utas_instructorFilters_v1', {}));
  const [blockedSlots, setBlockedSlots] = useState(() => loadState('utas_blockedSlots_v1', []));

  // Generated schedules
  const [schedules, setSchedules] = useState([]);
  const [scheduleIndex, setScheduleIndex] = useState(0);
  const [generating, setGenerating] = useState(false);

  // Load data on mount
  useEffect(() => {
    let mounted = true;

    async function load() {
      try {
        const data = await loadCourses();
        if (mounted) {
          // Optimization: Pre-parse time slots to avoid repeated regex operations
          data.forEach(course => {
            course.sections.forEach(section => {
              section.parsedSlots = parseTimeSlots(section.time || '');
            });
          });
          setAllCourses(data);
          setLoading(false);
        }
      } catch (err) {
        console.error('Failed to load courses:', err);
        if (mounted) setLoading(false);
      }
    }

    load();
    return () => { mounted = false; };
  }, []);

  // Persist state
  useEffect(() => { saveState('utas_selectedCourses_v1', selectedCourses); }, [selectedCourses]);
  useEffect(() => { saveState('utas_instructorFilters_v1', instructorFilters); }, [instructorFilters]);
  useEffect(() => { saveState('utas_blockedSlots_v1', blockedSlots); }, [blockedSlots]);

  // Handlers
  const handleToggleCourse = useCallback((course) => {
    setSelectedCourses(prev => {
      const exists = prev.find(c => c.id === course.id);
      return exists ? prev.filter(c => c.id !== course.id) : [...prev, course];
    });
    setSchedules([]);
    setScheduleIndex(0);
  }, []);

  const handleToggleInstructor = useCallback((courseId, instructor) => {
    setInstructorFilters(prev => {
      const course = selectedCourses.find(c => c.id === courseId);
      if (!course) return prev;

      const allInstructors = [...new Set(course.sections.map(s => s.instructor))];
      const current = prev[courseId];

      let newAllowed;
      if (!current) {
        newAllowed = allInstructors.filter(i => i !== instructor);
      } else {
        newAllowed = current.includes(instructor)
          ? current.filter(i => i !== instructor)
          : [...current, instructor];
      }

      if (newAllowed.length === allInstructors.length) {
        const copy = { ...prev };
        delete copy[courseId];
        return copy;
      }
      return { ...prev, [courseId]: newAllowed };
    });
  }, [selectedCourses]);

  const handleToggleDay = useCallback((day) => {
    const hours = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17];
    setBlockedSlots(prev => {
      const allBlocked = hours.every(h => prev.includes(`${day}-${h}`));
      if (allBlocked) {
        return prev.filter(key => !key.startsWith(`${day}-`));
      } else {
        const newSlots = [...prev];
        hours.forEach(h => {
          const key = `${day}-${h}`;
          if (!newSlots.includes(key)) newSlots.push(key);
        });
        return newSlots;
      }
    });
  }, []);

  const handleToggleSlot = useCallback((day, hour) => {
    const key = `${day}-${hour}`;
    setBlockedSlots(prev =>
      prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]
    );
  }, []);

  const handleGenerate = useCallback(() => {
    if (selectedCourses.length === 0) return;

    setGenerating(true);

    setTimeout(() => {
      const filtered = selectedCourses.map(course => {
        const allowed = instructorFilters[course.id];
        const sections = course.sections.filter(section => {
          if (allowed && !allowed.includes(section.instructor)) return false;

          if (allowed && !allowed.includes(section.instructor)) return false;

          const slots = section.parsedSlots || parseTimeSlots(section.time);
          const hasConflict = slots.some(slot => {
            const startH = Math.floor(slot.start / 60);
            const endH = Math.floor((slot.end - 1) / 60);
            for (let h = startH; h <= endH; h++) {
              if (blockedSlots.includes(`${slot.day}-${h}`)) return true;
            }
            return false;
          });

          return !hasConflict;
        });

        return { ...course, sections };
      });

      const results = generateSchedules(filtered);
      setSchedules(results);
      setScheduleIndex(0);
      setGenerating(false);

      if (results.length === 0) {
        alert(t.noSchedulesFound);
      }
    }, 50);
  }, [selectedCourses, instructorFilters, blockedSlots, t]);

  const handleClearCourses = useCallback(() => {
    if (window.confirm(t.confirmClearCourses)) {
      setSelectedCourses([]);
      setSchedules([]);
      setScheduleIndex(0);
      setInstructorFilters({});
    }
  }, [t]);

  const handleResetFilters = useCallback(() => {
    if (window.confirm(t.confirmResetFilters)) {
      setBlockedSlots([]);
    }
  }, [t]);

  if (loading) {
    return <LoadingScreen message={t.loadingCourses} />;
  }

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-left">
          <header className="mb-8 text-center animate-fade-in-down">
            <h1 className="text-4xl md:text-5xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-violet-600 to-amber-500 dark:from-violet-400 dark:to-amber-300 mb-4 transition-all duration-300">
              UTAS Sohar Schedule Maker <span className="text-4xl">🏛️</span>
            </h1>
            <p className="text-gray-600 dark:text-gray-400 text-lg transition-colors duration-300">
              Plan your semester effortlessly
            </p>
          </header>
          <DataSourceBadge sourceDate="2026-01-11" />
        </div>
        <div className="header-right">
          <LanguageToggle />
          <ThemeToggle />
          <button
            className="primary-btn"
            onClick={handleGenerate}
            disabled={selectedCourses.length === 0 || generating}
          >
            {generating ? t.processing : `${t.generate} (${selectedCourses.length})`}
          </button>
        </div>
      </header>

      <main className="app-main">
        <aside className="sidebar">
          <CourseSelector
            courses={allCourses}
            selectedCourses={selectedCourses}
            onToggleCourse={handleToggleCourse}
            instructorFilters={instructorFilters}
            onToggleInstructor={handleToggleInstructor}
            onClearAll={handleClearCourses}
            t={t}
          />
          <FilterPanel
            blockedSlots={blockedSlots}
            onToggleSlot={handleToggleSlot}
            onToggleDay={handleToggleDay}
            onReset={handleResetFilters}
            t={t}
          />
        </aside>

        <section className="main-content">
          <ScheduleViewer
            schedule={schedules[scheduleIndex]}
            scheduleIndex={scheduleIndex}
            totalSchedules={schedules.length}
            onNext={() => setScheduleIndex(i => Math.min(schedules.length - 1, i + 1))}
            onPrev={() => setScheduleIndex(i => Math.max(0, i - 1))}
            t={t}
          />
        </section>
      </main>
    </div>
  );
}

export default function App() {
  return (
    <LanguageProvider>
      <ThemeProvider>
        <AppContent />
      </ThemeProvider>
    </LanguageProvider>
  );
}
