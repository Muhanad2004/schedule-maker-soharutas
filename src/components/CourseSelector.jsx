import { useState, useMemo, useEffect } from 'react';
import { getInstructorScheduleSummary } from '../utils/timeUtils';

export default function CourseSelector({
  courses,
  selectedCourses,
  onToggleCourse,
  instructorFilters,
  onToggleInstructor,
  onClearAll,
  t
}) {
  const [searchTerm, setSearchTerm] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [activeTab, setActiveTab] = useState('all');

  // Debounce search
  useEffect(() => {
    const timer = setTimeout(() => setDebouncedSearch(searchTerm), 200);
    return () => clearTimeout(timer);
  }, [searchTerm]);

  const filteredCourses = useMemo(() => {
    if (!debouncedSearch.trim()) return [];
    const terms = debouncedSearch.toLowerCase().split(/\s+/).filter(Boolean);
    return courses.filter(c => {
      const code = c.code.toLowerCase();
      const name = c.name.toLowerCase();
      return terms.every(term => code.includes(term) || name.includes(term));
    });
  }, [courses, debouncedSearch]);

  const isSelected = (id) => selectedCourses.some(c => c.id === id);

  return (
    <div className="course-selector">
      <div className="selector-header">
        <input
          type="text"
          className="search-input"
          placeholder={t.searchPlaceholder}
          value={searchTerm}
          onChange={e => setSearchTerm(e.target.value)}
        />
      </div>

      <div className="selector-tabs">
        <button
          className={`tab ${activeTab === 'all' ? 'active' : ''}`}
          onClick={() => setActiveTab('all')}
        >
          {t.tabs.all}
        </button>
        <button
          className={`tab ${activeTab === 'selected' ? 'active' : ''}`}
          onClick={() => setActiveTab('selected')}
        >
          {t.tabs.selected} ({selectedCourses.length})
        </button>
      </div>

      <div className="selector-content">
        {activeTab === 'all' ? (
          <SearchResults
            searchTerm={debouncedSearch}
            courses={filteredCourses}
            isSelected={isSelected}
            onToggle={onToggleCourse}
            t={t}
          />
        ) : (
          <SelectedCourses
            courses={selectedCourses}
            instructorFilters={instructorFilters}
            onToggle={onToggleCourse}
            onToggleInstructor={onToggleInstructor}
            onClearAll={onClearAll}
            t={t}
          />
        )}
      </div>
    </div>
  );
}

function SearchResults({ searchTerm, courses, isSelected, onToggle, t }) {
  if (!searchTerm.trim()) {
    return <div className="empty-state">{t.emptySearchPrompt}</div>;
  }
  if (courses.length === 0) {
    return <div className="empty-state">{t.noMatchesFound}</div>;
  }
  return (
    <div className="course-list">
      {courses.map(course => (
        <div
          key={course.id}
          className={`course-item ${isSelected(course.id) ? 'selected' : ''}`}
          onClick={() => onToggle(course)}
        >
          <div className="course-info">
            <div className="course-code">{course.code}</div>
            <div className="course-name">{course.name}</div>
          </div>
          <span className="course-action">{isSelected(course.id) ? '✓' : '+'}</span>
        </div>
      ))}
    </div>
  );
}

function SelectedCourses({ courses, instructorFilters, onToggle, onToggleInstructor, onClearAll, t }) {
  if (courses.length === 0) {
    return <div className="empty-state">{t.noCoursesSelected}</div>;
  }

  return (
    <div className="course-list">
      {courses.length > 0 && (
        <button className="clear-all-btn" onClick={onClearAll}>
          Clear All
        </button>
      )}

      {courses.map(course => {
        const instructorMap = {};
        course.sections.forEach(s => {
          if (!instructorMap[s.instructor]) instructorMap[s.instructor] = [];
          instructorMap[s.instructor].push(s);
        });
        const instructors = Object.keys(instructorMap);
        const allowed = instructorFilters[course.id];

        return (
          <div key={course.id} className="selected-course-card">
            <div className="card-header">
              <div>
                <div className="course-code">{course.code}</div>
                <div className="course-name">{course.name}</div>
              </div>
              <button className="remove-btn" onClick={() => onToggle(course)}>✕</button>
            </div>

            <div className="instructor-filters">
              <span className="filter-label">{t.instructorTimesLabel}:</span>
              <div className="instructor-chips">
                {instructors.map(inst => {
                  const isIncluded = !allowed || allowed.includes(inst);
                  const timeSummary = getInstructorScheduleSummary(instructorMap[inst]);
                  return (
                    <button
                      key={inst}
                      className={`chip ${isIncluded ? 'included' : 'excluded'}`}
                      onClick={() => onToggleInstructor(course.id, inst)}
                      title={timeSummary}
                    >
                      <span className="chip-name">{inst}</span>
                      <span className="chip-time">{timeSummary}</span>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}
