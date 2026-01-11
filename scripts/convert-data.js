import fs from 'fs';
import path from 'path';
import * as XLSX_Module from 'xlsx';
import { fileURLToPath } from 'url';

// Handle ESM/CommonJS interop for XLSX
const XLSX = XLSX_Module.default || XLSX_Module;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const RAW_DATA_DIR = path.join(__dirname, '../raw-data');
const OUTPUT_FILE = path.join(__dirname, '../public/data.json');

function convertData() {
    if (!fs.existsSync(RAW_DATA_DIR)) {
        console.error(`Directory not found: ${RAW_DATA_DIR}`);
        process.exit(1);
        console.log(`Looking for files in: ${RAW_DATA_DIR}`);
        if (fs.existsSync(RAW_DATA_DIR)) {
            console.log('Directory exists.');
        } else {
            console.error('Directory does NOT exist.');
            process.exit(1);
        }

        const files = fs.readdirSync(RAW_DATA_DIR).filter(file => file.endsWith('.xls') || file.endsWith('.xlsx'));

        if (files.length === 0) {
            console.warn('No Excel files found in raw-data directory.');
            return;
        }

        let allCourses = [];

        files.forEach(file => {
            console.log(`Processing ${file}...`);
            const workbook = XLSX.readFile(path.join(RAW_DATA_DIR, file));
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (rows.length === 0) return;

            let currentCourse = null;

            // Iterate through all rows
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];

                // Check if this row is a Course Header row
                // Format: [Code, Section, Name, Credit, Lecturer]
                // Example: ["BSCR3001", "10", "Entrepreneurship", "3", "Naktal Al Kharousi"]
                const code = String(row[0] || '').trim();

                // Regex for course code: 4 letters + 4 digits (e.g., BSCR3001)
                // Or sometimes just generic check if it looks like a code and has a section number next to it
                // Adjust regex as needed based on data.
                const isCourseHeader = /^[A-Z]{4}\d{4}$/.test(code) && row[1];

                if (isCourseHeader) {
                    // If we have a current course building, push it to list before starting new one
                    if (currentCourse) {
                        allCourses.push(currentCourse);
                    }

                    currentCourse = {
                        code: code,
                        section: String(row[1] || '').trim(),
                        name: String(row[2] || '').trim(),
                        instructor: String(row[4] || '').trim(),
                        times: [],
                        rooms: []
                    };
                } else if (currentCourse) {
                    // If we are currently parsing a course, check subsequent rows for Time/Room info
                    // Format: [Day Time, Room]
                    // Example: ["MON 13:00-13:50", "T/T002"]

                    const timeStr = String(row[0] || '').trim();
                    const roomStr = String(row[1] || '').trim();

                    // It is a time row if it contains days or time patterns
                    // Simple check: has numbers and potentially day names
                    // The provided sample showed "MON 13:00-13:50"
                    const hasTime = /\d{2}:\d{2}/.test(timeStr);

                    if (hasTime) {
                        currentCourse.times.push(timeStr);
                        if (roomStr) {
                            currentCourse.rooms.push(roomStr);
                        }
                    } else {
                        // If row is empty or completely unrelated (like "College Requirement" header), 
                        // we might want to stop associating with current course?
                        // However, sometimes there are blank rows between time slots.
                        // For safety, we only "reset" currentCourse if we hit a new course header, which is handled above.
                        // But if we hit a known non-course header line, maybe we should stop.

                        const firstCell = String(row[0] || '').toLowerCase();
                        if (firstCell.includes('course code') || firstCell.includes('college requirement') || (!firstCell && row[1])) {
                            // Looks like a header or sub-header, stop current course context
                            allCourses.push(currentCourse);
                            currentCourse = null;
                        }
                    }
                }
            }
            // Push the last one
            if (currentCourse) {
                allCourses.push(currentCourse);
            }
        });

        const processedData = processCourses(allCourses);
    }
    console.log(`Looking for files in: ${RAW_DATA_DIR}`);
    if (fs.existsSync(RAW_DATA_DIR)) {
        console.log('Directory exists.');
    } else {
        console.error('Directory does NOT exist.');
        process.exit(1);
    }

    const files = fs.readdirSync(RAW_DATA_DIR).filter(file => file.endsWith('.xls') || file.endsWith('.xlsx'));

    if (files.length === 0) {
        console.warn('No Excel files found in raw-data directory.');
        return;
    }

    let allCourses = [];

    files.forEach(file => {
        console.log(`Processing ${file}...`);
        const workbook = XLSX.readFile(path.join(RAW_DATA_DIR, file));
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (rows.length === 0) return;

        let currentCourse = null;

        // Iterate through all rows
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            // Check if this row is a Course Header row
            // Format: [Code, Section, Name, Credit, Lecturer]
            // Example: ["BSCR3001", "10", "Entrepreneurship", "3", "Naktal Al Kharousi"]
            const code = String(row[0] || '').trim();

            // Regex for course code: 4 letters + 4 digits (e.g., BSCR3001)
            // Or sometimes just generic check if it looks like a code and has a section number next to it
            // Adjust regex as needed based on data.
            const isCourseHeader = /^[A-Z]{4}\d{4}$/.test(code) && row[1];

            if (isCourseHeader) {
                // If we have a current course building, push it to list before starting new one
                if (currentCourse) {
                    allCourses.push(currentCourse);
                }

                currentCourse = {
                    code: code,
                    section: String(row[1] || '').trim(),
                    name: String(row[2] || '').trim(),
                    instructor: String(row[4] || '').trim(),
                    times: [],
                    rooms: []
                };
            } else if (currentCourse) {
                // If we are currently parsing a course, check subsequent rows for Time/Room info
                // Format: [Day Time, Room]
                // Example: ["MON 13:00-13:50", "T/T002"]

                const timeStr = String(row[0] || '').trim();
                const roomStr = String(row[1] || '').trim();

                // It is a time row if it contains days or time patterns
                // Simple check: has numbers and potentially day names
                // The provided sample showed "MON 13:00-13:50"
                const hasTime = /\d{2}:\d{2}/.test(timeStr);

                if (hasTime) {
                    currentCourse.times.push(timeStr);
                    if (roomStr) {
                        currentCourse.rooms.push(roomStr);
                    }
                } else {
                    // If row is empty or completely unrelated (like "College Requirement" header), 
                    // we might want to stop associating with current course?
                    // However, sometimes there are blank rows between time slots.
                    // For safety, we only "reset" currentCourse if we hit a new course header, which is handled above.
                    // But if we hit a known non-course header line, maybe we should stop.

                    const firstCell = String(row[0] || '').toLowerCase();
                    if (firstCell.includes('course code') || firstCell.includes('college requirement') || (!firstCell && row[1])) {
                        // Looks like a header or sub-header, stop current course context
                        allCourses.push(currentCourse);
                        currentCourse = null;
                    }
                }
            }
        }
        // Push the last one
        if (currentCourse) {
            allCourses.push(currentCourse);
        }
    });

    const processedData = processCourses(allCourses);

    fs.writeFileSync(OUTPUT_FILE, JSON.stringify(processedData, null, 2));
    console.log(`Data converted successfully! Saved to ${OUTPUT_FILE}`);
}

function processCourses(flatData) {
    const coursesMap = {};

    flatData.forEach(row => {
        const { code, name, section, instructor, times, rooms } = row;

        // Join times and rooms for storage
        const timeStr = times.join(' | ');
        const roomStr = rooms.join(' / ');

        if (!coursesMap[code]) {
            coursesMap[code] = {
                id: code,
                code: code,
                name: name || code,
                sectionsMap: {} // Temporary map to merge sections
            };
        }

        const course = coursesMap[code];

        if (!course.sectionsMap[section]) {
            course.sectionsMap[section] = {
                section,
                instructor,
                time: timeStr,
                room: roomStr,
                exam: '' // No exam info in this format yet
            };
        } else {
            // If section exists (unlikely in this linear parse, but possible if file has duplicates)
            // merge logic if needed, but for now simple overwrite or append
            const existing = course.sectionsMap[section];
            if (timeStr && !existing.time.includes(timeStr)) {
                existing.time += ` | ${timeStr}`;
            }
        }
    });

    // Convert map to array
    return Object.values(coursesMap).map(c => ({
        id: c.id,
        code: c.code,
        name: c.name,
        sections: Object.values(c.sectionsMap)
    }));
}

// Remove the automatic call, make it exportable or wrapped if needed, 
// but since we run it as a script, keep the call.
convertData();
