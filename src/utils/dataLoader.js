/**
 * Data loader - fetches course data
 */

export async function loadCourses() {
    const baseUrl = import.meta.env.BASE_URL || '/';
    const url = `${baseUrl}data.json`.replace('//', '/');

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`Failed to load courses: ${response.status}`);
    }

    return response.json();
}
