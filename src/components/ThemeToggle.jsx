import { useTheme } from './ThemeContext';

export default function ThemeToggle() {
    const { theme, toggleTheme } = useTheme();

    return (
        <button
            className="icon-btn"
            onClick={toggleTheme}
            title={theme === 'light' ? 'Switch to Dark' : 'Switch to Light'}
        >
            {theme === 'light' ? '🌙' : '☀️'}
        </button>
    );
}
