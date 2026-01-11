import { useLanguage } from './LanguageContext';

export default function LanguageToggle() {
    const { lang, toggleLang } = useLanguage();

    return (
        <button className="icon-btn" onClick={toggleLang} title="Switch Language">
            {lang === 'en' ? 'عربي' : 'EN'}
        </button>
    );
}
