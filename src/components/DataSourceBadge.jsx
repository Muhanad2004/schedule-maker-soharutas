import { useLanguage } from './LanguageContext';

export default function DataSourceBadge() {
    const { t } = useLanguage();
    return (
        <div className="data-badge">
            <span className="badge-icon">📊</span>
            <span className="badge-text">{t.dataUpdate}</span>
        </div>
    );
}
