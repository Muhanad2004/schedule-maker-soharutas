import './LoadingScreen.css';

export default function LoadingScreen({ message = 'Loading...' }) {
    return (
        <div className="loading-screen">
            <div className="loading-content">
                <div className="loading-spinner"></div>
                <p className="loading-message">{message}</p>
            </div>
        </div>
    );
}
