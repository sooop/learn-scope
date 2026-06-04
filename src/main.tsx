import { createRoot } from 'react-dom/client';
import './styles.css';
import { App } from './components/App';

const root = document.getElementById('root');
if (!root) throw new Error('root element not found');

createRoot(root).render(<App />);
