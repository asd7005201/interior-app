import { Routes, Route } from 'react-router-dom'
import LandingPage from './pages/LandingPage'
import SurveyPage from './pages/SurveyPage'
import ResultPage from './pages/ResultPage'
import AdminPage from './pages/AdminPage'

function App() {
  return (
    <Routes>
      <Route path="/" element={<LandingPage />} />
      <Route path="/survey" element={<SurveyPage />} />
      <Route path="/result" element={<ResultPage />} />
      <Route path="/admin" element={<AdminPage />} />
    </Routes>
  )
}

export default App
