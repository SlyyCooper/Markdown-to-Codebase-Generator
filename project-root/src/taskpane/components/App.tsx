import * as React from "react";
import ChatInterface from "./ChatInterface";
import SplashScreen from "./SplashScreen";
import Header from "./Header";
import SettingsPanel from "./SettingsPanel";

interface AppProps {
  title: string;
}

const App: React.FC<AppProps> = ({ title }) => {
  const [isLoading, setIsLoading] = React.useState(false);
  const [showSettings, setShowSettings] = React.useState(false);
  const [apiKey, setApiKey] = React.useState("YOUR_API_KEY");
  const [splashOpacity, setSplashOpacity] = React.useState(1);

  React.useEffect(() => {
    setTimeout(() => setSplashOpacity(0), 2000);
    setTimeout(() => setSplashOpacity(-1), 3000);
  }, []);

  const handleSettingsClick = () => setShowSettings(true);
  const handleCloseSettings = () => setShowSettings(false);
  const handleApiKeyChange = (key: string) => setApiKey(key);

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      {splashOpacity >= 0 && <SplashScreen opacity={splashOpacity} />}
      <Header logo="assets/logo-filled.png" title={title} onSettingsClick={handleSettingsClick} />
      <ChatInterface apiKey={apiKey} setIsLoading={setIsLoading} />
      {showSettings && <SettingsPanel onClose={handleCloseSettings} apiKey={apiKey} onApiKeyChange={handleApiKeyChange} />}
    </div>
  );
};

export default App;
