import Home from "./pages/Home";
import { ThemeProvider } from "./contexts/ThemeContext";

function App() {
  return (
    <ThemeProvider defaultTheme="dark" switchable>
      <Home />
    </ThemeProvider>
  );
}

export default App;
