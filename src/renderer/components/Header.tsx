import { useEffect, useState } from "react";

const Header = () => {
  const [darkMode, setDarkMode] = useState(false);

  useEffect(() => {
    const saved = localStorage.getItem("theme");
    if (saved === "dark") {
      document.documentElement.classList.add("dark");
      setDarkMode(true);
    }
  }, []);

  const toggleTheme = () => {
    const isDark = document.documentElement.classList.contains("dark");

    if (isDark) {
      document.documentElement.classList.remove("dark");
      localStorage.setItem("theme", "light");
      setDarkMode(false);
    } else {
      document.documentElement.classList.add("dark");
      localStorage.setItem("theme", "dark");
      setDarkMode(true);
    }
  };

  return (
    <header className="h-12 flex items-center justify-end px-4 gap-4 bg-white dark:bg-gray-900 border-b border-gray-200 dark:border-gray-700">
      <button
        onClick={toggleTheme}
        className="px-3 py-1 rounded bg-indigo-600 hover:bg-indigo-700 text-white transition"
      >
        Toggle Theme
      </button>
    </header>
  );
};

export default Header;