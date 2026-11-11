interface Window {
  OmniaTheme: {
    getSettings(): { palette: "teal" | "blue" | "gold" | "red"; mode: "auto" | "dark" | "light" };
    setSettings(value: { palette: string; mode: string }): void;
    refresh(): void;
    observeOffice(): void;
  };
}
