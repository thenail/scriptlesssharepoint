import { createContext, useContext, useEffect, useState } from "react";
import { webDarkTheme, webLightTheme } from "@fluentui/react-components";

import type { Api } from "./bridge";
import type { AppInfo } from "./types";

export interface AppContextValue {
  api: Api;
  info: AppInfo | null;
  /** Shows a short message in a toast. */
  notify(message: string, intent?: "success" | "info" | "warning" | "error"): void;
}

export const AppContext = createContext<AppContextValue | null>(null);

export function useApp(): AppContextValue {
  const value = useContext(AppContext);
  if (!value) {
    throw new Error("useApp outside AppContext");
  }
  return value;
}

/** Fluent's web theme, light or dark as Windows is. */
export function useSystemTheme() {
  const query = "(prefers-color-scheme: dark)";
  const [dark, setDark] = useState(() => matchMedia(query).matches);
  useEffect(() => {
    const m = matchMedia(query);
    const onChange = () => setDark(m.matches);
    m.addEventListener("change", onChange);
    return () => m.removeEventListener("change", onChange);
  }, []);
  return dark ? webDarkTheme : webLightTheme;
}
