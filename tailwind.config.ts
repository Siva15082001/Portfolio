import type { Config } from "tailwindcss";

const config: Config = {
  content: [
    "./pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./components/**/*.{js,ts,jsx,tsx,mdx}",
    "./app/**/*.{js,ts,jsx,tsx,mdx}",
    "./data/**/*.{js,ts}",
  ],
  theme: {
    extend: {
      fontFamily: {
        display: ["'Bebas Neue'", "Impact", "sans-serif"],
        mono: ["'DM Mono'", "'Courier New'", "monospace"],
        sans: ["'Instrument Sans'", "system-ui", "sans-serif"],
      },
      colors: {
        bg: "#080808",
        bg2: "#0f0f0f",
        bg3: "#161616",
        border: "#1f1f1f",
        "border-bright": "#2e2e2e",
        muted: "#666666",
        accent: "#FF6200",
      },
    },
  },
  plugins: [],
};
export default config;
