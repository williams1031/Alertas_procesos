import type { Config } from "tailwindcss";

const config: Config = {
  content: [
    "./app/**/*.{js,ts,jsx,tsx,mdx}",
    "./components/**/*.{js,ts,jsx,tsx,mdx}"
  ],
  theme: {
    extend: {
      colors: {
        brand: {
          50: "#f3fbfa",
          100: "#d8f4ef",
          200: "#afe7dc",
          300: "#7ad4ca",
          400: "#42b6af",
          500: "#239790",
          600: "#1b7a75",
          700: "#195f5c",
          800: "#194b49",
          900: "#173f3d"
        },
        ink: "#132026",
        soft: "#f7faf9"
      },
      fontFamily: {
        sans: ["Poppins", "ui-sans-serif", "system-ui"]
      },
      boxShadow: {
        glow: "0 20px 50px -12px rgba(35,151,144,0.35)"
      }
    }
  },
  plugins: []
};

export default config;
