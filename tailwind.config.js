/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./app/**/*.{js,ts,jsx,tsx}",
    "./components/**/*.{js,ts,jsx,tsx}",
    "./pages/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        primary: {
          50: '#edf8ff',
          100: '#d7edff',
          200: '#b0dcff',
          300: '#7cc5ff',
          400: '#3f9fff',
          500: '#127dff',
          600: '#005fed',
          700: '#0049c0',
          800: '#003d99',
          900: '#00337a'
        }
      }
    }
  },
  plugins: [],
};
