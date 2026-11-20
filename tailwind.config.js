/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./public/index.html', './public/js/**/*.js'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'system-ui', '-apple-system', 'Segoe UI', 'Roboto', 'sans-serif'],
      },
      colors: {
        accent: {
          50: '#eef2ff',
          200: '#c7d2fe',
          500: '#4f46e5',
          600: '#4338ca',
          700: '#3730a3',
        },
      },
    },
  },
  plugins: [],
};
