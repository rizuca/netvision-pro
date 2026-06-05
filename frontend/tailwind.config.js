/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        network: {
          dark: '#0B0F19',
          card: 'rgba(17, 25, 40, 0.75)',
          border: 'rgba(255, 255, 255, 0.125)',
          green: '#10B981',
          red: '#EF4444',
          blue: '#3B82F6',
          purple: '#8B5CF6',
        }
      },
      backdropBlur: {
        xs: '2px',
      }
    },
  },
  plugins: [],
}
