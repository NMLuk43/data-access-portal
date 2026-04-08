/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./index.html", "./src/**/*.{css,js}"],
  theme: {
    extend: {
      colors: {
        primary: "#003865",
        muted: "#4d5462",
        borderSoft: "#e7e8ea",
        cardBorder: "#e0e0e0"
      },
      borderRadius: {
        card: "0.375rem",
        pill: "0.8rem"
      },
      boxShadow: {
        card: "0 1px 2px rgba(0, 0, 0, 0.07)",
        hoverSoft: "0 4px 8px rgba(0, 0, 0, 0.15)"
      }
    }
  },
  plugins: []
};
