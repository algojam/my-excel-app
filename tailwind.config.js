/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./public/**/*.html", // Ito ang nagsasabi sa Tailwind na i-scan ang lahat ng HTML files sa 'public' folder
    "./public/js/**/*.js", // Kung mayroon kang JavaScript files na gumagamit ng Tailwind classes
    // Maaari kang magdagdag pa ng iba pang paths kung saan mo ginagamit ang Tailwind classes
  ],
  theme: {
    extend: {},
  },
  plugins: [],
}