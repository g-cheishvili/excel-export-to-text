import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  // Relative base so the build works from any GitHub Pages path
  // (e.g. https://user.github.io/excel-export/) without hardcoding the repo name.
  base: './',
  plugins: [react()],
})
