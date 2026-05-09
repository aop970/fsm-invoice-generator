import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// https://vite.dev/config/
export default defineConfig({
  base: '/fsm-invoice-generator/',
  plugins: [
    tailwindcss(),
    react(),
  ],
  server: {
    proxy: {
      '/generate': 'http://localhost:3001',
    },
  },
})
