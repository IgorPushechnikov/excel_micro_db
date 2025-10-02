import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  // Для Electron указываем базовый путь, если будем использовать относительные пути
  // base: './', 
  build: {
    // Для Electron приложений часто используют relative base
    rollupOptions: {
      // Если нужно указать, что все ресурсы находятся относительно html
      // output: {
      //   entryFileNames: `assets/[name].[hash].js`,
      //   chunkFileNames: `assets/[name].[hash].js`,
      //   assetFileNames: `assets/[name].[hash].[ext]`
      // }
    }
  }
})
