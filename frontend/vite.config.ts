// frontend/vite.config.ts
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
// import tailwindcss from "@tailwindcss/vite"; // Убран, так как используем v3 через postcss.config.js
// Проверим, как импортируется yaml. Документация vite-plugin-yaml показывает, что он экспортируется как default.
// https://www.npmjs.com/package/vite-plugin-yaml
import yaml from "vite-plugin-yaml"; // Импортируем yaml как default export

// @ts-expect-error process is a nodejs global
const host = process.env.TAURI_DEV_HOST;

// https://vite.dev/config/
// Используем синхронный экспорт для упрощения и избежания ошибок типов с async defineConfig
export default defineConfig({
  plugins: [
    react(),
    // tailwindcss(), // Убран, так как используем v3 через postcss.config.js
    yaml, // Передаём импортированный плагин напрямую, без вызова (). Он должен быть совместимым объектом плагина.
  ],

  // Vite options tailored for Tauri development and only applied in `tauri dev` or `tauri build`
  //
  // 1. prevent Vite from obscuring rust errors
  clearScreen: false,
  // 2. tauri expects a fixed port, fail if that port is not available
  server: {
    port: 1420,
    strictPort: true,
    host: host || false,
    hmr: host
      ? {
          protocol: "ws",
          host,
          port: 1421,
        }
      : undefined,
    watch: {
      // 3. tell Vite to ignore watching `src-tauri`
      ignored: ["**/src-tauri/**"],
    },
  },
});
