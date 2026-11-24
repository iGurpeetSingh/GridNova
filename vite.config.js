import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    host: '0.0.0.0',
  },
  build: {
    // Limit simultaneous file reads/transforms to reduce peak memory usage on
    // Vercel's standard build container. This does not change app output.
    rollupOptions: {
      maxParallelFileOps: 1,
    },
  },
})
