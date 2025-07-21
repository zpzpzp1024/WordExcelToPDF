import { defineConfig } from 'vite';
import plugin from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
    plugins: [plugin(),
    ],
    server: {
        port: 5173,
    },
    base: "/",
    build: {
        outDir: "../wpfApp/wwwroot",
        emptyOutDir: true,
    }
})
