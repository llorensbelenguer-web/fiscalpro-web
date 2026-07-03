import { defineConfig } from 'astro/config';
import tailwind from '@astrojs/tailwind';

export default defineConfig({
  site: 'https://llorensbelenguer-web.github.io',
  base: '/fiscalpro-web',
  integrations: [tailwind()],
});
