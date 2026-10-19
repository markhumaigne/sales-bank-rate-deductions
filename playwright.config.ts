import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './e2e',
  timeout: 60000,
  workers: 1,
  fullyParallel: false,
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4300',
    browserName: 'chromium',
    channel: 'chrome',
    headless: true,
    timezoneId: 'Asia/Manila',
    viewport: { width: 1440, height: 1100 },
    trace: 'off',
    screenshot: 'off',
    video: 'off',
  },
  // Dedicated test ports keep an open accounting session untouched and ensure fresh API code.
  webServer: [
    {
      command: 'npm run start -w backend',
      url: 'http://127.0.0.1:3100/api/config/defaults',
      env: { PORT: '3100' },
      reuseExistingServer: false,
      timeout: 60000,
    },
    {
      command:
        'npm exec -w frontend -- ng serve --host 127.0.0.1 --port 4300 --proxy-config proxy.e2e.json',
      url: 'http://127.0.0.1:4300',
      reuseExistingServer: false,
      timeout: 60000,
    },
  ],
});
