import { NestFactory } from '@nestjs/core';
import { json } from 'express';
import { AppModule } from './app.module';
export async function createApplication(quiet = false) {
  const app = await NestFactory.create(AppModule, {
    bodyParser: false,
    ...(quiet ? { logger: false as const } : {}),
  });
  app.use(json({ limit: '25mb' }));
  app.use(
    (
      _request: unknown,
      response: { setHeader: (key: string, value: string) => void },
      next: () => void,
    ) => {
      response.setHeader('Cache-Control', 'no-store');
      response.setHeader('X-Content-Type-Options', 'nosniff');
      next();
    },
  );
  return app;
}
