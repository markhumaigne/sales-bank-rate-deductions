import { BadRequestException } from '@nestjs/common';
import { memoryStorage } from 'multer';
export const uploadOptions = {
  storage: memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024, files: 1, fields: 8, fieldSize: 1024 * 1024 },
  fileFilter: (
    _request: unknown,
    file: Express.Multer.File,
    callback: (error: Error | null, accept: boolean) => void,
  ) => {
    if (!/\.xlsx?$/i.test(file.originalname))
      callback(new BadRequestException('Upload an .xlsx or .xls file.'), false);
    else callback(null, true);
  },
};
