import { Module } from '@nestjs/common';
import { ExcelController } from './excel/excel.controller';
import { ExcelService } from './excel/excel.service';
import { ExcelExportService } from './excel/excel-export.service';
import { CalculationService } from './calculation/calculation.service';
import { ConfigService } from './config/config.service';
@Module({
  controllers: [ExcelController],
  providers: [ExcelService, ExcelExportService, CalculationService, ConfigService],
})
export class AppModule {}
