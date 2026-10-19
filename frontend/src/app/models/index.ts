export * from '../../../../shared/models';
import { Selection, SheetPreview } from '../../../../shared/models';
export interface ReportUpload {
  file: File;
  selection: Selection;
  preview: SheetPreview;
  valid: boolean;
}
