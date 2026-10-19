import { bootstrapApplication } from '@angular/platform-browser';
import { provideHttpClient } from '@angular/common/http';
import { DashboardComponent } from './app/components/dashboard/dashboard.component';
bootstrapApplication(DashboardComponent, { providers: [provideHttpClient()] }).catch(console.error);
