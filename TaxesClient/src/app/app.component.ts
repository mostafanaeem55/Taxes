import { CommonModule } from '@angular/common';
import { HttpClient, HttpClientModule, HttpErrorResponse, HttpEvent, HttpEventType } from '@angular/common/http';
import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet,CommonModule,HttpClientModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  file: File | null = null;
  message: { type: string, text: string } | null = null;

  constructor(private http: HttpClient) { }

  onFileChange(event: Event) {
    const element = event.target as HTMLInputElement;
    if (element.files && element.files.length) {
      this.file = element.files[0];
    }
  }

  uploadFile() {
    if (!this.file) {
      this.showMessage('danger', 'Please select a file.');
      return;
    }
  
    const formData = new FormData();
    formData.append('file', this.file);
  
    this.http.post('https://localhost:7038/api/Excel', formData, {
      reportProgress: true,
      observe: 'events',
      responseType: 'blob'
    }).subscribe(
      (event: HttpEvent<any>) => {
        switch (event.type) {
          case HttpEventType.UploadProgress:
            if (event.total) {
              const progress = Math.round(100 * event.loaded / event.total);
              console.log(`Upload progress: ${progress}%`);
            } else {
              console.log(`Uploaded ${event.loaded} bytes`);
            }
            break;
          case HttpEventType.Response:
            // Handle the downloaded Excel file (Blob)
            const blob = new Blob([event.body!], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'modified_taxes.xlsx';
            a.click();
            window.URL.revokeObjectURL(url);
            this.showMessage('success', 'File uploaded and processed successfully.');
            break;
        }
      },
      (error: HttpErrorResponse) => {
        if (error.error instanceof ErrorEvent) {
          // Client-side error
          this.showMessage('danger', `Error: ${error.error.message}`);
        } else {
          // Backend returned an unsuccessful response code
          this.showMessage('danger', `Error: ${error.status} - ${error.statusText}`);
          console.error('Backend Error:', error.error);
        }
      }
    );
  }

  private showMessage(type: string, text: string) {
    this.message = { type, text };
    setTimeout(() => {
      this.message = null;
    }, 5000);
  }}
