import { Injectable, PLATFORM_ID, Inject } from '@angular/core';
import { isPlatformBrowser } from '@angular/common';
import * as XLSX from 'xlsx';
import { PaymentData } from '../models/payment.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {
  private fileName: string = 'pagos_banco_robles.xlsx';
  private sheetName: string = 'Pagos';
  private storageKey: string = 'excelData';
  private isBrowser: boolean;

  constructor(@Inject(PLATFORM_ID) private platformId: Object) {
    this.isBrowser = isPlatformBrowser(this.platformId);
    // Verificar si el archivo existe, si no, crear uno nuevo con headers
    if (this.isBrowser) {
      this.initializeExcelFile();
    }
  }

  private initializeExcelFile(): void {
    if (!this.isBrowser) {
      return;
    }
    try {
      // Intentar leer el archivo existente
      const existingData = localStorage.getItem(this.storageKey);
      if (!existingData) {
        // Crear archivo nuevo con headers
        const headers = [
          ['Nombre', 'Apellido', 'Empresa', 'Importe', 'Forma de Pago', 'Fecha']
        ];
        const ws = XLSX.utils.aoa_to_sheet(headers);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, this.sheetName);
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        localStorage.setItem(this.storageKey, this.arrayBufferToBase64(excelBuffer));
      }
    } catch (error) {
      console.error('Error al inicializar archivo Excel:', error);
    }
  }

  saveToExcel(data: PaymentData, autoDownload: boolean = true): void {
    if (!this.isBrowser) {
      console.warn('ExcelService: No se puede guardar en el servidor');
      return;
    }
    try {
      // Leer datos existentes desde localStorage
      const existingData = localStorage.getItem(this.storageKey);
      let workbook: XLSX.WorkBook;
      let worksheet: XLSX.WorkSheet;

      if (existingData) {
        // Convertir base64 a ArrayBuffer
        const arrayBuffer = this.base64ToArrayBuffer(existingData);
        workbook = XLSX.read(arrayBuffer, { type: 'array' });
        worksheet = workbook.Sheets[this.sheetName];
      } else {
        // Crear nuevo workbook
        workbook = XLSX.utils.book_new();
        const headers = [['Nombre', 'Apellido', 'Empresa', 'Importe', 'Forma de Pago', 'Fecha']];
        worksheet = XLSX.utils.aoa_to_sheet(headers);
        XLSX.utils.book_append_sheet(workbook, worksheet, this.sheetName);
      }

      // Convertir worksheet a JSON para agregar nueva fila
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as unknown[][];
      
      // Agregar nueva fila
      jsonData.push([
        data.nombre,
        data.apellido,
        data.empresa,
        data.importe,
        data.formaPago,
        data.fecha
      ]);

      // Crear nuevo worksheet con los datos actualizados
      const newWorksheet = XLSX.utils.aoa_to_sheet(jsonData);
      workbook.Sheets[this.sheetName] = newWorksheet;

      // Guardar en localStorage
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      localStorage.setItem(this.storageKey, this.arrayBufferToBase64(excelBuffer));

      // Descargar el archivo actualizado si autoDownload es true
      if (autoDownload) {
        this.downloadExcel(workbook);
      }
    } catch (error) {
      console.error('Error al guardar en Excel:', error);
      if (this.isBrowser) {
        alert('Error al guardar los datos. Por favor, intente nuevamente.');
      }
    }
  }

  private downloadExcel(workbook: XLSX.WorkBook): void {
    if (!this.isBrowser) {
      return;
    }
    try {
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(data);
      const link = document.createElement('a');
      link.href = url;
      link.download = this.fileName;
      link.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Error al descargar Excel:', error);
    }
  }

  downloadExcelFile(): void {
    if (!this.isBrowser) {
      return;
    }
    try {
      const existingData = localStorage.getItem(this.storageKey);
      if (existingData) {
        const arrayBuffer = this.base64ToArrayBuffer(existingData);
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        this.downloadExcel(workbook);
      } else {
        if (this.isBrowser) {
          alert('No hay datos para descargar');
        }
      }
    } catch (error) {
      console.error('Error al descargar archivo Excel:', error);
      if (this.isBrowser) {
        alert('Error al descargar el archivo. Por favor, intente nuevamente.');
      }
    }
  }

  getAllPayments(): PaymentData[] {
    if (!this.isBrowser) {
      return [];
    }
    try {
      const existingData = localStorage.getItem(this.storageKey);
      if (existingData) {
        const arrayBuffer = this.base64ToArrayBuffer(existingData);
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheet = workbook.Sheets[this.sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as unknown[][];
        
        // Omitir la primera fila (headers) y convertir a PaymentData
        return jsonData.slice(1).map((row: unknown[]) => ({
          nombre: String(row[0] || ''),
          apellido: String(row[1] || ''),
          empresa: String(row[2] || ''),
          importe: Number(row[3] || 0),
          formaPago: String(row[4] || ''),
          fecha: String(row[5] || '')
        }));
      }
      return [];
    } catch (error) {
      console.error('Error al obtener pagos:', error);
      return [];
    }
  }

  clearData(): void {
    if (!this.isBrowser) {
      return;
    }
    localStorage.removeItem(this.storageKey);
    this.initializeExcelFile();
  }

  private arrayBufferToBase64(buffer: ArrayBuffer): string {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  private base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binaryString = atob(base64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  }
}

