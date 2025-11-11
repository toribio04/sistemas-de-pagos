import { Component, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { ExcelService } from '../../services/excel.service';
import { PaymentData } from '../../models/payment.model';

@Component({
  selector: 'app-payment-form',
  templateUrl: './payment-form.component.html',
  styleUrls: ['./payment-form.component.css']
})
export class PaymentFormComponent implements OnInit {
  paymentForm: FormGroup;
  empresas: string[] = [
    'Empresa A',
    'Empresa B',
    'Empresa C',
    'Institutos Educativos Parroquiales',
    'Otra Empresa'
  ];

  constructor(
    private fb: FormBuilder,
    private excelService: ExcelService
  ) {
    this.paymentForm = this.fb.group({
      nombre: ['', [Validators.required]],
      apellido: ['', [Validators.required]],
      empresa: ['', [Validators.required]],
      importe: ['', [Validators.required, Validators.min(0.01)]],
      formaPago: ['Tarjeta', [Validators.required]]
    });
  }

  ngOnInit(): void {
  }

  onSubmit(): void {
    if (this.paymentForm.valid) {
      const formData: PaymentData = {
        nombre: this.paymentForm.value.nombre,
        apellido: this.paymentForm.value.apellido,
        empresa: this.paymentForm.value.empresa,
        importe: parseFloat(this.paymentForm.value.importe),
        formaPago: this.paymentForm.value.formaPago,
        fecha: new Date().toLocaleString('es-ES')
      };

      this.excelService.saveToExcel(formData, true);
      alert('Pago confirmado y guardado en Excel');
      this.paymentForm.reset({ formaPago: 'Tarjeta' });
    } else {
      alert('Por favor, complete todos los campos correctamente');
      this.markFormGroupTouched();
    }
  }

  private markFormGroupTouched(): void {
    Object.keys(this.paymentForm.controls).forEach(key => {
      const control = this.paymentForm.get(key);
      control?.markAsTouched();
    });
  }
}

