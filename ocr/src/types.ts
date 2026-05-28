export interface Vendor {
  name: string;
  taxId: string;
  address: string;
  phone: string;
  email: string;
  website: string;
}

export interface Customer {
  name: string;
  taxId: string;
  address: string;
  email: string;
}

export interface LineItem {
  description: string;
  quantity: number;
  unitPrice: number;
  taxRate: number; // e.g. 21 representing 21%
  amount: number;
}

export interface TaxLine {
  taxRate: number; // e.g. 21
  taxableAmount: number;
  taxAmount: number;
}

export interface InvoiceData {
  invoiceNumber: string;
  invoiceDate: string;
  dueDate: string;
  poNumber: string;
  paymentTerms: string;
  vendor: Vendor;
  customer: Customer;
  lineItems: LineItem[];
  taxes: TaxLine[];
  subtotal: number;
  discount: number;
  taxTotal: number;
  totalAmount: number;
  currency: string;
  paymentInstructions: string;
  summaryOfAccuracy: string;
}

export interface OCRHistoryEntry {
  id: string;
  fileName: string;
  fileType: string;
  uploadedAt: string;
  invoiceData: InvoiceData;
  fileData?: string; // Cache base64 preview
}
