
export enum ScholarshipStep {
  TALENT = 'TALENT',
  HARDSHIP = 'HARDSHIP',
  INTERNATIONAL = 'INTERNATIONAL',
  ACADEMIC = 'ACADEMIC',
}

export type ScholarshipLevel = 'Xuất sắc' | 'Giỏi' | 'Khá' | 'Không đạt';

export interface Student {
  tt: number;
  hoTen: string;
  ngaySinh: string;
  maSV: string;
  lop: string;
  khoa: string;
  nganh: string;
  loaiHinhDaoTao: string;
  hbTaiNangFlag: boolean;
  hoanCanh: string;
  quocTich: string;
  hocChuyenTiep: boolean;
  coKhoaLuan: boolean;
  soTinChi: number;
  soTinChiNo: number;
  diem4: number;
  diem10: number;
  diemRenLuyen: number;
  hocPhi: number;
  // Computed fields
  xepLoaiHB: ScholarshipLevel;
  maHocKy?: string;
  soTienHB?: number;
  congDonTienHB?: number;
  soTienPhanBo?: number;
  soTienPhanBoConLai?: number;
  ketLuan?: string;
  ghiChu?: string;
  diemTBTichLuy?: number;
  originalRow?: any; // Lưu trữ row nguyên bản từ file Excel nhập vào
}

export interface BudgetEntry {
  stt: number;
  noiDung?: string;
  nganh?: string;
  phanBo: Record<string, number>; // Key: "Khóa 48", Value: Amount
  maHocKy: string;
}

export interface ProcessingResult {
  step: ScholarshipStep;
  students: Student[];
  totalAllocated: number;
  remainingBudget: number;
}
