
export const formatCurrency = (value: number | undefined | null): string => {
  if (value === undefined || value === null) return '0';
  // Sử dụng dấu CHẤM (.) cho phần phân cách hàng nghìn
  return Math.round(value).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
};

export const formatDecimal = (value: number | undefined | null): string => {
  if (value === undefined || value === null) return '0,00';
  // Sử dụng dấu PHẨY (,) cho phần thập phân
  return value.toFixed(2).replace('.', ',');
};

export const splitName = (fullName: string): { hoLot: string; ten: string } => {
  const parts = (fullName || '').trim().split(/\s+/);
  if (parts.length <= 1) return { hoLot: '', ten: fullName || '' };
  const ten = parts.pop() || '';
  const hoLot = parts.join(' ');
  return { hoLot, ten };
};

/**
 * Trích xuất Khóa từ tên Lớp. 
 * Ví dụ: 48K01.1 -> 48K
 */
export const extractKHoa = (lop: string): string => {
  if (!lop) return 'N/A';
  // Tìm chuỗi số kết hợp với chữ K ở đầu lớp
  const match = lop.match(/^(\d+K)/i);
  if (match) return match[1].toUpperCase();
  
  // Trường hợp không có chữ K, lấy 2 số đầu
  const digits = lop.match(/^\d{2}/);
  return digits ? digits[0] : 'N/A';
};
