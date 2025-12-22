
import { ScholarshipLevel } from './types';

/**
 * Ma trận xếp loại Học bổng dựa trên sự kết hợp giữa Học tập và Rèn luyện
 * Một sinh viên đạt loại Giỏi học tập nhưng chỉ Tốt rèn luyện thì chỉ nhận HB Khá.
 */
export const SCHOLARSHIP_MATRIX: Record<string, Record<string, ScholarshipLevel>> = {
  'Xuất sắc': {
    'Xuất sắc': 'Xuất sắc',
    'Giỏi': 'Giỏi',
    'Tốt': 'Khá',
    'Khá': 'Khá',
  },
  'Giỏi': {
    'Xuất sắc': 'Giỏi',
    'Giỏi': 'Giỏi',
    'Tốt': 'Khá',
    'Khá': 'Khá',
  },
  'Khá': {
    'Xuất sắc': 'Khá',
    'Giỏi': 'Khá',
    'Tốt': 'Khá',
    'Khá': 'Khá',
  },
};

export const STEP_NAMES = {
  TALENT: 'Học bổng Tài năng',
  HARDSHIP: 'Học bổng Vượt khó',
  INTERNATIONAL: 'Học bổng Quốc tế',
  ACADEMIC: 'Học bổng Học tập & Rèn luyện',
};

export const THRESHOLDS = {
  TALENT: { minDiem4: 3.2, minRenLuyen: 80, minCredits: 15, minCreditsLow: 10 },
  HARDSHIP: { minDiem4: 2.5, minRenLuyen: 65, minCredits: 15, minCreditsLow: 10 },
  INTERNATIONAL: { minDiem4: 2.5, minRenLuyen: 65, minCredits: 15, minCreditsLow: 10 },
  ACADEMIC: { minDiem4: 2.5, minRenLuyen: 65, minCredits: 15, minCreditsLow: 10 },
};
