
import { Student, BudgetEntry, ScholarshipStep, ScholarshipLevel } from '../types';
import { THRESHOLDS, SCHOLARSHIP_MATRIX } from '../constants';
import { extractKHoa } from '../utils/formatters';

/**
 * Tính số tiền học bổng dựa trên loại xếp loại và học phí
 * Xuất sắc: 110%, Giỏi: 105%, Khá: 100%
 */
const calculateAmountByLevel = (tuition: number, level: ScholarshipLevel): number => {
  switch (level) {
    case 'Xuất sắc': return tuition * 1.1;
    case 'Giỏi': return tuition * 1.05;
    case 'Khá': return tuition;
    default: return 0;
  }
};

export const calculateScholarshipLevel = (diem4: number, diemRL: number): ScholarshipLevel => {
  let hocTap = '';
  if (diem4 >= 3.6) hocTap = 'Xuất sắc';
  else if (diem4 >= 3.2) hocTap = 'Giỏi';
  else if (diem4 >= 2.5) hocTap = 'Khá';
  else return 'Không đạt';

  let renLuyen = '';
  if (diemRL >= 90) renLuyen = 'Xuất sắc';
  else if (diemRL >= 80) renLuyen = 'Giỏi';
  else if (diemRL >= 70) renLuyen = 'Tốt';
  else if (diemRL >= 65) renLuyen = 'Khá';
  else return 'Không đạt';

  return SCHOLARSHIP_MATRIX[hocTap]?.[renLuyen] || 'Không đạt';
};

export const getSortOrder = (a: Student, b: Student) => {
  const levelRank: Record<string, number> = { 'Xuất sắc': 3, 'Giỏi': 2, 'Khá': 1, 'Không đạt': 0 };
  if (levelRank[a.xepLoaiHB] !== levelRank[b.xepLoaiHB]) {
    return levelRank[b.xepLoaiHB] - levelRank[a.xepLoaiHB];
  }
  if (a.diem4 !== b.diem4) return b.diem4 - a.diem4;
  if (a.diem10 !== b.diem10) return b.diem10 - a.diem10;
  if (a.diemRenLuyen !== b.diemRenLuyen) return b.diemRenLuyen - a.diemRenLuyen;
  if (a.soTinChi !== b.soTinChi) return b.soTinChi - a.soTinChi;
  return 0;
};

export const processScholarship = (
  step: ScholarshipStep,
  students: Student[],
  budgets: BudgetEntry[],
  alreadyReceived: Set<string>
): Student[] => {
  const threshold = THRESHOLDS[step];
  const results: Student[] = [];

  const candidates = students.filter(s => {
    if (alreadyReceived.has(s.maSV)) return false;
    if (step === ScholarshipStep.TALENT) return s.hbTaiNangFlag === true;
    if (step === ScholarshipStep.HARDSHIP) return s.hoanCanh && s.hoanCanh.trim() !== '';
    if (step === ScholarshipStep.INTERNATIONAL) return s.quocTich && s.quocTich.toLowerCase() !== 'việt nam';
    return true;
  });

  const stepStudents = candidates.map(s => {
    let reason = '';
    const minTC = s.coKhoaLuan ? threshold.minCreditsLow : threshold.minCredits;
    
    if (s.xepLoaiHB === 'Không đạt') {
      reason = s.diem4 < 2.5 ? 'Điểm học tập < Khá' : 'Điểm rèn luyện < Khá';
    }
    else if (s.soTinChiNo > 0) reason = 'Còn nợ tín chỉ';
    else if (s.soTinChi < minTC) reason = `Tín chỉ < ${minTC}`;
    else if (s.diem4 < threshold.minDiem4) reason = `GPA < ${threshold.minDiem4}`;
    else if (s.diemRenLuyen < threshold.minRenLuyen) reason = `DRL < ${threshold.minRenLuyen}`;

    return { ...s, tempReason: reason };
  }) as (Student & { tempReason: string })[];

  const eligible = stepStudents.filter(s => !s.tempReason).sort(getSortOrder);
  const ineligible = stepStudents.filter(s => s.tempReason);

  if (step !== ScholarshipStep.ACADEMIC) {
    const budgetMap: Record<string, number> = {};
    const relevantBudgets = budgets.filter(b => {
      const noiDung = (b.noiDung || '').toLowerCase();
      if (step === ScholarshipStep.TALENT) return noiDung.includes('tài năng');
      if (step === ScholarshipStep.HARDSHIP) return noiDung.includes('vượt khó');
      if (step === ScholarshipStep.INTERNATIONAL) return noiDung.includes('quốc tế');
      return false;
    });

    relevantBudgets.forEach(b => {
      Object.entries(b.phanBo).forEach(([k, val]) => {
        budgetMap[k.trim().toLowerCase()] = (budgetMap[k.trim().toLowerCase()] || 0) + val;
      });
    });

    const runningLuyKe: Record<string, number> = {};

    eligible.forEach(s => {
      const khoaHienTai = extractKHoa(s.lop);
      const khoaKey = `khóa ${khoaHienTai}`.toLowerCase(); 
      const phanBoHienTai = budgetMap[khoaKey] || 0;
      const currentSpent = runningLuyKe[khoaKey] || 0;
      const amount = calculateAmountByLevel(s.hocPhi, s.xepLoaiHB);

      const isOverBudget = (currentSpent + amount) > phanBoHienTai;
      runningLuyKe[khoaKey] = currentSpent + amount;

      results.push({
        ...s, 
        soTienHB: amount, 
        congDonTienHB: runningLuyKe[khoaKey],
        soTienPhanBo: phanBoHienTai, 
        soTienPhanBoConLai: phanBoHienTai - runningLuyKe[khoaKey],
        ketLuan: 'Đạt', 
        ghiChu: [s.ghiChu, isOverBudget ? 'Thiếu tiền (Vượt quỹ)' : ''].filter(Boolean).join('; ')
      });
    });
  } else {
    // Logic cho Học tập & Rèn luyện
    const majorGroups = new Map<string, (Student & { tempReason: string })[]>();
    eligible.forEach(s => {
      const key = (s.nganh || '').trim().toLowerCase();
      const list = majorGroups.get(key) || [];
      list.push(s);
      majorGroups.set(key, list);
    });

    majorGroups.forEach((group, nganhKey) => {
      const budgetEntry = budgets.find(b => (b.nganh || '').trim().toLowerCase() === nganhKey);

      const groupByKhoa = new Map<string, Student[]>();
      group.forEach(s => {
        const khoaKey = `khóa ${extractKHoa(s.lop)}`.toLowerCase();
        const list = groupByKhoa.get(khoaKey) || [];
        list.push(s);
        groupByKhoa.set(khoaKey, list);
      });

      groupByKhoa.forEach((khoaStudents, khoaKey) => {
        let phanBoTotal = 0;
        if (budgetEntry) {
            const rawKey = Object.keys(budgetEntry.phanBo).find(k => k.trim().toLowerCase() === khoaKey);
            phanBoTotal = rawKey ? budgetEntry.phanBo[rawKey] : 0;
        }

        let currentSpent = 0;
        const awardedIds = new Set<string>();

        const trainingTypes = Array.from(new Set(khoaStudents.map(s => s.loaiHinhDaoTao)));
        trainingTypes.sort().forEach(tt => {
          const topInType = khoaStudents.filter(s => s.loaiHinhDaoTao === tt).sort(getSortOrder)[0];
          if (topInType) {
            const amount = calculateAmountByLevel(topInType.hocPhi, topInType.xepLoaiHB);
            awardedIds.add(topInType.maSV);
            currentSpent += amount;
            const isOverBudget = currentSpent > phanBoTotal;
            results.push({ 
              ...topInType, 
              soTienHB: amount, 
              congDonTienHB: currentSpent, 
              soTienPhanBo: phanBoTotal, 
              soTienPhanBoConLai: phanBoTotal - currentSpent, 
              ketLuan: 'Đạt', 
              ghiChu: [topInType.ghiChu, isOverBudget ? 'Ưu tiên tối thiểu (Vượt quỹ)' : 'Ưu tiên tối thiểu'].filter(Boolean).join('; ') 
            });
          }
        });

        let isExhausted = currentSpent >= phanBoTotal; 
        khoaStudents.filter(s => !awardedIds.has(s.maSV)).sort(getSortOrder).forEach(s => {
          const amount = calculateAmountByLevel(s.hocPhi, s.xepLoaiHB);
          const remaining = phanBoTotal - currentSpent;

          if (isExhausted) {
            results.push({ ...s, soTienHB: 0, congDonTienHB: currentSpent, soTienPhanBo: phanBoTotal, soTienPhanBoConLai: phanBoTotal - currentSpent, ketLuan: 'Dự phòng', ghiChu: [s.ghiChu, 'Hết ngân sách'].filter(Boolean).join('; ') });
            return;
          }

          if (amount <= remaining) {
            currentSpent += amount;
            results.push({ ...s, soTienHB: amount, congDonTienHB: currentSpent, soTienPhanBo: phanBoTotal, soTienPhanBoConLai: phanBoTotal - currentSpent, ketLuan: 'Đạt', ghiChu: [s.ghiChu, 'Xếp hạng điểm'].filter(Boolean).join('; ') });
          } else if (remaining >= 0.5 * amount) {
            currentSpent += amount;
            isExhausted = true; 
            results.push({ ...s, soTienHB: amount, congDonTienHB: currentSpent, soTienPhanBo: phanBoTotal, soTienPhanBoConLai: phanBoTotal - currentSpent, ketLuan: 'Đạt', ghiChu: [s.ghiChu, 'Dư > 50% HP'].filter(Boolean).join('; ') });
          } else {
            isExhausted = true; 
            results.push({ ...s, soTienHB: 0, congDonTienHB: currentSpent, soTienPhanBo: phanBoTotal, soTienPhanBoConLai: phanBoTotal - currentSpent, ketLuan: 'Dự phòng', ghiChu: [s.ghiChu, 'Hết ngân sách'].filter(Boolean).join('; ') });
          }
        });
      });
    });
  }

  ineligible.forEach(s => {
    results.push({ ...s, soTienHB: 0, ketLuan: 'Loại', ghiChu: [s.ghiChu, s.tempReason].filter(Boolean).join('; ') });
  });

  return results;
};
