
import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileUp, Settings, Users, CreditCard, Play, Download, 
  BarChart3, FileSpreadsheet, Layers, GraduationCap, 
  CheckCircle2, Info, Calendar, ArrowRightCircle, AlertTriangle, 
  Search, Eye, CheckCircle, ChevronRight, FileText, TrendingDown
} from 'lucide-react';
import { Student, BudgetEntry, ScholarshipStep } from './types';
import { STEP_NAMES } from './constants';
import { calculateScholarshipLevel, processScholarship } from './services/scholarshipLogic';
import { formatCurrency, formatDecimal, splitName, extractKHoa } from './utils/formatters';

interface Notification {
  type: 'success' | 'error' | 'info' | 'warning';
  message: string;
}

const App: React.FC = () => {
  const [semesterCode, setSemesterCode] = useState('252');
  const [budgetClassData, setBudgetClassData] = useState<BudgetEntry[]>([]);
  const [budgetMajorData, setBudgetMajorData] = useState<BudgetEntry[]>([]);
  const [studentData, setStudentData] = useState<Student[]>([]);
  const [results, setResults] = useState<Record<ScholarshipStep, Student[]>>({
    [ScholarshipStep.TALENT]: [],
    [ScholarshipStep.HARDSHIP]: [],
    [ScholarshipStep.INTERNATIONAL]: [],
    [ScholarshipStep.ACADEMIC]: [],
  });
  const [isProcessing, setIsProcessing] = useState(false);
  const [activeTab, setActiveTab] = useState<'import' | 'data-check' | 'process' | 'results'>('import');
  const [notification, setNotification] = useState<Notification | null>(null);
  const [searchTerm, setSearchTerm] = useState('');

  const COLORS = {
    orange: '#F37021',
    green: '#00A651',
    blue: '#0054A6',
    background: '#F8FAFC',
    text: '#1E293B',
    border: '#E2E8F0'
  };

  useEffect(() => {
    if (notification) {
      const timer = setTimeout(() => setNotification(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [notification]);

  const showNotify = (message: string, type: 'success' | 'error' | 'info' | 'warning' = 'success') => {
    setNotification({ message, type });
  };

  const downloadTemplate = (type: 'BUDGET_CLASS' | 'BUDGET_MAJOR' | 'STUDENT') => {
    const wb = XLSX.utils.book_new();
    let data: any[] = [];
    if (type === 'BUDGET_CLASS') {
      data = [
        { 'STT': 1, 'Nội dung': 'Học bổng Tài năng', 'Mã học kỳ': semesterCode, 'Khóa 48K': 500000000, 'Khóa 49K': 400000000, 'Khóa 50K': 300000000, 'Khóa 51K': 300000000 },
        { 'STT': 2, 'Nội dung': 'Học bổng Vượt khó', 'Mã học kỳ': semesterCode, 'Khóa 48K': 300000000, 'Khóa 49K': 300000000, 'Khóa 50K': 300000000, 'Khóa 51K': 300000000 },
        { 'STT': 3, 'Nội dung': 'Học bổng Quốc tế', 'Mã học kỳ': semesterCode, 'Khóa 48K': 200000000, 'Khóa 49K': 200000000, 'Khóa 50K': 200000000, 'Khóa 51K': 200000000 }
      ];
    } else if (type === 'BUDGET_MAJOR') {
      const allMajors = ['Kinh tế', 'Kế toán', 'Kiểm toán', 'Quản trị kinh doanh', 'Marketing', 'Tài chính', 'Ngân hàng', 'Luật', 'Du lịch'];
      data = allMajors.map((m, i) => ({
        'STT': i + 1, 'Ngành': m, 'Mã học kỳ': semesterCode, 'Khóa 48K': 400000000, 'Khóa 49K': 350000000, 'Khóa 50K': 300000000, 'Khóa 51K': 300000000
      }));
    } else {
      const khoas = ['48K', '49K', '50K', '51K'];
      for (let i = 1; i <= 1500; i++) {
        const k = khoas[i % 4];
        data.push({
          'STT': i,
          'Họ lót': `Sinh viên`,
          'Tên': `${i}`,
          'Ngày sinh': '01/01/2005',
          'Mã sinh viên': `${k}${i}`,
          'Lớp': `${k}.S`,
          'Khoa quản lý': 'KINH TẾ',
          'Ngành': 'Kinh tế',
          'Loại hình đào tạo': 'S',
          'CTSV nhập ưu tiên(HB Tài năng Hoàn cảnh của sinh viên)': i % 10 === 0 ? 'Tài năng' : (i % 15 === 0 ? 'Hộ nghèo' : false),
          'Quốc tịch': 'Việt Nam',
          'Khóa luận TN': false,
          'Tín chỉ học lần đầu': 20,
          'Tín chỉ nợ': 0,
          'Điểm rèn luyện': 85,
          'Điểm Thang 4': 3.5,
          'Điểm Thang 10': 8.75,
          'Học phí L1': 15000000
        });
      }
    }
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, "Dữ liệu");
    XLSX.writeFile(wb, `MAU_${type}.xlsx`);
  };

  const handleBudgetUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'CLASS' | 'MAJOR') => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]) as any[];
      const parsed = json.map((row: any) => ({
        stt: row['STT'] || 0, noiDung: row['Nội dung'] || '', nganh: row['Ngành'] || '', 
        maHocKy: String(row['Mã học kỳ'] || semesterCode),
        phanBo: Object.keys(row).reduce((acc: any, key) => { 
          if (key.includes('Khóa')) acc[key.trim().toLowerCase()] = Number(row[key]) || 0; 
          return acc; 
        }, {})
      }));
      if (type === 'CLASS') setBudgetClassData(parsed); else setBudgetMajorData(parsed);
      showNotify("Nhập ngân sách thành công.");
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const handleStudentUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]) as any[];
      const parsed: Student[] = json.map((row: any) => {
        const d4 = Number(row['Điểm Thang 4'] || row['Điểm học tập (Theo thang 4)'] || 0);
        const drl = Number(row['Điểm rèn luyện'] || 0);
        
        let ngaySinh = row['Ngày sinh'];
        if (typeof ngaySinh === 'number') {
          const date = new Date(Math.round((ngaySinh - 25569) * 86400 * 1000));
          ngaySinh = `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}`;
        } else {
          ngaySinh = String(ngaySinh || '');
        }

        const hoLot = String(row['Họ lót'] || '').trim();
        const ten = String(row['Tên'] || '').trim();
        const hoTen = hoLot ? `${hoLot} ${ten}` : (ten || String(row['Họ và tên'] || ''));
        
        const uuTienStr = String(row['CTSV nhập ưu tiên(HB Tài năng Hoàn cảnh của sinh viên)'] || row['CTSV nhập ưu tiên'] || row['Hoàn cảnh của sinh viên'] || '').trim();
        const isFalse = uuTienStr.toLowerCase() === 'false' || uuTienStr === '0' || uuTienStr === '';
        
        const thucTapTN = row['Khóa luận TN'] ?? row['Thực tập TN'] ?? row['Khóa luận/ Báo cáo thực tập/Đề án TN'];

        return {
          tt: Number(row['STT']) || 0, 
          hoTen, 
          ngaySinh,
          maSV: String(row['Mã sinh viên']), 
          lop: String(row['Lớp'] || ''), 
          khoa: String(row['Khoa quản lý'] || row['Khoa'] || '').toUpperCase(),
          nganh: String(row['Ngành'] || '').trim(), 
          loaiHinhDaoTao: String(row['Loại hình đào tạo'] || 'S').trim().toUpperCase(),
          hbTaiNangFlag: !isFalse && uuTienStr.toLowerCase().includes('tài năng'), 
          hoanCanh: isFalse ? '' : uuTienStr,
          quocTich: row['Quốc tịch'] || 'Việt Nam', 
          hocChuyenTiep: String(row['Học chuyển tiếp'] || '').toLowerCase() === 'x',
          coKhoaLuan: (String(thucTapTN).toLowerCase() === 'x' || String(thucTapTN) === '1' || String(thucTapTN).toLowerCase() === 'true'),
          soTinChi: Number(row['Tín chỉ học lần đầu'] || row['Số tín chỉ lần đầu']) || 0, 
          soTinChiNo: Number(row['Tín chỉ nợ'] || row['Số tín chỉ nợ']) || 0,
          diem4: d4, 
          diem10: Number(row['Điểm Thang 10'] || row['Điểm học tập (Theo thang 10)']) || 0, 
          diemRenLuyen: drl,
          hocPhi: Number(row['Học phí L1'] || row['Học phí (ĐVT: Đồng)']) || 0, 
          xepLoaiHB: calculateScholarshipLevel(d4, drl)
        } as Student;
      });
      setStudentData(parsed);
      showNotify(`Đã nhập ${parsed.length} sinh viên.`);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const runProcessing = () => {
    if (studentData.length === 0) return;
    setIsProcessing(true);
    setTimeout(() => {
      const rSet = new Set<string>();
      const bCombined = [...budgetClassData, ...budgetMajorData];
      const talent = processScholarship(ScholarshipStep.TALENT, studentData, bCombined, rSet);
      talent.filter(s => s.ketLuan === 'Đạt').forEach(s => rSet.add(s.maSV));
      const hardship = processScholarship(ScholarshipStep.HARDSHIP, studentData, bCombined, rSet);
      hardship.filter(s => s.ketLuan === 'Đạt').forEach(s => rSet.add(s.maSV));
      const international = processScholarship(ScholarshipStep.INTERNATIONAL, studentData, bCombined, rSet);
      international.filter(s => s.ketLuan === 'Đạt').forEach(s => rSet.add(s.maSV));
      const academic = processScholarship(ScholarshipStep.ACADEMIC, studentData, bCombined, rSet);
      setResults({ 
        [ScholarshipStep.TALENT]: talent, [ScholarshipStep.HARDSHIP]: hardship, 
        [ScholarshipStep.INTERNATIONAL]: international, [ScholarshipStep.ACADEMIC]: academic 
      });
      setIsProcessing(false);
      showNotify("Xét duyệt hoàn tất!");
      setActiveTab('results');
    }, 1500);
  };

  const getExportKhoas = () => {
    // FIX: Use Array.from to ensure typed string array from Set (fixes line 187 potential unknown[])
    return Array.from(new Set(studentData.map(s => extractKHoa(s.lop)))).sort();
  };

  const transformTo29Columns = (students: Student[], stepLabel: string) => {
    return students.map(s => {
      const { hoLot, ten } = splitName(s.hoTen);
      return {
        'STT': s.tt, 'Họ lót': hoLot, 'Tên': ten, 'Ngày sinh': s.ngaySinh, 'Mã sinh viên': s.maSV, 'Lớp': s.lop,
        'Khoa': s.khoa, 'Khóa': extractKHoa(s.lop), 'Ngành': s.nganh, 'Loại hình đào tạo': s.loaiHinhDaoTao,
        'HB Tài năng': s.hbTaiNangFlag ? 'x' : '', 'Hoàn cảnh của sinh viên': s.hoanCanh, 'Quốc tịch': s.quocTich,
        'Học chuyển tiếp': s.hocChuyenTiep ? 'x' : '', 'Khóa luận/ Báo cáo thực tập/Đề án TN': s.coKhoaLuan ? 'x' : '',
        'Số tín chỉ lần đầu': s.soTinChi, 'Số tín chỉ nợ': s.soTinChiNo, 'Điểm học tập (Theo thang 4)': formatDecimal(s.diem4),
        'Điểm học tập (Theo thang 10)': formatDecimal(s.diem10), 'Điểm rèn luyện': s.diemRenLuyen, 'Xếp loại học bổng': s.xepLoaiHB,
        'Học phí (ĐVT: Đồng)': s.hocPhi, 'Mã học kỳ': semesterCode, 'Số tiền học bổng': Math.round(s.soTienHB || 0),
        'Cộng dồn tiền HB': Math.round(s.congDonTienHB || 0), 'Số tiền phân bổ': Math.round(s.soTienPhanBo || 0),
        'Số tiền phân bổ còn lại': Math.round(s.soTienPhanBoConLai || 0), 'Kết luận': `${s.ketLuan} (${stepLabel})`,
        'Ghi chú': s.ghiChu || ''
      };
    });
  };

  const exportMasterExcel = () => {
    const wb = XLSX.utils.book_new();
    const summaryData = studentData.map(s => {
      let found: Student | undefined;
      let stepName = 'Không';
      for (const k of Object.keys(results) as ScholarshipStep[]) {
        const m = results[k].find(r => r.maSV === s.maSV);
        if (m && (!found || (m.ketLuan === 'Đạt' && found.ketLuan !== 'Đạt'))) { found = m; stepName = STEP_NAMES[k]; }
        if (m?.ketLuan === 'Đạt') break;
      }
      return transformTo29Columns([found || s], stepName)[0];
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryData), "TỔNG HỢP");
    (Object.keys(results) as ScholarshipStep[]).forEach(step => {
      const data = transformTo29Columns(results[step], STEP_NAMES[step]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), step);
    });
    XLSX.writeFile(wb, `BAO_CAO_TONG_HOP_HK${semesterCode}.xlsx`);
  };

  const exportAdministrativeTN_QT = (step: ScholarshipStep, title: string) => {
    const sList = results[step].filter(s => s.ketLuan === 'Đạt');
    const uniqueKhoas = getExportKhoas();
    const h1 = ['TT', 'Khoa/ Ngành'];
    uniqueKhoas.forEach(k => h1.push(`Khóa ${k}`, ''));
    h1.push('Ghi chú');
    const h2 = ['', ''];
    uniqueKhoas.forEach(() => h2.push('Số SV', 'Tiền cấp'));
    h2.push('');
    const aoa: any[][] = [h1, h2];
    const merges: XLSX.Range[] = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
      { s: { r: 0, c: h1.length - 1 }, e: { r: 1, c: h1.length - 1 } }
    ];
    uniqueKhoas.forEach((_, i) => merges.push({ s: { r: 0, c: 2 + i * 2 }, e: { r: 0, c: 3 + i * 2 } }));
    // FIX: Use Array.from for facultyList to resolve unknown[] error (line 223)
    const facultyList: string[] = Array.from(new Set(studentData.map(s => s.khoa))).sort();
    let stt = 1;
    facultyList.forEach(f => {
      const fStudents = sList.filter(s => s.khoa === f);
      const rowF = [stt++, f];
      uniqueKhoas.forEach(k => {
        const kS = fStudents.filter(s => extractKHoa(s.lop) === k);
        rowF.push(kS.length, Math.round(kS.reduce((a, b) => a + (b.soTienHB || 0), 0)));
      });
      aoa.push(rowF);
      // FIX: Use Array.from for majorList to resolve unknown[] error (line 233)
      const majorList: string[] = Array.from(new Set(studentData.filter(s => s.khoa === f).map(s => s.nganh))).sort();
      majorList.forEach(m => {
        const mS = fStudents.filter(s => s.nganh === m);
        // FIX: Explicitly type rowM as any[] to allow pushing numbers (line 239)
        const rowM: any[] = ['', `- ${m}`];
        uniqueKhoas.forEach(k => {
          const kS = mS.filter(s => extractKHoa(s.lop) === k);
          rowM.push(kS.length, Math.round(kS.reduce((a, b) => a + (b.soTienHB || 0), 0)));
        });
        aoa.push(rowM);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "KẾT QUẢ");
    XLSX.writeFile(wb, `${title}_HK${semesterCode}.xlsx`);
  };

  const exportAdministrativeHardship = () => {
    const sList = results[ScholarshipStep.HARDSHIP].filter(s => s.ketLuan === 'Đạt');
    const uniqueKhoas = getExportKhoas();
    const h1 = ['TT', 'Khoa/ Ngành'];
    uniqueKhoas.forEach(k => h1.push(`Khóa ${k}`, '', '', ''));
    h1.push('TỔNG CỘNG', '');
    const h2 = ['', ''];
    uniqueKhoas.forEach(() => h2.push('Hộ nghèo', '', 'Hộ cận nghèo', ''));
    h2.push('Tổng SV', 'Tổng tiền');
    const h3 = ['', ''];
    uniqueKhoas.forEach(() => h3.push('Số SV', 'Tiền cấp', 'Số SV', 'Tiền cấp'));
    h3.push('', '');
    const aoa: any[][] = [h1, h2, h3];
    const merges: XLSX.Range[] = [
      { s: { r: 0, c: 0 }, e: { r: 2, c: 0 } }, { s: { r: 0, c: 1 }, e: { r: 2, c: 1 } },
      { s: { r: 0, c: h1.length - 2 }, e: { r: 1, c: h1.length - 1 } }
    ];
    uniqueKhoas.forEach((_, i) => {
      merges.push({ s: { r: 0, c: 2 + i * 4 }, e: { r: 0, c: 5 + i * 4 } });
      merges.push({ s: { r: 1, c: 2 + i * 4 }, e: { r: 1, c: 3 + i * 4 } });
      merges.push({ s: { r: 1, c: 4 + i * 4 }, e: { r: 1, c: 5 + i * 4 } });
    });
    // FIX: Use Array.from for facultyList to resolve unknown[] error (line 273)
    const facultyList: string[] = Array.from(new Set(studentData.map(s => s.khoa))).sort();
    let stt = 1;
    facultyList.forEach(f => {
      const fS = sList.filter(s => s.khoa === f);
      const rowF = [stt++, f];
      let tSV = 0, tMoney = 0;
      uniqueKhoas.forEach(k => {
        const kS = fS.filter(s => extractKHoa(s.lop) === k);
        const p = kS.filter(s => s.hoanCanh.includes('Hộ nghèo'));
        const np = kS.filter(s => s.hoanCanh.includes('Hộ cận nghèo'));
        const pM = Math.round(p.reduce((a, b) => a + (b.soTienHB || 0), 0));
        const npM = Math.round(np.reduce((a, b) => a + (b.soTienHB || 0), 0));
        rowF.push(p.length, pM, np.length, npM);
        tSV += kS.length; tMoney += (pM + npM);
      });
      rowF.push(tSV, tMoney);
      aoa.push(rowF);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "5-VK");
    XLSX.writeFile(wb, `MAU_5_VK_HK${semesterCode}.xlsx`);
  };

  const exportAdministrativeAcademic = () => {
    const sList = results[ScholarshipStep.ACADEMIC].filter(s => s.ketLuan === 'Đạt');
    const uniqueKhoas = getExportKhoas();
    const h1 = ['TT', 'Khoa/ Ngành'];
    uniqueKhoas.forEach(k => h1.push(`KHÓA ${k}`, '', '', ''));
    h1.push('TỔNG CỘNG', '');
    const h2 = ['', ''];
    uniqueKhoas.forEach(() => h2.push('Số SV', 'Phân bổ', 'Dự kiến cấp', 'Chênh lệch'));
    h2.push('Tổng SV', 'Tổng tiền');
    const aoa: any[][] = [h1, h2];
    const merges: XLSX.Range[] = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
      { s: { r: 0, c: h1.length - 2 }, e: { r: 0, c: h1.length - 1 } }
    ];
    uniqueKhoas.forEach((_, i) => merges.push({ s: { r: 0, c: 2 + i * 4 }, e: { r: 0, c: 5 + i * 4 } }));
    // FIX: Use Array.from for facultyList to resolve unknown[] error (line 313)
    const facultyList: string[] = Array.from(new Set(studentData.map(s => s.khoa))).sort();
    let stt = 1;
    facultyList.forEach(f => {
      const fS = sList.filter(s => s.khoa === f);
      const rowF = [stt++, f];
      let tSV = 0, tMoney = 0;
      uniqueKhoas.forEach(k => {
        const kS = fS.filter(s => extractKHoa(s.lop) === k);
        const actual = Math.round(kS.reduce((a, b) => a + (b.soTienHB || 0), 0));
        const pb = Math.round(kS.length > 0 ? (kS[0].soTienPhanBo || 0) : 0);
        rowF.push(kS.length, pb, actual, pb - actual);
        tSV += kS.length; tMoney += actual;
      });
      rowF.push(tSV, tMoney);
      aoa.push(rowF);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "4-HTRL");
    XLSX.writeFile(wb, `MAU_4_HTRL_HK${semesterCode}.xlsx`);
  };

  const totalStats = useMemo(() => {
    let t = 0, c = 0;
    (Object.values(results).flat() as Student[]).forEach(s => { 
      if (s.ketLuan === 'Đạt') { t += (s.soTienHB || 0); c++; } 
    });
    return { t, c };
  }, [results]);

  const advancedStats = useMemo(() => {
    const allAwarded = (Object.values(results).flat() as Student[]).filter(s => s.ketLuan === 'Đạt');
    
    const byKhoaMap = new Map<string, { count: number, money: number }>();
    const byNganhMap = new Map<string, { count: number, money: number }>();

    allAwarded.forEach(s => {
      const k = s.khoa || 'Khác';
      const n = s.nganh || 'Khác';
      
      const currK = byKhoaMap.get(k) || { count: 0, money: 0 };
      byKhoaMap.set(k, { count: currK.count + 1, money: currK.money + (s.soTienHB || 0) });

      const currN = byNganhMap.get(n) || { count: 0, money: 0 };
      byNganhMap.set(n, { count: currN.count + 1, money: currN.money + (s.soTienHB || 0) });
    });

    return {
      byKhoa: Array.from(byKhoaMap.entries()).map(([name, data]) => ({ name, ...data })).sort((a,b) => b.money - a.money),
      byNganh: Array.from(byNganhMap.entries()).map(([name, data]) => ({ name, ...data })).sort((a,b) => b.money - a.money)
    };
  }, [results]);

  const TABS_LABEL = {
    import: 'CẤU HÌNH, NHẬP LIỆU & KIỂM TRA',
    results: 'KẾT QUẢ'
  };

  return (
    <div className="min-h-screen flex flex-col font-sans" style={{ backgroundColor: COLORS.background, color: COLORS.text }}>
      <header className="bg-white border-b-2 border-slate-100 shadow-sm sticky top-0 z-50 px-8 py-4">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-6">
            <img src="https://due.udn.vn/portals/_default/skins/dhkt/img/front/logoDUE.png" alt="Logo DUE" className="h-12 w-auto" />
            <div className="h-8 w-px bg-slate-200 hidden md:block"></div>
            <div>
              <h1 className="text-xl font-black uppercase tracking-tight" style={{ color: COLORS.blue }}>
                Học bổng <span style={{ color: COLORS.orange }}>Pro</span> <span style={{ color: COLORS.green }}>DUE</span>
              </h1>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">HỆ THỐNG XÉT HỌC BỔNG SIU VUÝP PRO CỦA DIỆU LINH VÀ VĨNH DIỆN</p>
            </div>
          </div>
          <nav className="hidden lg:flex items-center gap-1 bg-slate-100 p-1 rounded-xl">
            {Object.entries(TABS_LABEL).map(([key, label]) => (
              <button 
                key={key} 
                onClick={() => setActiveTab(key as any)} 
                className={`px-4 py-2 rounded-lg text-xs font-black transition-all ${activeTab === key ? 'bg-white shadow-sm' : 'text-slate-400 hover:text-slate-600'}`} 
                style={activeTab === key ? { color: COLORS.blue } : {}}
              >
                {label}
              </button>
            ))}
          </nav>
        </div>
      </header>

      {notification && (
        <div className="fixed top-24 right-6 z-50 animate-slide-in p-4 bg-white border border-slate-200 rounded-2xl shadow-2xl flex items-center gap-4">
          <div className="w-1.5 h-10 rounded-full" style={{ backgroundColor: notification.type === 'error' ? 'red' : COLORS.orange }}></div>
          <div><p className="font-black text-xs text-slate-400 uppercase">Thông báo</p><p className="font-bold text-slate-800 text-sm">{notification.message}</p></div>
        </div>
      )}

      <main className="flex-1 max-w-7xl mx-auto w-full p-8">
        {activeTab === 'import' && (
          <div className="space-y-10 animate-fade-in">
            <div className="bg-white p-10 rounded-3xl border border-slate-100 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-6">
              <div className="flex items-center gap-4">
                <div className="w-16 h-16 rounded-2xl bg-blue-50 flex items-center justify-center">
                  <Calendar size={28} style={{ color: COLORS.blue }} />
                </div>
                <div>
                  <h2 className="text-xl font-black text-slate-800">Thiết lập học kỳ</h2>
                  <p className="text-xs text-slate-500 font-bold mt-1 uppercase tracking-widest">Nhập mã học kỳ áp dụng</p>
                </div>
              </div>
              <input 
                  type="text" 
                  value={semesterCode} 
                  onChange={e => setSemesterCode(e.target.value)} 
                  className="w-full sm:w-64 px-6 py-4 bg-slate-50 border-2 rounded-xl text-center text-2xl font-black focus:border-blue-300 outline-none transition-all" 
                  style={{ color: COLORS.blue }} 
                  placeholder="Mã học kỳ"
              />
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              {[ {id: 'CLASS', label: 'Ngân sách Quỹ Chung', data: budgetClassData}, 
                 {id: 'MAJOR', label: 'Ngân sách Ngành học', data: budgetMajorData} ].map(t => (
                  <div key={t.id} className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex flex-col justify-between hover:shadow-md transition-all border-l-4 border-l-orange-500">
                    <div className="flex justify-between items-start mb-4">
                      <div>
                        <h3 className="text-sm font-black text-slate-700 uppercase">{t.label}</h3>
                        <p className="text-xs text-slate-400 mt-1">{t.data.length ? `Đã nạp ${t.data.length} dòng` : 'Chưa có dữ liệu'}</p>
                      </div>
                      {t.id === 'CLASS' ? <Layers size={18} style={{ color: COLORS.orange }} /> : <BarChart3 size={18} style={{ color: COLORS.orange }} />}
                    </div>
                    <div className="flex gap-2 w-full">
                      <button onClick={() => downloadTemplate(`BUDGET_${t.id}` as any)} className="flex-1 px-2 py-2.5 text-sm font-black border border-slate-200 rounded-xl hover:bg-slate-50 transition-colors uppercase tracking-widest text-slate-600">TẢI MẪU</button>
                      <label className="flex-1 px-2 py-2.5 text-sm font-black bg-blue-50 text-blue-600 rounded-xl cursor-pointer hover:bg-blue-100 transition-colors text-center uppercase tracking-widest border border-blue-100">
                        <input type="file" className="hidden" onChange={e => handleBudgetUpload(e, t.id as any)} />
                        TẢI LÊN
                      </label>
                    </div>
                  </div>
              ))}
              
              <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex flex-col justify-between hover:shadow-md transition-all border-l-4 border-l-emerald-500">
                <div className="flex justify-between items-start mb-4">
                  <div>
                    <h3 className="text-sm font-black text-slate-700 uppercase">Danh sách Sinh viên</h3>
                    <p className="text-xs text-slate-400 mt-1">{studentData.length ? `Đã nạp ${studentData.length} SV` : 'Chưa có dữ liệu'}</p>
                  </div>
                  <Users size={18} style={{ color: COLORS.green }} />
                </div>
                <div className="flex gap-2 w-full">
                  <button onClick={() => downloadTemplate('STUDENT')} className="flex-1 px-2 py-2.5 text-sm font-black border border-slate-200 rounded-xl hover:bg-slate-50 transition-colors uppercase tracking-widest text-slate-600">TẢI MẪU</button>
                  <label className="flex-1 px-2 py-2.5 text-sm font-black bg-emerald-500 text-white rounded-xl cursor-pointer hover:bg-emerald-600 shadow-sm transition-colors text-center uppercase tracking-widest border border-emerald-600">
                    <input type="file" className="hidden" onChange={handleStudentUpload} />
                    TẢI LÊN
                  </label>
                </div>
              </div>
            </div>
          
          {studentData.length > 0 && (
            <div className="space-y-10 animate-fade-in mt-10 border-t-2 border-slate-100 pt-10">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
              {[
                { label: 'Tổng sinh viên', val: studentData.length, icon: <Users size={20}/>, color: COLORS.blue },
                { label: 'Diện khó khăn', val: studentData.filter(x => x.hoanCanh.includes('Hộ')).length, icon: <AlertTriangle size={20}/>, color: COLORS.orange },
                { label: 'Tài năng', val: studentData.filter(x => x.hbTaiNangFlag).length, icon: <GraduationCap size={20}/>, color: COLORS.green },
                { label: 'Lưu học sinh', val: studentData.filter(x => x.quocTich !== 'Việt Nam').length, icon: <Info size={20}/>, color: COLORS.blue }
              ].map((x, i) => (
                <div key={i} className="bg-white p-8 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-5">
                  <div className="w-12 h-12 rounded-xl flex items-center justify-center" style={{ backgroundColor: x.color + '10', color: x.color }}>{x.icon}</div>
                  <div><p className="text-xs font-black text-slate-400 uppercase tracking-widest">{x.label}</p><p className="text-3xl font-black text-slate-800">{x.val}</p></div>
                </div>
              ))}
            </div>
            <div className="bg-white p-10 rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
              <div className="flex flex-col md:flex-row justify-between items-center gap-6 mb-8">
                <h3 className="font-black text-lg uppercase tracking-tight">Rà soát dữ liệu nhập vào</h3>
                <div className="relative w-full md:w-80"><Search size={16} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" /><input type="text" placeholder="Tìm tên hoặc mã sinh viên..." value={searchTerm} onChange={e => setSearchTerm(e.target.value)} className="w-full pl-12 pr-6 py-3 bg-slate-50 border rounded-xl text-xs font-bold outline-none focus:border-blue-200 transition-all" /></div>
              </div>
              <div className="max-h-[450px] overflow-y-auto pr-2 custom-scrollbar border rounded-2xl">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-50/90 backdrop-blur-sm text-slate-500 font-black uppercase text-sm tracking-widest sticky top-0 z-10 border-b">
                    <tr>
                      <th className="px-4 py-4 whitespace-nowrap">Tên sinh viên</th>
                      <th className="px-4 py-4">MSSV</th>
                      <th className="px-4 py-4">Lớp</th>
                      <th className="px-4 py-4">Khoa</th>
                      <th className="px-4 py-4">Ngành</th>
                      <th className="px-4 py-4 text-center">Hệ</th>
                      <th className="px-4 py-4 text-center">GPA</th>
                      <th className="px-4 py-4 text-center">ĐRL</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {studentData.filter(x => x.hoTen.toLowerCase().includes(searchTerm.toLowerCase()) || x.maSV.includes(searchTerm)).slice(0, 100).map(s => (
                      <tr key={s.maSV} className="hover:bg-blue-50/50 transition-all bg-white">
                        <td className="px-4 py-4 font-black text-slate-700">{s.hoTen || '-'}</td>
                        <td className="px-4 py-4 text-slate-500 font-bold">{s.maSV}</td>
                        <td className="px-4 py-4 font-bold text-slate-600">{s.lop}</td>
                        <td className="px-4 py-4 text-xs text-slate-500 max-w-[120px] truncate" title={s.khoa}>{s.khoa}</td>
                        <td className="px-4 py-4 text-xs text-slate-500 max-w-[120px] truncate" title={s.nganh}>{s.nganh}</td>
                        <td className="px-4 py-4 text-center">
                          <span className="text-xs font-black px-2 py-1 rounded text-white uppercase shadow-sm" style={{ backgroundColor: s.loaiHinhDaoTao === 'E' ? COLORS.blue : s.loaiHinhDaoTao === 'P' ? COLORS.orange : COLORS.green }}>HỆ {s.loaiHinhDaoTao}</span>
                        </td>
                        <td className="px-4 py-4 text-center font-black text-blue-600 tabular-nums">{s.diem4}</td>
                        <td className="px-4 py-4 text-center font-black text-orange-600 tabular-nums">{s.diemRenLuyen}</td>
                      </tr>
                    ))}
                    {studentData.length > 100 && (
                      <tr>
                        <td colSpan={8} className="px-6 py-4 text-center text-xs font-bold text-slate-400 italic bg-slate-50">
                          Đang hiển thị 100 sinh viên đầu tiên. Vui lòng sử dụng thanh tìm kiếm.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
              <div className="mt-8 pt-8 border-t flex justify-end">
                <button onClick={runProcessing} disabled={isProcessing} className="px-10 py-4 rounded-xl text-white font-black text-xs uppercase tracking-[0.2em] shadow-lg hover:brightness-110 active:scale-95 transition-all disabled:opacity-50" style={{ backgroundColor: COLORS.blue }}>
                    {isProcessing ? 'ĐANG XỬ LÝ...' : 'TIẾN HÀNH XÉT DUYỆT TỰ ĐỘNG'}
                </button>
              </div>
            </div>
          </div>
        )}
        </div>
      )}

        {activeTab === 'results' && (
          <div className="space-y-12 animate-fade-in pb-20">
            <div className="bg-white p-12 rounded-[2.5rem] shadow-2xl flex flex-col lg:flex-row items-center justify-between border-4 border-white gap-10 relative overflow-hidden">
              <div className="absolute top-0 right-0 w-80 h-80 opacity-5 rotate-12 -mr-16 -mt-16" style={{ color: COLORS.blue }}><GraduationCap size={320} /></div>
              <div className="relative z-10 text-center lg:text-left">
                <p className="text-sm font-black text-slate-400 uppercase tracking-[0.4em] mb-3">TỔNG GIẢI NGÂN HỌC KỲ {semesterCode}</p>
                <div className="text-5xl md:text-6xl font-black tabular-nums tracking-tighter" style={{ color: COLORS.blue }}>{formatCurrency(totalStats.t)} <span className="text-2xl opacity-30 font-bold ml-2 uppercase">VNĐ</span></div>
                <div className="flex items-center gap-4 mt-8 justify-center lg:justify-start"><span className="text-xs font-black bg-emerald-500 text-white px-5 py-2 rounded-full uppercase tracking-widest">{totalStats.c} SINH VIÊN ĐẠT HỌC BỔNG</span></div>
              </div>
              <div className="flex flex-col gap-4 relative z-10 w-full lg:w-auto">
                <button onClick={exportMasterExcel} className="bg-white border-2 px-8 py-4 rounded-xl font-black text-xs flex items-center justify-center gap-4 shadow-xl uppercase tracking-widest hover:bg-slate-50 transition-all active:scale-95" style={{ color: COLORS.blue, borderColor: COLORS.blue }}><FileSpreadsheet size={20} /> XUẤT TỔNG HỢP 5 SHEET (29 CỘT)</button>
                <div className="grid grid-cols-2 gap-4">
                  <button onClick={() => exportAdministrativeTN_QT(ScholarshipStep.TALENT, 'MAU_5_TN')} className="bg-orange-50 px-6 py-4 rounded-xl font-black text-xs text-orange-600 border border-orange-100 hover:bg-orange-100 transition-all uppercase tracking-widest">MẪU 5-TÀI NĂNG</button>
                  <button onClick={exportAdministrativeHardship} className="bg-emerald-50 px-6 py-4 rounded-xl font-black text-xs text-emerald-600 border border-emerald-100 hover:bg-emerald-100 transition-all uppercase tracking-widest">MẪU 5-VƯỢT KHÓ</button>
                  <button onClick={() => exportAdministrativeTN_QT(ScholarshipStep.INTERNATIONAL, 'MAU_4_QT')} className="bg-blue-50 px-6 py-4 rounded-xl font-black text-xs text-blue-600 border border-blue-100 hover:bg-blue-100 transition-all uppercase tracking-widest">MẪU 4-QUỐC TẾ</button>
                  <button onClick={exportAdministrativeAcademic} className="bg-slate-50 px-6 py-4 rounded-xl font-black text-xs text-slate-600 border border-slate-100 hover:bg-slate-100 transition-all uppercase tracking-widest">MẪU 4-HT&RL</button>
                </div>
              </div>
            </div>

            <div className="space-y-6">
              <h3 className="font-black text-sm uppercase tracking-widest text-slate-400">1. THỐNG KÊ THEO LOẠI HỌC BỔNG</h3>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                {(Object.keys(results) as ScholarshipStep[]).map(step => {
                  const stepAwarded = results[step].filter(x => x.ketLuan === 'Đạt');
                  const stepMoney = stepAwarded.reduce((sum, x) => sum + (x.soTienHB || 0), 0);
                  return (
                    <div key={step} className="bg-white p-8 rounded-3xl border border-slate-100 shadow-sm flex flex-col gap-4 hover:shadow-lg transition-all">
                      <div className="flex items-center gap-3"><div className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: step === ScholarshipStep.TALENT ? COLORS.orange : step === ScholarshipStep.HARDSHIP ? COLORS.green : COLORS.blue }}></div><span className="text-xs font-black uppercase text-slate-700 tracking-widest">{STEP_NAMES[step]}</span></div>
                      <div>
                        <div className="text-2xl font-black text-slate-800 tabular-nums">{formatCurrency(stepMoney)} <span className="text-xs font-bold text-slate-300">VNĐ</span></div>
                        <div className="text-xs font-bold text-slate-400 mt-1">{stepAwarded.length} sinh viên</div>
                      </div>
                    </div>
                  );
                })}
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-6">
                <div>
                  <h3 className="font-black text-sm uppercase tracking-widest text-slate-400 mb-6">2. THỐNG KÊ THEO KHOA QUẢN LÝ</h3>
                  <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-slate-400 font-black uppercase text-sm tracking-widest">
                        <tr><th className="px-6 py-4">Khoa</th><th className="px-6 py-4 text-center">Số SV</th><th className="px-6 py-4 text-right">Tổng tiền (VNĐ)</th></tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50">
                        {advancedStats.byKhoa.map(k => (
                          <tr key={k.name} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-6 py-4 font-bold text-slate-700">{k.name}</td>
                            <td className="px-6 py-4 text-center font-bold text-slate-500">{k.count}</td>
                            <td className="px-6 py-4 text-right font-black tabular-nums text-blue-600">{formatCurrency(k.money)}</td>
                          </tr>
                        ))}
                        {advancedStats.byKhoa.length === 0 && <tr><td colSpan={3} className="px-6 py-8 text-center text-slate-400 italic">Chưa có dữ liệu</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div>
                  <h3 className="font-black text-sm uppercase tracking-widest text-slate-400 mb-6">3. THỐNG KÊ THEO NGÀNH HỌC</h3>
                  <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden max-h-[400px] overflow-y-auto custom-scrollbar">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-slate-400 font-black uppercase text-sm tracking-widest sticky top-0">
                        <tr><th className="px-6 py-4">Ngành học</th><th className="px-6 py-4 text-center">Số SV</th><th className="px-6 py-4 text-right">Tổng tiền (VNĐ)</th></tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50">
                        {advancedStats.byNganh.map(n => (
                          <tr key={n.name} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-6 py-4 font-bold text-slate-700">{n.name}</td>
                            <td className="px-6 py-4 text-center font-bold text-slate-500">{n.count}</td>
                            <td className="px-6 py-4 text-right font-black tabular-nums text-emerald-600">{formatCurrency(n.money)}</td>
                          </tr>
                        ))}
                        {advancedStats.byNganh.length === 0 && <tr><td colSpan={3} className="px-6 py-8 text-center text-slate-400 italic">Chưa có dữ liệu</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>

            <div className="space-y-12">
              {(Object.entries(results) as [ScholarshipStep, Student[]][]).map(([step, data]) => (
                <div key={step} className="bg-white rounded-[2rem] border border-slate-100 overflow-hidden shadow-sm">
                  <div className="px-10 py-6 border-b flex justify-between items-center bg-slate-50/50">
                    <h4 className="font-black text-xs uppercase tracking-[0.2em]" style={{ color: COLORS.blue }}>Xem nhanh: {STEP_NAMES[step]}</h4>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-slate-50/20 text-slate-400 font-black uppercase text-sm tracking-widest">
                        <tr><th className="px-10 py-5">Sinh viên & MSSV</th><th className="px-6 py-5">Hệ & Ngành</th><th className="px-6 py-5 text-right">Số tiền</th><th className="px-10 py-5 text-center">Kết luận</th><th className="px-6 py-5">Ghi chú</th></tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50">
                        {data.slice(0, 10).map(s => (
                          <tr key={s.maSV} className="hover:bg-slate-50/30 transition-all">
                            <td className="px-10 py-5"><p className="font-black text-slate-700 text-xs">{s.hoTen}</p><p className="text-sm text-slate-400 font-bold mt-1 uppercase tracking-tighter">{s.maSV}</p></td>
                            <td className="px-6 py-5"><p className="font-bold text-slate-600">Hệ {s.loaiHinhDaoTao}</p><p className="text-sm text-slate-400 italic truncate max-w-[150px]">{s.nganh}</p></td>
                            <td className="px-6 py-5 text-right font-black text-sm tabular-nums" style={{ color: COLORS.blue }}>{s.soTienHB ? formatCurrency(s.soTienHB) : '—'}</td>
                            <td className="px-10 py-5 text-center"><span className={`text-sm font-black uppercase px-4 py-1.5 rounded-full shadow-sm ${s.ketLuan === 'Đạt' ? 'bg-emerald-500 text-white' : 'bg-slate-100 text-slate-400'}`}>{s.ketLuan}</span></td>
                            <td className="px-6 py-5 italic text-slate-400 text-sm truncate max-w-[150px]">{s.ghiChu}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {data.length > 10 && <div className="p-4 bg-slate-50/50 text-center text-xs font-black text-slate-300 uppercase tracking-[0.4em] border-t italic">... Vui lòng xem đầy đủ 29 cột trong tệp báo cáo tổng hợp ...</div>}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </main>

      <footer className="py-12 bg-white border-t border-slate-100 text-center mt-auto">
        <p className="text-slate-400 text-xs font-black uppercase tracking-[0.8em] mb-4">STARS DUE &bull; TRƯỜNG ĐẠI HỌC KINH TẾ - ĐH ĐÀ NẴNG</p>
        <p className="text-slate-300 text-sm font-bold uppercase tracking-[0.2em] italic">Hệ thống quản lý học bổng KKHT phiên bản chuyên nghiệp v11.0 by Le Vinh Dien</p>
      </footer>

      <style>{`
        @keyframes slide-in { from { transform: translateX(100%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        .animate-slide-in { animation: slide-in 0.4s cubic-bezier(0.16, 1, 0.3, 1); }
        .animate-fade-in { animation: fadeIn 0.5s ease-out; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .custom-scrollbar::-webkit-scrollbar { width: 4px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
      `}</style>
    </div>
  );
};

export default App;
