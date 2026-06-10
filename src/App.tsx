import React, { useState, useEffect, useMemo, useRef } from 'react';
// ... 後續原本的程式碼
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously, signInWithCustomToken, onAuthStateChanged } from 'firebase/auth';
import { getFirestore, collection, onSnapshot, addDoc, deleteDoc, doc, serverTimestamp, setDoc, updateDoc } from 'firebase/firestore';
import { 
  MapPin, 
  Calculator, 
  Download, 
  History, 
  FileSpreadsheet, 
  Trash2, 
  Car, 
  Bike, 
  Train, 
  CheckSquare, 
  Square, 
  FileText, 
  Settings, 
  AlertCircle, 
  CheckCircle2, 
  Clock, 
  Filter, 
  ChevronLeft, 
  ChevronRight,
  Users,
  Sparkles,
  Printer,
  Lock,
  Unlock,
  ExternalLink,
  CalendarDays,
  Calendar,
  Search,
  Plus,        
  CheckCircle, 
  Check,       
  XCircle,
  FileDown
} from 'lucide-react';
import XLSX from 'xlsx-js-style';

// =========================================================================
// 1. 全局環境變數與外部擴充宣告 (徹底解決 Cannot find name 錯誤)
// =========================================================================
declare const __firebase_config: string | undefined;
declare const __app_id: string | undefined;
declare const __initial_auth_token: string | undefined;

interface Window {
  XLSX: any;
}

// =========================================================================
// 2. 資料結構與型別定義 (TypeScript 核心規格表)
// =========================================================================

// 差旅紀錄的資料結構
interface TravelRecord {
  id: string;
  name: string;
  group: string;
  title: string;
  status: string;              // 狀態，例如：'待審核'、'已核准'、'已駁回'
  date: string;                // 開始日期 (YYYY-MM-DD)
  isMultiDay?: boolean;        // 是否為多日
  endDate?: string;            // 結束日期
  days?: number;               // 天數
  destination: string;         // 目的地
  reason: string;              // 出差事由
  transportMode: string;       // 交通工具
  hsrStation?: string;         // 高鐵站
  totalDistance?: number;      // 總里程
  duration?: number;           // 時數
  transportFee: number;        // 交通費
  accommodationFee: number;    // 住宿費
  mealFee: number;             // 膳雜費
  totalFee: number;            // 總費用
}

// 假別/結餘時數結構
interface LeaveBalance {
  d: number; // 天
  h: number; // 時
}

// 同仁基本資料與特休結構
interface Employee {
  name: string;
  title: string;
  group: string;
  hireDate: string;
  remainingTe: LeaveBalance;
  remainingBu: LeaveBalance;
  takenShi: LeaveBalance;
  takenBing: LeaveBalance;
  takenSang: LeaveBalance;
}

// 下拉選單群組限定型別
type GroupType = '全部' | '行政人員' | '資訊人員' | '輔導人員' | string;


// =========================================================================
// 3. Firebase 資料庫初始化與安全連線設定 (明確指定 any 避免隱式型別錯誤)
// =========================================================================
let app: any, auth: any, db: any, appId: string;
try {
  // 1. 改用 window 物件來讀取全域變數，這樣 TypeScript 就不會報錯中斷
  const globalConfig = (window as any).__firebase_config;
  
  const firebaseConfig = globalConfig ? JSON.parse(globalConfig) : {
    apiKey: "AIzaSyDjPezJDSRZ-Vvb-XySDEX9D8iR3WSuS2I",
    authDomain: "yunlin-digital-travel.firebaseapp.com",
    projectId: "yunlin-digital-travel",
    storageBucket: "yunlin-digital-travel.firebasestorage.app",
    messagingSenderId: "696155447858",
    appId: "1:696155447858:web:f70ae90592e901ba90922b",
    measurementId: "G-S83BJ89QTV"
  };

  app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  db = getFirestore(app);
  
  // 2. 同理，改用 window 讀取 appId，若無則套用後方的預設名稱
  appId = (window as any).__app_id || 'travel-expense-app';
} catch (e) {
  console.error("Firebase 連線初始化失敗:", e);
}


// =========================================================================
// 4. 靜態設定資料 (加上型別限制，防止物件 Key 的對應報錯)
// =========================================================================

// 預設之 115 年度同仁留用與出勤結餘資料
const INITIAL_EMPLOYEES: Employee[] = [
  { name: '康芳鈞', title: '資訊人員', group: '中區小組', hireDate: '111/06/01', remainingTe: { d: 7, h: 2 }, remainingBu: { d: 0, h: 0 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '戴君櫟', title: '輔導人員', group: '中區小組', hireDate: '111/06/01', remainingTe: { d: 7, h: 0 }, remainingBu: { d: 1, h: 4 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '李佩縜', title: '輔導人員', group: '西區小組', hireDate: '111/07/01', remainingTe: { d: 6, h: 7 }, remainingBu: { d: 0, h: 7 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '王靖瑜', title: '輔導人員', group: '東區小組', hireDate: '111/07/01', remainingTe: { d: 0, h: 0 }, remainingBu: { d: 0, h: 1 }, takenShi: { d: 5, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '龔柏逢', title: '資訊人員', group: '北區小組', hireDate: '111/07/05', remainingTe: { d: 6, h: 7 }, remainingBu: { d: 3, h: 5 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '新進人員(待定)', title: '資訊人員', group: '東區小組', hireDate: '', remainingTe: { d: 0, h: 0 }, remainingBu: { d: 0, h: 0 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } }, 
  { name: '呂其樺', title: '資訊人員', group: '南區小組', hireDate: '112/08/07', remainingTe: { d: 6, h: 7 }, remainingBu: { d: 0, h: 6 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '蒲信宏', title: '資訊人員', group: '西區小組', hireDate: '114/07/01', remainingTe: { d: 4, h: 0 }, remainingBu: { d: 2, h: 7 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '曾欒閔', title: '行政人員', group: '行政', hireDate: '115/02/01', remainingTe: { d: 0, h: 0 }, remainingBu: { d: 1, h: 1 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '陳奕瑄', title: '輔導人員', group: '南區小組', hireDate: '115/02/09', remainingTe: { d: 0, h: 0 }, remainingBu: { d: 2, h: 2 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } },
  { name: '黃銘麒', title: '輔導人員', group: '北區小組', hireDate: '115/02/01', remainingTe: { d: 0, h: 0 }, remainingBu: { d: 3, h: 2 }, takenShi: { d: 0, h: 0 }, takenBing: { d: 0, h: 0 }, takenSang: { d: 0, h: 0 } }
];

// 💡 加上 Record<string, number> 解決 TS7053 的對應索引報錯
const ROLE_SORT_ORDER: Record<string, number> = {
  '行政人員': 1,
  '資訊人員': 2,
  '輔導人員': 3
};

const COMMON_LOCATIONS = [
  { name: '越港國小', address: '雲林縣土庫鎮林森路120號' },
  { name: '雲林縣政府', address: '雲林縣斗六市府前街2號' },
  { name: '土庫國小', address: '雲林縣土庫鎮光明路59號' },
  { name: '安慶國小', address: '雲林縣虎尾鎮民主路36號' },
  { name: '大同國小', address: '雲林縣莿桐鄉莿桐村大同路1號' },
  { name: '崙背國小', address: '雲林縣崙背鄉大同路10號' },
  { name: '文安國小', address: '雲林縣西螺鎮延平路411號' },
  { name: '雲林國小', address: '雲林縣斗六市莊敬路100號' }
];

// 💡 加上 Record<string, number> 解決高鐵站動態查詢時的型別錯誤
const HSR_FARES_FROM_YUNLIN: Record<string, number> = {
  '南港': 970, '台北': 930, '板橋': 900, '桃園': 780,
  '新竹': 640, '苗栗': 500, '台中': 230, '彰化': 110,
  '嘉義': 150, '台南': 420, '左營': 560
};

const HSR_FARES_NON_RESERVED: Record<string, number> = {
  '南港': 940, '台北': 900, '板橋': 870, '桃園': 755,
  '新竹': 620, '苗栗': 485, '台中': 220, '彰化': 105,
  '嘉義': 145, '台南': 405, '左營': 540
};

const BASE_ADDRESS = '雲林縣土庫鎮林森路120號';
const CURRENT_ROC_YEAR = new Date().getFullYear() - 1911;

export default function App() {
  const [user, setUser] = useState(null);
  const [activeTab, setActiveTab] = useState('attendance_dashboard'); 
  const [notification, setNotification] = useState(null);
  const [applyType, setApplyType] = useState('leave'); 

  // 請假與加班狀態
  const [employees, setEmployees] = useState(INITIAL_EMPLOYEES);
  const [requests, setRequests] = useState([]);
  const [overtimeRequests, setOvertimeRequests] = useState([]);
  const [syncedRequests, setSyncedRequests] = useState([]);

  // 主管解鎖機制
  const [inputPassword, setInputPassword] = useState('');
  const [isUnlocked, setIsUnlocked] = useState(false);
  const correctPassword = "115"; 

  // 請假篩選
  const [leaveSearch, setLeaveSearch] = useState('');
  const [leaveFilterYear, setLeaveFilterYear] = useState('全部');
  const [leaveFilterMonth, setLeaveFilterMonth] = useState('全部');
  const [leaveFilterStatus, setLeaveFilterStatus] = useState('全部');
  const [leavePage, setLeavePage] = useState(1);
  const recordsPerPage = 10;

  // 加班篩選
  const [otSearch, setOtSearch] = useState('');
  const [otFilterYear, setOtFilterYear] = useState('全部');
  const [otFilterMonth, setOtFilterMonth] = useState('全部');
  const [otFilterStatus, setOtFilterStatus] = useState('全部');
  const [otPage, setOtPage] = useState(1);

  // 請假表單狀態
  const [formApplicant, setFormApplicant] = useState('');
  const [formLeaveType, setFormLeaveType] = useState('特休');
  const [formStartDate, setFormStartDate] = useState(new Date().toISOString().split('T')[0]);
  const [formEndDate, setFormEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [formStartTime, setFormStartTime] = useState('08:30');
  const [formEndTime, setFormEndTime] = useState('17:30');
  const [formHours, setFormHours] = useState(8);
  const [formAgent, setFormAgent] = useState('');
  const [formLocation, setFormLocation] = useState('越港國小');
  const [formReason, setFormReason] = useState('');
  const [formIsBusinessTrip, setFormIsBusinessTrip] = useState(false);

  // 加班表單狀態
  const [otDate, setOtDate] = useState(new Date().toISOString().split('T')[0]);
  const [otStartTime, setOtStartTime] = useState('08:00');
  const [otEndTime, setOtEndTime] = useState('17:00');
  const [otHours, setOtHours] = useState(8);
  const [otActivityName, setOtActivityName] = useState('');
  const [otParticipants, setOtParticipants] = useState([]);

  const [showAgreementModal, setShowAgreementModal] = useState(false);
  const [selectedOtRequest, setSelectedOtRequest] = useState(null);

  // 差旅狀態
  const [records, setRecords] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [gmapsKey, setGmapsKey] = useState(() => localStorage.getItem('gmaps_api_key') || '');
  const [currentUser, setCurrentUser] = useState('主管端'); 
  const [rosterSearch, setRosterSearch] = useState('');
  const [selectedGroup, setSelectedGroup] = useState('全部');

  const [formData, setFormData] = useState({
    date: new Date().toISOString().split('T')[0],
    memberNames: [INITIAL_EMPLOYEES[0].name],
    reason: '辦理研習',
    startLocation: '越港國小',
    destination: '',
    transportMode: 'car',
    distance: '',
    isRoundTrip: true,
    hsrTicketType: 'standard', 
    hsrTicketTypeOutbound: 'standard',
    hsrTicketTypeReturn: 'standard',
    isHsrSameTicketType: true, 
    hsrStation: '台中',
    hsrStationReturn: '台中',
    isHsrSameStation: true,
    isMultiDay: false,
    endDate: new Date().toISOString().split('T')[0],
    isOutCounty: false,
    duration: 'morning', 
    accommodationFee: '', 
    claimMealFee: true, 
  });
  
  const [isDistanceManual, setIsDistanceManual] = useState(false);
  const [selectedRecordIds, setSelectedRecordIds] = useState([]); 
  const distanceInputRef = useRef(null);

  const [exportMonth, setExportMonth] = useState('all');
  const [exportStatus, setExportStatus] = useState('pending');

  const [filterMonth, setFilterMonth] = useState('all');
  const [filterName, setFilterName] = useState('all');
  const [filterStatus, setFilterStatus] = useState('pending');
  const [currentTravelPage, setCurrentTravelPage] = useState(1);
  const travelRecordsPerPage = 15;

  const [dialog, setDialog] = useState({
    isOpen: false, type: 'alert', title: '', message: '', inputValue: '', inputType: 'text', onConfirm: null,
  });

  const [deleteTarget, setDeleteTarget] = useState(null);
  const [deletePassword, setDeletePassword] = useState('');
  const correctDeletePassword = "056341014";

  // Toast / Dialog 提示系統
  const showToast = (message, type = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 4000);
  };

  const showDialog = (type, title, message, onConfirm = null, inputType = 'text') => {
    setDialog({ isOpen: true, type, title, message, inputValue: '', inputType, onConfirm });
  };
  const closeDialog = () => setDialog(p => ({ ...p, isOpen: false }));

  // 日期與時間格式化輔助函數
  const getWeekdayText = (dateStr) => {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    if (parts.length !== 3) return '';
    let year = parseInt(parts[0], 10);
    if (year < 1000) year += 1911; 
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    const date = new Date(year, month, day);
    const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
    return `星期${weekdays[date.getDay()]}`;
  };

  const formatMinguoDateText = (dateStr) => {
    if (!dateStr) return ' 年  月  日';
    const parts = dateStr.split('-');
    if (parts.length === 3) {
      let year = parseInt(parts[0], 10);
      if (year > 1900) year -= 1911; 
      return `${year} 年 ${parts[1]} 月 ${parts[2]} 日`;
    }
    return dateStr;
  };

  // 年資及特休估算 (符合您的合約累積天數與年資級距)
  const calculateSeniorityInfo = (hireDateStr) => {
    if (!hireDateStr) return { years: 0, months: 0, text: '無 (已離職/待定)', standardTe: 0, nextMilestone: '人員待定中' };
    const parts = hireDateStr.split('/');
    if (parts.length !== 3) return { years: 0, months: 0, text: '無資料', standardTe: 0, nextMilestone: '' };
    
    const hireYear = parseInt(parts[0]) + 1911;
    const hireMonth = parseInt(parts[1]) - 1;
    const hireDay = parseInt(parts[2]);
    const hireDate = new Date(hireYear, hireMonth, hireDay);
    const today = new Date(2026, 5, 9); // 當前專案鎖定民國115年6月基準點
    
    let diffMs = today.getTime() - hireDate.getTime();
    if (diffMs < 0) diffMs = 0;
    const diffDate = new Date(diffMs);
    const years = diffDate.getUTCFullYear() - 1970;
    const months = diffDate.getUTCMonth();
    const days = diffDate.getUTCDate() - 1;

    let seniorityText = '';
    if (years > 0) seniorityText += `${years} 年 `;
    if (months > 0) seniorityText += `${months} 個月 `;
    if (years === 0 && months === 0) seniorityText = `${days} 天`;

    let standardTe = 0;
    const totalMonths = years * 12 + months;
    if (totalMonths >= 6 && totalMonths < 12) standardTe = 3;
    else if (totalMonths >= 12 && totalMonths < 24) standardTe = 7;
    else if (totalMonths >= 24 && totalMonths < 36) standardTe = 10;
    else if (totalMonths >= 36 && totalMonths < 60) standardTe = 14;
    else if (totalMonths >= 60 && totalMonths < 120) standardTe = 15;
    else if (totalMonths >= 120) standardTe = Math.min(30, 15 + Math.floor(years - 10));

    let nextMilestone = '';
    if (totalMonths < 6) {
      const td = new Date(hireYear, hireMonth + 6, hireDay);
      nextMilestone = `於 ${td.getFullYear() - 1911}/${td.getMonth() + 1}/${td.getDate()} 滿半年享有 3 天特休`;
    } else if (totalMonths < 12) {
      const td = new Date(hireYear, hireMonth + 12, hireDay);
      nextMilestone = `於 ${td.getFullYear() - 1911}/${td.getMonth() + 1}/${td.getDate()} 滿 1 年享有 7 天特休`;
    } else if (totalMonths < 24) {
      nextMilestone = `滿 2 年將升至 10 天特休`;
    } else if (totalMonths < 36) {
      nextMilestone = `滿 3 年將升至 14 天特休`;
    } else {
      nextMilestone = `特休天數已累計，穩定留用中`;
    }
    return { years, months, text: seniorityText, standardTe, nextMilestone };
  };

  const chunkParticipants = (list) => {
    const chunked = [];
    if (!list) return chunked;
    for (let i = 0; i < list.length; i += 2) {
      chunked.push([list[i], list[i + 1] || '']);
    }
    return chunked;
  };

  function numberToChinese(num) {
    if (isNaN(num) || num < 0) return "";
    if (num === 0) return "零元整";
    const digits = ["零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"];
    const units = ["", "拾", "佰", "仟"];
    const bigUnits = ["", "萬", "億"];
    let numStr = Math.floor(num).toString();
    let result = "";
    let len = numStr.length;
    let groups = [];
    for (let i = len; i > 0; i -= 4) groups.push(numStr.substring(Math.max(0, i - 4), i));
    
    for (let g = 0; g < groups.length; g++) {
      let part = groups[g];
      let partResult = "";
      let partLen = part.length;
      let zeroFlag = false;
      for (let i = 0; i < partLen; i++) {
        let d = parseInt(part[i]);
        let unitIdx = partLen - 1 - i;
        if (d === 0) zeroFlag = true;
        else {
          if (zeroFlag) { partResult += digits[0]; zeroFlag = false; }
          partResult += digits[d] + units[unitIdx];
        }
      }
      if (partResult !== "") result = partResult + bigUnits[g] + result;
    }
    result = result.replace(/零+/g, '零').replace(/零萬/g, '萬').replace(/零元/g, '元');
    if (result.endsWith('零')) result = result.substring(0, result.length - 1);
    return "新台幣：" + result + "元整";
  }

  function getRocDateRange(recordsList) {
    const defaultRange = `中華民國${CURRENT_ROC_YEAR}年01月01日至${CURRENT_ROC_YEAR}年12月31日`;
    if (!recordsList || recordsList.length === 0) return defaultRange;
    const dates = [];
    recordsList.forEach(r => {
      if (r.date) dates.push(r.date);
      if (r.isMultiDay && r.endDate) dates.push(r.endDate);
    });
    dates.sort();
    if (dates.length === 0) return defaultRange;
    const minD = dates[0];
    const maxD = dates[dates.length - 1];
    const formatSingle = (dateStr) => {
      const parts = dateStr.split('-');
      const rocYear = parseInt(parts[0], 10) - 1911;
      return `${rocYear}年${parts[1]}月${parts[2]}日`;
    };
    return `中華民國${formatSingle(minD)}至${formatSingle(maxD)}`;
  }

  // 會計旅費計算模組 (完美維持 4、5、6 公里進位邏輯)
  const calculateFees = (data) => {
    let transportFee = 0;
    let mealFee = 0;
    const accomFee = Math.ceil(parseFloat(data.accommodationFee) || 0);
    let days = 1;

    if (data.isMultiDay && data.endDate && data.endDate !== data.date) {
      const start = new Date(data.date);
      const end = new Date(data.endDate);
      if (end > start) days = Math.ceil((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;
    }

    if (data.transportMode === 'car' || data.transportMode === 'scooter') {
      const dist = parseFloat(data.distance) || 0;
      if (dist >= 5) {
        const rate = data.transportMode === 'car' ? 3 : 2;
        const roundedDist = Math.ceil(dist); 
        transportFee = roundedDist * 2 * rate; 
      }
    } else if (data.transportMode === 'hsr') {
      const outboundTicketType = data.hsrTicketTypeOutbound || data.hsrTicketType || 'standard';
      const outboundFareTable = outboundTicketType === 'non-reserved' ? HSR_FARES_NON_RESERVED : HSR_FARES_FROM_YUNLIN;
      const outboundFare = outboundFareTable[data.hsrStation] || 0;

      const returnStation = data.isHsrSameStation !== false ? data.hsrStation : (data.hsrStationReturn || data.hsrStation);
      const returnTicketType = data.isHsrSameTicketType !== false ? outboundTicketType : (data.hsrTicketTypeReturn || 'standard');
      const returnFareTable = returnTicketType === 'non-reserved' ? HSR_FARES_NON_RESERVED : HSR_FARES_FROM_YUNLIN;
      const returnFare = returnFareTable[returnStation] || 0;

      transportFee = Math.ceil(outboundFare + returnFare); 
    }

    if (data.claimMealFee !== false) { 
      const fullAllowance = data.isOutCounty ? 400 : 200;
      const halfAllowance = data.isOutCounty ? 200 : 100;
      if (days === 1) {
        mealFee = data.duration === 'full' ? fullAllowance : halfAllowance;
      } else {
        const midDays = days - 2;
        const midMeal = midDays * fullAllowance;
        const edgeMeal = 2 * (data.duration === 'full' ? fullAllowance : halfAllowance);
        mealFee = midMeal + edgeMeal;
      }
    }
    
    return { transportFee, mealFee, accommodationFee: accomFee, totalFee: transportFee + mealFee + accomFee, days };
  };

  const expandRecords = (personRecords) => {
    const sortedData = [...personRecords].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    const expandedData = [];
    
    sortedData.forEach(r => {
      const rDays = Number(r.days) || 1;
      if (r.isMultiDay && rDays > 1 && r.endDate) {
        const fullAllowance = r.isOutCounty ? 400 : 200;
        const halfAllowance = r.isOutCounty ? 200 : 100;
        const edgeAllowance = r.duration === 'full' ? fullAllowance : halfAllowance;

        let outboundFare = 0;
        let returnFare = 0;
        if (r.transportMode === 'hsr') {
          const outType = r.hsrTicketTypeOutbound || r.hsrTicketType || 'standard';
          const outFareTable = outType === 'non-reserved' ? HSR_FARES_NON_RESERVED : HSR_FARES_FROM_YUNLIN;
          outboundFare = outFareTable[r.hsrStation] || 0;

          const retStation = r.isHsrSameStation !== false ? r.hsrStation : (r.hsrStationReturn || r.hsrStation);
          const retType = r.isHsrSameTicketType !== false ? outType : (r.hsrTicketTypeReturn || 'standard');
          const retFareTable = retType === 'non-reserved' ? HSR_FARES_NON_RESERVED : HSR_FARES_FROM_YUNLIN;
          returnFare = retFareTable[retStation] || 0;
        }

        for (let i = 0; i < rDays; i++) {
          const startDateObj = new Date(r.date);
          startDateObj.setDate(startDateObj.getDate() + i);
          const curDateStr = startDateObj.toISOString().split('T')[0];
          const isFirstDay = i === 0;
          const isLastDay = i === rDays - 1;

          let dayTransportFee = 0;
          if (r.transportMode === 'car' || r.transportMode === 'scooter') {
            dayTransportFee = isFirstDay ? (Number(r.transportFee) || 0) : 0;
          } else if (r.transportMode === 'hsr') {
            if (isFirstDay) dayTransportFee = outboundFare;
            if (isLastDay) dayTransportFee = returnFare;
          }

          const dayAccommodationFee = isFirstDay ? (Number(r.accommodationFee) || 0) : 0;
          const dayMealFee = (isFirstDay || isLastDay) ? (r.claimMealFee !== false ? edgeAllowance : 0) : (r.claimMealFee !== false ? fullAllowance : 0);
          
          let dayHalf = 0; let dayFull = 0;
          if (isFirstDay || isLastDay) {
            if (r.duration === 'full') dayFull = 1; else dayHalf = 0.5;
          } else {
            dayFull = 1;
          }

          let dayRemarks = '';
          if (r.transportMode === 'car') {
            if (isFirstDay) dayRemarks = '汽車(去回程合併申報)';
          } else if (r.transportMode === 'scooter') {
            if (isFirstDay) dayRemarks = '機車(去回程合併申報)';
          } else if (r.transportMode === 'hsr') {
            if (isFirstDay) {
              const outType = r.hsrTicketTypeOutbound || r.hsrTicketType || 'standard';
              const outTypeTxt = outType === 'non-reserved' ? '自由座' : '全票';
              dayRemarks = `高鐵去程:${r.hsrStation}(${outTypeTxt})`;
            }
            if (isLastDay) {
              const retStation = r.isHsrSameStation !== false ? r.hsrStation : (r.hsrStationReturn || r.hsrStation);
              const outType = r.hsrTicketTypeOutbound || r.hsrTicketType || 'standard';
              const retType = r.isHsrSameTicketType !== false ? outType : (r.hsrTicketTypeReturn || 'standard');
              const retTypeTxt = retType === 'non-reserved' ? '自由座' : '全票';
              dayRemarks = `高鐵回程:${retStation}(${retTypeTxt})`;
            }
          }
          if (r.accommodationFee && isFirstDay) dayRemarks += ` [住宿費:NT$${r.accommodationFee}]`;
          if (rDays > 1) dayRemarks += ` (多天出差 第${i + 1}/${rDays}天)`;

          expandedData.push({
            ...r, date: curDateStr, transportFee: dayTransportFee, accommodationFee: dayAccommodationFee,
            mealFee: dayMealFee, totalFee: dayTransportFee + dayAccommodationFee + dayMealFee,
            excelHalf: dayHalf, excelFull: dayFull, remarks: dayRemarks
          });
        }
      } else {
        const isHalf = (r.duration === 'morning' || r.duration === 'afternoon' || r.duration === 'half');
        let remarks = '';
        if (r.transportMode === 'car') remarks = '汽車(來回)';
        else if (r.transportMode === 'scooter') remarks = '機車(來回)';
        else if (r.transportMode === 'hsr') {
          const returnStation = r.isHsrSameStation !== false ? r.hsrStation : (r.hsrStationReturn || r.hsrStation);
          const sameStationText = r.isHsrSameStation !== false ? '來回' : `去回不同站(${r.hsrStation}→${returnStation})`;
          const outType = r.hsrTicketTypeOutbound || r.hsrTicketType || 'standard';
          const retType = r.isHsrSameTicketType !== false ? outType : (r.hsrTicketTypeReturn || 'standard');
          const outTypeText = outType === 'non-reserved' ? '自由座' : '全票';
          const retTypeText = retType === 'non-reserved' ? '自由座' : '全票';
          let ticketTypeText = (r.isHsrSameTicketType !== false && outType === retType) ? (outType === 'non-reserved' ? '自由座' : '全票') : `去:${outTypeText}/回:${retTypeText}`;
          remarks = `高鐵${ticketTypeText}(${sameStationText})`;
        }
        expandedData.push({ ...r, excelHalf: isHalf ? 0.5 : 0, excelFull: isHalf ? 0 : 1, remarks });
      }
    });
    return expandedData;
  };

  // Google Calendar 串接 API 服務產生器
  const getGoogleCalendarUrl = (req) => {
    const convertToWestern = (dateStr) => {
      if (!dateStr) return '';
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        let year = parseInt(parts[0], 10);
        if (year < 1000) year += 1911; 
        return `${year}-${parts[1]}-${parts[2]}`;
      }
      return dateStr;
    };
    const westernStartDate = convertToWestern(req.startDate || req.date);
    const westernEndDate = convertToWestern(req.endDate || req.date);
    const formattedStartDate = westernStartDate.replace(/-/g, '');
    const formattedEndDate = westernEndDate.replace(/-/g, '');
    const startHour = req.startTime ? req.startTime.replace(/:/g, '') : '0830';
    const endHour = req.endTime ? req.endTime.replace(/:/g, '') : '1730';
    const startStr = `${formattedStartDate}T${startHour}00`;
    const endStr = `${formattedEndDate}T${endHour}00`;
    const title = `【請假-${req.leaveType || '出差'}】${req.applicant || req.name} (代理人: ${req.agent || '無'})`;
    let details = `請假同仁：${req.applicant || req.name}\n請假類別：${req.leaveType || '出差公出'}\n時數：${req.hours || 8} 小時\n事由：${req.reason || '未填寫'}`;
    if (req.leaveType === '公假') details += `\n起訖地點：${req.location || '越港國小'}\n出差旅費核銷：${req.isBusinessTrip ? '可請領差旅費' : '不具出差性質'}`;
    const location = req.location || req.destination || '越港國小';
    return `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(title)}&dates=${startStr}/${endStr}&details=${encodeURIComponent(details)}&location=${encodeURIComponent(location)}`;
  };

  const triggerCalendarSync = (req, url) => {
    window.open(url, '_blank');
    setSyncedRequests(p => [...p, req.id]);
    showToast('已安全開啟 Google 行事曆新增頁面！');
  };

  // 取得與試算路徑距離
  const getTravelDistance = async (start, dest) => {
    if (!gmapsKey) return null;
    const startAddr = COMMON_LOCATIONS.find(l => l.name === start)?.address || start;
    const destAddr = COMMON_LOCATIONS.find(l => l.name === dest)?.address || dest;
    try {
      const response = await fetch('https://routes.googleapis.com/v1/computeRoutes', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Goog-Api-Key': gmapsKey,
          'X-Goog-FieldMask': 'routes.distanceMeters'
        },
        body: JSON.stringify({
          origin: { address: startAddr },
          destination: { address: destAddr },
          travelMode: 'DRIVE',
          routeModifiers: { avoidTolls: true, avoidHighways: false }
        })
      });
      const data = await response.json();
      if (data.routes && data.routes[0]) return Math.round((data.routes[0].distanceMeters / 1000) * 10) / 10;
    } catch (e) {
      console.error("試算路徑失敗:", e);
    }
    return null;
  };

  // ==========================================
  // 3. Firebase 認證與即時數據流監聽
  // ==========================================
  useEffect(() => {
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) await signInWithCustomToken(auth, __initial_auth_token);
        else await signInAnonymously(auth);
      } catch (err) {
        console.error('驗證失敗:', err);
      }
    };
    initAuth();
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => setUser(currentUser));
    return () => unsubscribe();
  }, []);

  // 差旅紀錄訂閱
  useEffect(() => {
    if (!user || !db) return;
    const recordsRef = collection(db, 'artifacts', appId, 'public', 'data', 'travel_expenses');
    const unsubscribe = onSnapshot(recordsRef, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      data.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
      setRecords(data);
    }, (error) => showDialog('alert', '連線異常', '無法取得最新差旅資料，請檢查網路。'));
    return () => unsubscribe();
  }, [user]);

  // 同仁資料訂閱
  useEffect(() => {
    if (!user || !db) return;
    const empCol = collection(db, 'artifacts', appId, 'public', 'data', 'employees');
    const unsub = onSnapshot(empCol, (snapshot) => {
      if (snapshot.empty) {
        INITIAL_EMPLOYEES.forEach(async (emp) => {
          await setDoc(doc(db, 'artifacts', appId, 'public', 'data', 'employees', emp.name), emp);
        });
      } else {
        const list = snapshot.docs.map(doc => doc.data() as any);
        setEmployees(list);
      }
    });
    return () => unsub();
  }, [user]);

  // 請假紀錄訂閱
  useEffect(() => {
    if (!user || !db) return;
    const leavesCol = collection(db, 'artifacts', appId, 'public', 'data', 'leaves');
    const unsub = onSnapshot(leavesCol, (snapshot) => {
      const list = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
      setRequests(list);
    });
    return () => unsub();
  }, [user]);

  // 加班紀錄訂閱
  useEffect(() => {
    if (!user || !db) return;
    const overtimesCol = collection(db, 'artifacts', appId, 'public', 'data', 'overtimes');
    const unsub = onSnapshot(overtimesCol, (snapshot) => {
      const list = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
      setOvertimeRequests(list);
    });
    return () => unsub();
  }, [user]);

  useEffect(() => {
    if (currentUser !== '主管端') {
      setIsUnlocked(false);
      setInputPassword('');
    }
  }, [currentUser]);

  // ==========================================
  // 4. 數據過濾與分頁運算 (useMemo Hooks)
  // ==========================================
  const leaveYears = useMemo(() => {
    const set = new Set<string>();
    requests.forEach((r: any) => { if (r.startDate) set.add(r.startDate.split('-')[0]); });
    return ['全部', ...Array.from(set).sort()];
  }, [requests]);

  const leaveMonths = useMemo(() => {
    const set = new Set<string>();
    requests.forEach((r: any) => { if (r.startDate) set.add(r.startDate.split('-')[1]); });
    return ['全部', ...Array.from(set).sort()];
  }, [requests]);

  const filteredLeaves = useMemo(() => {
    return requests.filter((r: any) => {
      const matchSearch = r.applicant.includes(leaveSearch) || (r.reason && r.reason.includes(leaveSearch));
      const matchStatus = leaveFilterStatus === '全部' || r.status === leaveFilterStatus;
      const parts = r.startDate ? r.startDate.split('-') : [];
      const matchYear = leaveFilterYear === '全部' || parts[0] === leaveFilterYear;
      const matchMonth = leaveFilterMonth === '全部' || parts[1] === leaveFilterMonth;
      return matchSearch && matchStatus && matchYear && matchMonth;
    }).sort((a: any, b: any) => b.startDate.localeCompare(a.startDate));
  }, [requests, leaveSearch, leaveFilterStatus, leaveFilterYear, leaveFilterMonth]);

  const paginatedLeaves = useMemo(() => {
    const offset = (leavePage - 1) * recordsPerPage;
    return filteredLeaves.slice(offset, offset + recordsPerPage);
  }, [filteredLeaves, leavePage]);

  const otYears = useMemo(() => {
    const set = new Set<string>();
    overtimeRequests.forEach((r: any) => { if (r.workDate) set.add(r.workDate.split('-')[0]); });
    return ['全部', ...Array.from(set).sort()];
  }, [overtimeRequests]);

  const otMonths = useMemo(() => {
    const set = new Set<string>();
    overtimeRequests.forEach((r: any) => { if (r.workDate) set.add(r.workDate.split('-')[1]); });
    return ['全部', ...Array.from(set).sort()];
  }, [overtimeRequests]);

  const filteredOvertimes = useMemo(() => {
    return overtimeRequests.filter((r: any) => {
      const matchSearch = r.activityName.includes(otSearch) || r.participants.join('、').includes(otSearch);
      const matchStatus = otFilterStatus === '全部' || r.status === otFilterStatus;
      const parts = r.workDate ? r.workDate.split('-') : [];
      const matchYear = otFilterYear === '全部' || parts[0] === otFilterYear;
      const matchMonth = otFilterMonth === '全部' || parts[1] === otFilterMonth;
      return matchSearch && matchStatus && matchYear && matchMonth;
    }).sort((a: any, b: any) => b.workDate.localeCompare(a.workDate));
  }, [overtimeRequests, otSearch, otFilterStatus, otFilterYear, otFilterMonth]);

  const paginatedOvertimes = useMemo(() => {
    const offset = (otPage - 1) * recordsPerPage;
    return filteredOvertimes.slice(offset, offset + recordsPerPage);
  }, [filteredOvertimes, otPage]);

  const sortedAndFilteredEmployees = useMemo(() => {
    return employees
      .filter(emp => {
        const matchesSearch = emp.name.includes(rosterSearch) || emp.group.includes(rosterSearch) || emp.title.includes(rosterSearch);
        const matchesGroup = selectedGroup === '全部' || emp.group === selectedGroup;
        return matchesSearch && matchesGroup;
      })
      .sort((a, b) => (ROLE_SORT_ORDER[a.title] || 99) - (ROLE_SORT_ORDER[b.title] || 99));
  }, [employees, rosterSearch, selectedGroup]);

  const availableMonths = useMemo(() => {
    const months = new Set<string>();
    records.forEach(r => { if (r.date) months.add(r.date.substring(0, 7)); });
    return Array.from(months).sort().reverse();
  }, [records]);

  useEffect(() => { setCurrentTravelPage(1); }, [filterMonth, filterName, filterStatus]);

  const filteredRecords = useMemo(() => {
    return records.filter(r => {
      const rStatus = r.status || 'pending';
      const matchMonth = filterMonth === 'all' || (r.date && r.date.startsWith(filterMonth));
      const matchName = filterName === 'all' || r.name === filterName;
      const matchStatus = filterStatus === 'all' || rStatus === filterStatus;
      return matchMonth && matchName && matchStatus;
    });
  }, [records, filterMonth, filterName, filterStatus]);

  const totalTravelPages = Math.ceil(filteredRecords.length / travelRecordsPerPage) || 1;
  const currentRecords = filteredRecords.slice((currentTravelPage - 1) * travelRecordsPerPage, currentTravelPage * travelRecordsPerPage);

  const { transportFee, mealFee, accommodationFee: currentAccomFee, totalFee } = useMemo(() => calculateFees(formData), [formData]);

  // 估算加班時數
  useEffect(() => {
    if (otStartTime && otEndTime) {
      const [startH, startM] = otStartTime.split(':').map(Number);
      const [endH, endM] = otEndTime.split(':').map(Number);
      let diffHrs = (endH + endM / 60) - (startH + startM / 60);
      if (diffHrs < 0) diffHrs += 24; 
      if (otStartTime <= '12:00' && otEndTime >= '13:00') {
        diffHrs = Math.max(1, diffHrs - 1); 
      }
      setOtHours(Math.round(diffHrs * 10) / 10);
    }
  }, [otStartTime, otEndTime]);

  // 出差地點及高鐵跨縣市判斷
  useEffect(() => {
    const dest = formData.destination;
    if (dest) {
      const outCountyKeywords = ['台北', '新北', '桃園', '台中', '台南', '高雄', '基隆', '新竹', '苗栗', '彰化', '南投', '嘉義', '屏東', '宜蘭', '花蓮', '台東', '澎湖', '金門', '馬祖'];
      const isYunlin = dest.includes('雲林');
      const hasOutCountyKeyword = outCountyKeywords.some(keyword => dest.includes(keyword));
      if (hasOutCountyKeyword && !isYunlin) setFormData(prev => ({ ...prev, isOutCounty: true }));
      else setFormData(prev => ({ ...prev, isOutCounty: false }));
    }
  }, [formData.destination]);

  useEffect(() => {
    if (formData.transportMode === 'hsr') setFormData(prev => ({ ...prev, isOutCounty: true }));
  }, [formData.transportMode, formData.hsrStation]);



  // 差旅表單輸入處理
  const handleInputChange = (e) => {
    const { name, value, type, checked } = e.target;
    setFormData(prev => {
      const updated = { ...prev, [name]: type === 'checkbox' ? checked : value };
      if (name === 'hsrStation' && prev.isHsrSameStation) updated.hsrStationReturn = value;
      if (name === 'hsrTicketTypeOutbound' && prev.isHsrSameTicketType) {
        updated.hsrTicketTypeReturn = value;
        updated.hsrTicketType = value;
      }
      return updated;
    });
  };

  const toggleMemberSelection = (name) => {
    setFormData(prev => {
      const isSelected = prev.memberNames.includes(name);
      return { ...prev, memberNames: isSelected ? prev.memberNames.filter(n => n !== name) : [...prev.memberNames, name] };
    });
  };

  const toggleOtParticipant = (name) => {
    if (otParticipants.includes(name)) setOtParticipants(otParticipants.filter(p => p !== name));
    else setOtParticipants([...otParticipants, name]);
  };

  // Google Maps / OSM 導航距離抓取
  const handleAutoFetchDistance = async () => {
    if (!formData.destination.trim()) return showDialog('alert', '提示', '請先填寫出差地點！\n建議輸入明確名稱或地址。');
    setIsProcessing(true);
    try {
      let finalDistance = 0;
      let usedGoogle = false;
      let gmapError = '';

      if (gmapsKey) {
        try {
          const isScooter = formData.transportMode === 'scooter';
          const reqBody = {
            origin: { address: formData.startLocation || BASE_ADDRESS },
            destination: { address: '台灣 ' + formData.destination },
            travelMode: 'DRIVE',
            routingPreference: 'TRAFFIC_AWARE_OPTIMAL', 
            computeAlternativeRoutes: true,
            routeModifiers: { avoidHighways: isScooter, avoidTolls: isScooter }
          };
          const res = await fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'X-Goog-Api-Key': gmapsKey, 'X-Goog-FieldMask': 'routes.distanceMeters,routes.duration' },
            body: JSON.stringify(reqBody)
          });
          const data = await res.json();
          if (res.ok && data.routes && data.routes.length > 0) {
            const routesDetails = data.routes.map(r => ({ distance: r.distanceMeters || 0, durationSecs: parseInt(r.duration || '0', 10) }));
            const minDuration = Math.min(...routesDetails.map(r => r.durationSecs));
            const reasonableRoutes = routesDetails.filter(r => r.durationSecs <= minDuration * 1.2);
            const maxMeters = Math.max(...reasonableRoutes.map(r => r.distance));
            finalDistance = maxMeters / 1000;
            usedGoogle = true;
          } else {
            gmapError = data.error?.message || 'Google API 發生錯誤';
            throw new Error(gmapError);
          }
        } catch (err) {
          console.warn('Google API 呼叫失敗，將啟用 OSRM 備援圖資:', err);
        }
      }

      if (!usedGoogle) {
        const geoRes = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent('台灣 ' + formData.destination)}`);
        const geoData = await geoRes.json();
        if (!geoData || geoData.length === 0) throw new Error('地圖伺服器無法辨識此地點，已為您解鎖請直接手動輸入里程。');
        
        const destLat = geoData[0].lat;
        const destLon = geoData[0].lon;
        const baseLat = 23.6826;
        const baseLon = 120.3969; 
        
        const routeRes = await fetch(`https://router.project-osrm.org/route/v1/driving/${baseLon},${baseLat};${destLon},${destLat}?alternatives=true&overview=false`);
        const routeData = await routeRes.json();
        
        if (routeData.code !== 'Ok' || !routeData.routes || routeData.routes.length === 0) throw new Error('無法計算路線里程，已解除鎖定請手動輸入。');
        
        finalDistance = Math.max(...routeData.routes.map(r => r.distance)) / 1000;
        if (gmapsKey) showDialog('alert', '系統提示', `Google API 連線發生問題 (${gmapError})。\n已自動啟動備援 OpenStreetMap 系統計算完畢。`);
        else showDialog('alert', '系統提示', '目前使用備用圖資計算完成，為求精確仍建議綁定學校專屬 Google API。');
      }

      setFormData(prev => ({ ...prev, distance: finalDistance.toFixed(1) }));
      setIsDistanceManual(false); 

    } catch (error) {
      showDialog('alert', '抓取失敗', error.message);
      setIsDistanceManual(true);
      setTimeout(() => { if (distanceInputRef.current) distanceInputRef.current.focus(); }, 300);
    } finally {
      setIsProcessing(false);
    }
  };

  // 請假送出
  const handleApplyLeave = async (e) => {
    e.preventDefault();
    if (!user || !db) return;
    if (!formApplicant) return showDialog('alert', '提示', '請選擇請假人！');
    if (!formAgent) return showDialog('alert', '提示', '請設定代理人！');
    if (formHours <= 0) return showDialog('alert', '提示', '請假時數不正確！');

    try {
      const leavesCol = collection(db, 'artifacts', appId, 'public', 'data', 'leaves');
      await addDoc(leavesCol, {
        applicant: formApplicant,
        leaveType: formLeaveType,
        startDate: formStartDate,
        endDate: formEndDate,
        startTime: formStartTime,
        endTime: formEndTime,
        hours: formHours,
        agent: formAgent,
        location: formLocation,
        reason: formReason,
        isBusinessTrip: formLeaveType === '公假' ? formIsBusinessTrip : false,
        status: '待審核',
        createdAt: serverTimestamp()
      });
      showDialog('alert', '送出成功', '請假申請已送出，請靜候主管簽核！');
      // 重設表單
      setFormReason('');
    } catch (err) {
      showDialog('alert', '錯誤', '送出失敗！');
    }
  };

  // 加班申報送出
  const handleApplyOvertime = async (e) => {
    e.preventDefault();
    if (!user || !db) return;
    if (otParticipants.length === 0) return showDialog('alert', '提示', '請至少勾選一位加班同仁！');
    if (!otActivityName.trim()) return showDialog('alert', '提示', '請填寫活動研習名稱！');

    try {
      const overtimesCol = collection(db, 'artifacts', appId, 'public', 'data', 'overtimes');
      await addDoc(overtimesCol, {
        workDate: otDate,
        startTime: otStartTime,
        endTime: otEndTime,
        hours: otHours,
        participants: otParticipants,
        activityName: otActivityName,
        status: '待審核',
        appliedAt: new Date().toLocaleDateString('zh-TW')
      });
      showDialog('alert', '申報成功', '團體加班補休協商單已送出，待主管核定後，時數將自動加總補休。');
      setOtParticipants([]);
      setOtActivityName('');
    } catch (err) {
      showDialog('alert', '錯誤', '申報儲存失敗。');
    }
  };

  // 差旅儲存
  const handleCalculateTravel = async (e) => {
    if (e) e.preventDefault();
    if (formData.memberNames.length === 0) return showDialog('alert', '錯誤', '請至少選擇一位出差同仁。');
    if (!formData.destination.trim()) return showDialog('alert', '錯誤', '請選擇或填寫目的地。');
    
    setIsProcessing(true);
    let finalDist = parseFloat(formData.distance);
    if (isNaN(finalDist) && formData.transportMode !== 'hsr') {
      const apiDist = await getTravelDistance(formData.startLocation, formData.destination);
      if (apiDist !== null) {
        finalDist = apiDist;
        setFormData(prev => ({ ...prev, distance: apiDist.toString() }));
      } else {
        showDialog('alert', '錯誤', '無法自動試算路徑，請手動輸入公里數。');
        setIsProcessing(false);
        return;
      }
    }

    try {
      const fees = calculateFees(formData);
      const recordsRef = collection(db, 'artifacts', appId, 'public', 'data', 'travel_expenses');
      
      await Promise.all(formData.memberNames.map(async (memberName) => {
        const memberInfo = employees.find(m => m.name === memberName);
        const recordData = {
          ...formData, name: memberInfo.name, title: memberInfo.title, group: memberInfo.group,
          ...fees, distance: finalDist || 0, totalDistance: (finalDist || 0) * (formData.isRoundTrip ? 2 : 1),
          createdAt: serverTimestamp(), userId: user.uid, status: 'pending'
        };
        delete recordData.memberNames; 
        await addDoc(recordsRef, recordData);
      }));
      
      showDialog('alert', '成功', `已成功新增 ${formData.memberNames.length} 筆差旅紀錄！`);
      setFormData({
        date: new Date().toISOString().split('T')[0],
        memberNames: [INITIAL_EMPLOYEES[0].name],
        reason: '辦理研習',
        startLocation: '越港國小',
        destination: '',
        transportMode: 'car',
        distance: '',
        isRoundTrip: true,
        hsrTicketType: 'standard', 
        hsrTicketTypeOutbound: 'standard',
        hsrTicketTypeReturn: 'standard',
        isHsrSameTicketType: true, 
        hsrStation: '台中',
        hsrStationReturn: '台中',
        isHsrSameStation: true,
        isMultiDay: false,
        endDate: new Date().toISOString().split('T')[0],
        isOutCounty: false,
        duration: 'morning', 
        accommodationFee: '', 
        claimMealFee: true, 
      });
      setIsDistanceManual(false);
      setActiveTab('history'); 
    } catch (error) {
      console.error(error);
      showDialog('alert', '錯誤', '新增失敗，請重試。');
    } finally {
      setIsProcessing(false);
    }
  };

  // 匯出 CSV (差勤總表)
  const handleExportCSV = () => {
    let csvContent = "\uFEFF"; 
    csvContent += "姓名,職稱,組別,到職日,剩餘特休天數,剩餘特休時數,剩餘補休天數,剩餘補休時數,已請事假天數,已請事假時數,已請病假天數,已請病假時數,已請喪假天數,已請喪假時數\r\n";
    const sortedEmployees = [...employees].sort((a, b) => (ROLE_SORT_ORDER[a.title] || 99) - (ROLE_SORT_ORDER[b.title] || 99));
    sortedEmployees.forEach(emp => {
      csvContent += `${emp.name},${emp.title},${emp.group},${emp.hireDate || '無'},${emp.remainingTe.d},${emp.remainingTe.h},${emp.remainingBu.d},${emp.remainingBu.h},${emp.takenShi.d},${emp.takenShi.h},${emp.takenBing.d},${emp.takenBing.h},${emp.takenSang.d},${emp.takenSang.h}\r\n`;
    });
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "雲林縣數位學習推辦_115年度最新差勤總表.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('已成功匯出差勤 CSV！', 'success');
  };

  const handleExportOvertimeExcel = () => {
    let csvContent = "\uFEFF"; 
    csvContent += "加班日期,星期,活動研習名稱/加班事由,開始時間,結束時間,申報時數,參與同仁名冊,審核狀態,申請時間\r\n";
    const sortedOvertimes = [...overtimeRequests].sort((a: any, b: any) => b.workDate.localeCompare(a.workDate));
    sortedOvertimes.forEach((ot: any) => {
      const weekday = getWeekdayText(ot.workDate);
      const listStr = ot.participants ? ot.participants.join('、') : '';
      csvContent += `"${ot.workDate}","${weekday}","${ot.activityName || ''}","${ot.startTime}","${ot.endTime}","${ot.hours}小時","${listStr}","${ot.status}","${ot.appliedAt || ''}"\r\n`;
    });
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "雲林縣數位學習推辦_115年度假日加班補休清冊.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('假日加班協商歷史清冊已匯出！', 'success');
  };

  // 100% 標楷體 Word 簽呈格式 - 差假請示單匯出
  const handleExportTravelWord = () => {
    if (selectedRecordIds.length === 0) {
      return showDialog('alert', '提示', '請先在列表中勾選要匯出的紀錄');
    }
    
    try {
      const selectedData = records.filter(r => selectedRecordIds.includes(r.id));
      if (selectedData.length === 0) throw new Error("找不到勾選的資料");
      
      const uniqueMembers = [];
      selectedData.forEach(r => {
        if (!uniqueMembers.find(m => m.name === r.name)) {
          uniqueMembers.push({ name: String(r.name || '未知'), title: String(r.title || '人員') });
        }
      });

      const titlePriority = { '行政人員': 1, '資訊人員': 2, '輔導人員': 3 };
      uniqueMembers.sort((a, b) => (titlePriority[a.title] || 99) - (titlePriority[b.title] || 99));

      const todayStr = `${CURRENT_ROC_YEAR}年 ${String(new Date().getMonth() + 1).padStart(2, '0')}月 ${String(new Date().getDate()).padStart(2, '0')}日`;
      
      let personsHtml = '';
      uniqueMembers.forEach((m, idx) => {
        if (idx === 0) {
          personsHtml += `
            <tr>
              <td style="border: 1px solid black; padding: 4px; font-weight:bold;">服務單位</td>
              <td colspan="4" style="border: 1px solid black; padding: 4px;">數位學習辦公室</td>
              <td style="border: 1px solid black; padding: 4px; font-weight:bold;">職別</td>
              <td colspan="2" style="border: 1px solid black; padding: 4px;">${m.title}</td>
              <td style="border: 1px solid black; padding: 4px; font-weight:bold;">姓名</td>
              <td colspan="2" style="border: 1px solid black; padding: 4px;">${m.name}</td>
            </tr>
          `;
        } else {
          personsHtml += `
            <tr>
              <td colspan="5" style="border: 1px solid black; padding: 4px;"></td>
              <td style="border: 1px solid black; padding: 4px; font-weight:bold;">職別</td>
              <td colspan="2" style="border: 1px solid black; padding: 4px;">${m.title}</td>
              <td style="border: 1px solid black; padding: 4px; font-weight:bold;">姓名</td>
              <td colspan="2" style="border: 1px solid black; padding: 4px;">${m.name}</td>
            </tr>
          `;
        }
      });

      const groupedEvents = {};
      selectedData.forEach(r => {
        const safeDate = String(r.date || '無日期');
        const safeReason = String(r.reason || '公出');
        const safeDest = String(r.destination || '未填寫地點');
        const key = `${safeDate}_${safeDest}_${safeReason}`;
        
        if (!groupedEvents[key]) {
          groupedEvents[key] = {
            date: safeDate, duration: r.duration, reason: safeReason, destination: safeDest,
            transportFee: Number(r.transportFee) || 0, 
            accommodationFee: Number(r.accommodationFee) || 0,
            mealFee: Number(r.mealFee) || 0, 
            totalFee: Number(r.totalFee) || 0,
            days: r.duration === 'full' ? 1 : 0.5
          };
        }
      });

      let totalDays = 0;
      let totalAmount = 0;
      let rowsHtml = '';
      const groupKeys = Object.keys(groupedEvents);

      groupKeys.forEach(key => {
        const g = groupedEvents[key];
        totalDays += g.days;
        totalAmount += g.totalFee;
        let isMorning = (g.duration === 'morning' || g.duration === 'half') ? 'ˇ' : '';
        let isAfternoon = g.duration === 'afternoon' ? 'ˇ' : '';
        let isFull = g.duration === 'full' ? 'ˇ' : '';
        const dateStr = String(g.date || '');
        const shortDate = dateStr.length >= 5 ? dateStr.substring(5).replace('-', '/') : dateStr;

        rowsHtml += `
          <tr>
            <td style="border: 1px solid black; padding: 6px;">${shortDate}</td>
            <td style="border: 1px solid black; padding: 6px;">${isMorning}</td>
            <td style="border: 1px solid black; padding: 6px;">${isAfternoon}</td>
            <td style="border: 1px solid black; padding: 6px;">${isFull}</td>
            <td colspan="2" style="border: 1px solid black; padding: 6px; text-align: left;">${g.reason}</td>
            <td style="border: 1px solid black; padding: 6px; text-align: left;">${g.destination}</td>
            <td style="border: 1px solid black; padding: 6px;">${g.transportFee}</td>
            <td style="border: 1px solid black; padding: 6px;">${g.accommodationFee || ''}</td>
            <td style="border: 1px solid black; padding: 6px;">${g.mealFee}</td>
            <td style="border: 1px solid black; padding: 6px; font-weight: bold;">${g.totalFee}</td>
          </tr>
        `;
      });

      // 補足空白列至6列
      for (let i = groupKeys.length; i < 6; i++) {
         rowsHtml += `<tr><td style="border: 1px solid black; padding: 6px;">&nbsp;</td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td><td colspan="2" style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td></tr>`;
      }

      const htmlContent = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <meta charset='utf-8'>
          <title>差假請示單</title>
          <style>
            @page WordSection1 { size: 21cm 29.7cm; margin: 1.5cm 1.5cm 1.5cm 1.5cm; mso-header-margin: 35.4pt; mso-footer-margin: 35.4pt; mso-paper-source: 0; }
            div.WordSection1 { page: WordSection1; }
            body { font-family: '標楷體', 'DFKai-SB', serif; font-size: 11pt; }
            table { border-collapse: collapse; width: 100%; text-align: center; }
            td, th { border: 1px solid black; }
          </style>
        </head>
        <body>
          <div class="WordSection1">
          <h2 align="center">雲林縣數位學習專案辦公室公(差)假請示單</h2>
          <p>請示日期：民國 ${todayStr}</p>
          <table border="1" cellspacing="0" cellpadding="5" width="100%">
             ${personsHtml}
             <tr>
               <td rowspan="2" style="border: 1px solid black; padding: 6px; font-weight:bold;">日期</td>
               <td colspan="3" style="border: 1px solid black; padding: 6px; font-weight:bold;">請假時間</td>
               <td rowspan="2" colspan="2" style="border: 1px solid black; padding: 6px; font-weight:bold;">事由(學校)</td>
               <td rowspan="2" style="border: 1px solid black; padding: 6px; font-weight:bold;">地點</td>
               <td rowspan="2" style="border: 1px solid black; padding: 6px; font-weight:bold;">交通費</td>
               <td rowspan="2" style="border: 1px solid black; padding: 6px; font-weight:bold;">住宿費</td>
               <td rowspan="2" style="border: 1px solid black; padding: 6px; font-weight:bold;">膳雜費</td>
               <td rowspan="2" style="border: 1px solid black; padding: 6px; font-weight:bold;">合計</td>
             </tr>
             <tr>
               <td style="border: 1px solid black; padding: 6px;">上午</td><td style="border: 1px solid black; padding: 6px;">下午</td><td style="border: 1px solid black; padding: 6px;">整天</td>
             </tr>
             ${rowsHtml}
             <tr>
               <td colspan="5" align="left" style="border: 1px solid black; padding: 6px; font-weight:bold;">差假共計: ${totalDays} 日</td>
               <td colspan="6" align="left" style="border: 1px solid black; padding: 6px; font-weight:bold;">金額共計新台幣: ${totalAmount} 元整</td>
             </tr>
             <tr style="height: 75px;" valign="top">
               <td colspan="2" align="left" style="border: 1px solid black; padding: 6px; font-weight:bold;">差假人<br>簽章</td>
               <td colspan="2" align="left" style="border: 1px solid black; padding: 6px; font-weight:bold;">職務代理人<br>簽章</td>
               <td colspan="3" align="left" style="border: 1px solid black; padding: 6px; font-weight:bold;">單位主管<br>簽章</td>
               <td colspan="2" align="left" style="border: 1px solid black; padding: 6px; font-weight:bold;">人事室<br>簽章</td>
               <td colspan="2" align="left" style="border: 1px solid black; padding: 6px; font-weight:bold;">中心主任<br>批示</td>
             </tr>
          </table>
          </div>
        </body>
        </html>
      `;

      const blob = new Blob(['\ufeff' + htmlContent], { type: 'application/msword;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `數辦出差請示單_${new Date().getTime()}.doc`;
      document.body.appendChild(link);
      link.click();
      setTimeout(() => { document.body.removeChild(link); URL.revokeObjectURL(url); }, 150);
      showToast('Word請示單導出成功！', 'success');
    } catch (err) {
      showDialog('alert', '匯出失敗', '無法產生檔案：' + err.message);
    }
  };

  // 100% 會計專用 標楷體雙欄 Excel 出差旅費報告表 (帶有動態公式與中文大寫)
  const exportExcel = (memberName = 'all') => {
    
    
    let dataToExport = records.filter(r => {
      const rStatus = r.status || 'pending';
      const matchMember = memberName === 'all' || r.name === memberName;
      const matchMonth = exportMonth === 'all' || (r.date && r.date.startsWith(exportMonth));
      const matchStatus = exportStatus === 'all' || rStatus === exportStatus;
      return matchMember && matchMonth && matchStatus;
    });

    if (dataToExport.length === 0) return showDialog('alert', '提示', '當前篩選條件下，無任何差旅紀錄可供匯出。');
    const wb = XLSX.utils.book_new();

    const createReportSheet = (personRecords, name, title, group) => {
      const expandedData = expandRecords(personRecords);
      const N = expandedData.length;
      const dateRangeText = getRocDateRange(expandedData);
      const titleText = "雲林縣土庫鎮越港國民小學出差旅費報告表";
      const gridCols = 15; 
      const totalRowIdx = 8 + N; 
      const totalRows = totalRowIdx + 4; 
      const sheetData: any[] = [];

      const baseStyle = { font: { name: "DFKai-SB", sz: 11 }, alignment: { vertical: "center", horizontal: "center", wrapText: true }, border: { top: { style: "thin", color: { rgb: "000000" } }, bottom: { style: "thin", color: { rgb: "000000" } }, left: { style: "thin", color: { rgb: "000000" } }, right: { style: "thin", color: { rgb: "000000" } } } };

      for (let r = 0; r < totalRows; r++) {
        const row: any[] = [];
        for (let c = 0; c < gridCols; c++) row.push({ t: 's', v: '', s: JSON.parse(JSON.stringify(baseStyle)) });
        sheetData.push(row);
      }

      for (let c = 0; c < gridCols; c++) sheetData[0][c].s.border = { top: {style: "none"}, bottom: {style: "none"}, left: {style: "none"}, right: {style: "none"} };
      sheetData[0][0].v = "科目"; sheetData[0][0].s.alignment = { horizontal: "left", vertical: "center" }; sheetData[0][0].s.font = { name: "DFKai-SB", sz: 10 };
      sheetData[0][1].v = "傳票編號："; sheetData[0][1].s.alignment = { horizontal: "left", vertical: "center" }; sheetData[0][1].s.font = { name: "DFKai-SB", sz: 10 };

      for (let c = 0; c < gridCols; c++) sheetData[1][c].s.border = { top: {style: "none"}, bottom: {style: "none"}, left: {style: "none"}, right: {style: "none"} };
      sheetData[1][0].v = titleText; sheetData[1][0].s.font = { name: "DFKai-SB", sz: 22, bold: true }; sheetData[1][0].s.alignment = { horizontal: "center", vertical: "center" };

      sheetData[2][0].v = "姓名"; sheetData[2][1].v = name; sheetData[2][4].v = "職稱"; sheetData[2][5].v = title; sheetData[2][8].v = "職等"; sheetData[2][9].v = ""; 
      sheetData[3][0].v = "出差事由"; sheetData[3][1].v = "如工作記要"; sheetData[3][1].s.alignment = { horizontal: "left", vertical: "center" };
      sheetData[4][0].v = "出差起訖日期"; sheetData[4][1].v = dateRangeText; sheetData[4][1].s.alignment = { horizontal: "left", vertical: "center" };

      const headers1 = [ "編號", `${CURRENT_ROC_YEAR}`, "年", "起(土庫)訖地點", "行程", "行程", "工作記要", "交通費", "交通費", "交通費", "交通費", "住宿費", "雜費", "小計", "備註" ];
      headers1.forEach((val, cIdx) => sheetData[5][cIdx].v = val);
      const headers2 = [ "", "月", "日", "", "半天", "整天", "", "飛機及高鐵", "汽車及捷運", "火車", "船舶", "", "", "", "" ];
      headers2.forEach((val, cIdx) => sheetData[6][cIdx].v = val);

      sheetData[7][0].v = "出差日合計"; sheetData[7][3].f = "E8+F8"; sheetData[7][3].t = "n"; sheetData[7][4].f = `SUM(E9:E${totalRowIdx})`; sheetData[7][4].t = "n"; sheetData[7][5].f = `SUM(F9:F${totalRowIdx})`; sheetData[7][5].t = "n"; sheetData[7][12].v = "費用總計"; sheetData[7][13].f = `N${totalRowIdx + 1}`; sheetData[7][13].t = "n"; sheetData[7][13].s.z = '#,##0';

      expandedData.forEach((r, idx) => {
        const rIdx = 8 + idx; const rowNum = rIdx + 1; 
        const d = r.date ? r.date.split('-') : ['', '', ''];
        const m = d[1] ? parseInt(d[1], 10) : ''; const day = d[2] ? parseInt(d[2], 10) : '';
        const isHsr = r.transportMode === 'hsr';
        const hsrFee = isHsr ? (Number(r.transportFee) || 0) : 0;
        const carFee = !isHsr ? (Number(r.transportFee) || 0) : 0;
        const accomFee = Number(r.accommodationFee) || 0;
        const mealFee = Number(r.mealFee) || 0;

        sheetData[rIdx][0].v = idx + 1;
        if (m) { sheetData[rIdx][1].v = m; sheetData[rIdx][1].t = "n"; }
        if (day) { sheetData[rIdx][2].v = day; sheetData[rIdx][2].t = "n"; }
        sheetData[rIdx][3].v = r.destination || ''; sheetData[rIdx][3].s.alignment = { horizontal: "left", vertical: "center" };
        if (r.excelHalf > 0) { sheetData[rIdx][4].v = r.excelHalf; sheetData[rIdx][4].t = "n"; }
        if (r.excelFull > 0) { sheetData[rIdx][5].v = r.excelFull; sheetData[rIdx][5].t = "n"; }
        sheetData[rIdx][6].v = r.reason || '公出'; sheetData[rIdx][6].s.alignment = { horizontal: "left", vertical: "center" };
        sheetData[rIdx][7].v = hsrFee; sheetData[rIdx][7].t = "n"; sheetData[rIdx][7].s.z = '#,##0';
        sheetData[rIdx][8].v = carFee; sheetData[rIdx][8].t = "n"; sheetData[rIdx][8].s.z = '#,##0';
        sheetData[rIdx][9].v = 0; sheetData[rIdx][9].t = "n"; sheetData[rIdx][9].s.z = '#,##0';
        sheetData[rIdx][10].v = 0; sheetData[rIdx][10].t = "n"; sheetData[rIdx][10].s.z = '#,##0';
        sheetData[rIdx][11].v = accomFee; sheetData[rIdx][11].t = "n"; sheetData[rIdx][11].s.z = '#,##0';
        sheetData[rIdx][12].v = mealFee; sheetData[rIdx][12].t = "n"; sheetData[rIdx][12].s.z = '#,##0';
        sheetData[rIdx][13].f = `H${rowNum}+I${rowNum}+J${rowNum}+K${rowNum}+L${rowNum}+M${rowNum}`; sheetData[rIdx][13].t = "n"; sheetData[rIdx][13].s.z = '#,##0';
        sheetData[rIdx][14].v = r.remarks || ''; sheetData[rIdx][14].s.alignment = { horizontal: "left", vertical: "center" };
      });

      sheetData[totalRowIdx][0].v = "合計";
      sheetData[totalRowIdx][7].f = `SUM(H9:H${totalRowIdx})`; sheetData[totalRowIdx][7].t = "n"; sheetData[totalRowIdx][7].s.z = '#,##0';
      sheetData[totalRowIdx][8].f = `SUM(I9:I${totalRowIdx})`; sheetData[totalRowIdx][8].t = "n"; sheetData[totalRowIdx][8].s.z = '#,##0';
      sheetData[totalRowIdx][9].f = `SUM(J9:J${totalRowIdx})`; sheetData[totalRowIdx][9].t = "n"; sheetData[totalRowIdx][9].s.z = '#,##0';
      sheetData[totalRowIdx][10].f = `SUM(K9:K${totalRowIdx})`; sheetData[totalRowIdx][10].t = "n"; sheetData[totalRowIdx][10].s.z = '#,##0';
      sheetData[totalRowIdx][11].f = `SUM(L9:L${totalRowIdx})`; sheetData[totalRowIdx][11].t = "n"; sheetData[totalRowIdx][11].s.z = '#,##0';
      sheetData[totalRowIdx][12].f = `SUM(M9:M${totalRowIdx})`; sheetData[totalRowIdx][12].t = "n"; sheetData[totalRowIdx][12].s.z = '#,##0';
      sheetData[totalRowIdx][13].f = `SUM(N9:N${totalRowIdx})`;
      sheetData[totalRowIdx][13].t = "n";
      sheetData[totalRowIdx][13].s.z = '#,##0';

      const rawGrandTotal = expandedData.reduce((sum, r) => sum + (Number(r.totalFee) || 0), 0);
      const chineseAmt = numberToChinese(rawGrandTotal);

      const amtRowIdx = totalRowIdx + 1;
      sheetData[amtRowIdx][0].v = "上列出差旅費計"; sheetData[amtRowIdx][3].v = chineseAmt; sheetData[amtRowIdx][3].s.alignment = { horizontal: "left", vertical: "center" }; sheetData[amtRowIdx][11].v = "具領人簽";

      const emptyRowIdx = totalRowIdx + 2;
      for (let c = 0; c < gridCols; c++) sheetData[emptyRowIdx][c].s.border = { top: {style: "none"}, bottom: {style: "none"}, left: {style: "none"}, right: {style: "none"} };

      const signRowIdx = totalRowIdx + 3;
      for (let c = 0; c < gridCols; c++) sheetData[signRowIdx][c].s.border = { top: {style: "none"}, bottom: {style: "none"}, left: {style: "none"}, right: {style: "none"} };
      sheetData[signRowIdx][0].v = "出差人"; sheetData[signRowIdx][3].v = "單位主管"; sheetData[signRowIdx][6].v = "人事單位"; sheetData[signRowIdx][9].v = "會計單位"; sheetData[signRowIdx][12].v = "校長";

      const ws = XLSX.utils.aoa_to_sheet([]);
      for (let r = 0; r < totalRows; r++) { for (let c = 0; c < gridCols; c++) { const cellRef = XLSX.utils.encode_cell({ r, c }); ws[cellRef] = sheetData[r][c]; } }
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: totalRows - 1, c: gridCols - 1 } });

      ws['!merges'] = [
        { s: { r: 1, c: 0 }, e: { r: 1, c: 14 } }, 
        { s: { r: 2, c: 1 }, e: { r: 2, c: 3 } },  
        { s: { r: 2, c: 5 }, e: { r: 2, c: 7 } },  
        { s: { r: 2, c: 9 }, e: { r: 2, c: 14 } }, 
        { s: { r: 3, c: 1 }, e: { r: 3, c: 14 } }, 
        { s: { r: 4, c: 1 }, e: { r: 4, c: 14 } }, 
        { s: { r: 5, c: 0 }, e: { r: 6, c: 0 } }, 
        { s: { r: 5, c: 1 }, e: { r: 5, c: 2 } }, 
        { s: { r: 5, c: 3 }, e: { r: 6, c: 3 } }, 
        { s: { r: 5, c: 4 }, e: { r: 5, c: 5 } }, 
        { s: { r: 5, c: 6 }, e: { r: 6, c: 6 } }, 
        { s: { r: 5, c: 7 }, e: { r: 5, c: 10 } }, 
        { s: { r: 5, c: 11 }, e: { r: 6, c: 11 } }, 
        { s: { r: 5, c: 12 }, e: { r: 6, c: 12 } }, 
        { s: { r: 5, c: 13 }, e: { r: 6, c: 13 } }, 
        { s: { r: 5, c: 14 }, e: { r: 6, c: 14 } }, 
        { s: { r: 7, c: 0 }, e: { r: 7, c: 2 } }, 
        { s: { r: totalRowIdx, c: 0 }, e: { r: totalRowIdx, c: 12 } }, 
        { s: { r: totalRowIdx + 1, c: 0 }, e: { r: totalRowIdx + 1, c: 2 } }, 
        { s: { r: totalRowIdx + 1, c: 3 }, e: { r: totalRowIdx + 1, c: 10 } }, 
        { s: { r: totalRowIdx + 1, c: 11 }, e: { r: totalRowIdx + 1, c: 12 } }, 
        { s: { r: totalRowIdx + 1, c: 13 }, e: { r: totalRowIdx + 1, c: 14 } }, 
        { s: { r: signRowIdx, c: 0 }, e: { r: signRowIdx, c: 2 } }, 
        { s: { r: signRowIdx, c: 3 }, e: { r: signRowIdx, c: 5 } }, 
        { s: { r: signRowIdx, c: 6 }, e: { r: signRowIdx, c: 8 } }, 
        { s: { r: signRowIdx, c: 9 }, e: { r: signRowIdx, c: 11 } }, 
        { s: { r: signRowIdx, c: 12 }, e: { r: signRowIdx, c: 14 } }, 
      ];
      ws['!views'] = [{ showGridLines: true }];
      ws['!cols'] = [ { wch: 5 }, { wch: 4 }, { wch: 4 }, { wch: 18 }, { wch: 6 }, { wch: 6 }, { wch: 25 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 11 }, { wch: 15 } ];
      return ws;
    };

    if (memberName === 'all') {
      const statusText = exportStatus === 'all' ? "全部狀態" : exportStatus === 'verified' ? "已核銷" : "待核銷";
      const monthText = exportMonth === 'all' ? "所有月份" : exportMonth.replace('-', '年 ') + '月';
      const summaryAoa = [
        ["雲林縣土庫鎮越港國民小學 (數位學習推動辦公室) 國內出差旅費報銷彙整總表"],
        [`篩選區間：${monthText} (${statusText})  |  列印日期：` + new Date().toLocaleDateString('zh-TW')],
        [],
        ["編號", "姓名", "職稱", "組別", "出差次數", "交通費合計", "住宿費合計", "膳雜費合計", "總計金額"]
      ];

      const recordsGrouped = {};
      employees.forEach(m => {
        const mRecs = dataToExport.filter(r => r.name === m.name);
        if (mRecs.length > 0) {
          const expandedRecs = expandRecords(mRecs); 
          recordsGrouped[m.name] = { member: m, recs: mRecs, expandedRecs: expandedRecs };
        }
      });

      let rowIndex = 4;
      Object.keys(recordsGrouped).forEach((name, i) => {
        const g = recordsGrouped[name];
        const count = g.recs.length; 
        summaryAoa.push([i + 1, name, g.member.title, g.member.group, count, 0, 0, 0, 0]);
        rowIndex++;
      });

      const summaryTotalRowIdx = rowIndex;
      summaryAoa.push(["合 計", "", "", "", "", 0, 0, 0, 0]);

      const summaryWs = XLSX.utils.aoa_to_sheet(summaryAoa);
      summaryWs['!merges'] = [ { s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 8 } }, { s: { r: summaryTotalRowIdx, c: 0 }, e: { r: summaryTotalRowIdx, c: 4 } } ];

      Object.keys(recordsGrouped).forEach((name, i) => {
        const rIdx = 4 + i;
        const safeSheetName = name.replace(/[\\*?:/[\]]/g, ''); 
        const pTabTotalRow = 9 + recordsGrouped[name].expandedRecs.length; 
        summaryWs[XLSX.utils.encode_cell({ r: rIdx, c: 5 })] = { t: 'n', f: `'${safeSheetName}'!H${pTabTotalRow}+'${safeSheetName}'!I${pTabTotalRow}+'${safeSheetName}'!J${pTabTotalRow}+'${safeSheetName}'!K${pTabTotalRow}` };
        summaryWs[XLSX.utils.encode_cell({ r: rIdx, c: 6 })] = { t: 'n', f: `'${safeSheetName}'!L${pTabTotalRow}` };
        summaryWs[XLSX.utils.encode_cell({ r: rIdx, c: 7 })] = { t: 'n', f: `'${safeSheetName}'!M${pTabTotalRow}` };
        summaryWs[XLSX.utils.encode_cell({ r: rIdx, c: 8 })] = { t: 'n', f: `'${safeSheetName}'!N${pTabTotalRow}` };
      });

      summaryWs[XLSX.utils.encode_cell({ r: summaryTotalRowIdx, c: 5 })] = { t: 'n', f: `SUM(F5:F${summaryTotalRowIdx})` };
      summaryWs[XLSX.utils.encode_cell({ r: summaryTotalRowIdx, c: 6 })] = { t: 'n', f: `SUM(G5:G${summaryTotalRowIdx})` };
      summaryWs[XLSX.utils.encode_cell({ r: summaryTotalRowIdx, c: 7 })] = { t: 'n', f: `SUM(H5:H${summaryTotalRowIdx})` };
      summaryWs[XLSX.utils.encode_cell({ r: summaryTotalRowIdx, c: 8 })] = { t: 'n', f: `SUM(I5:I${summaryTotalRowIdx})` };

      const summaryRange = XLSX.utils.decode_range(summaryWs['!ref']);
      for (let r = summaryRange.s.r; r <= summaryRange.e.r; r++) {
        for (let c = summaryRange.s.c; c <= summaryRange.e.c; c++) {
          const cellAddr = XLSX.utils.encode_cell({ r, c });
          if (!summaryWs[cellAddr]) summaryWs[cellAddr] = { t: 's', v: '' };
          const cell = summaryWs[cellAddr]; cell.s = cell.s || {}; cell.s.font = { name: "DFKai-SB", sz: 11 };
          if (r > 1) {
            cell.s.border = { top: { style: "thin", color: { rgb: "000000" } }, bottom: { style: "thin", color: { rgb: "000000" } }, left: { style: "thin", color: { rgb: "000000" } }, right: { style: "thin", color: { rgb: "000000" } } };
            cell.s.alignment = { vertical: "center", horizontal: "center" };
          } else {
            cell.s.border = { top: {style: "none"}, bottom: {style: "none"}, left: {style: "none"}, right: {style: "none"} };
            cell.s.alignment = { vertical: "center", horizontal: "center" };
            if (r === 0) cell.s.font = { name: "DFKai-SB", sz: 16, bold: true };
          }
        }
      }

      summaryWs['!views'] = [{ showGridLines: true }];
      summaryWs['!cols'] = [ { wch: 8 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 20 } ];
      XLSX.utils.book_append_sheet(wb, summaryWs, "彙整總表");

      Object.keys(recordsGrouped).forEach(name => {
        const g = recordsGrouped[name];
        const personWs = createReportSheet(g.recs, name, g.member.title, g.member.group);
        XLSX.utils.book_append_sheet(wb, personWs, name.replace(/[\\*?:/[\]]/g, ''));
      });

      XLSX.writeFile(wb, `雲林數辦_全體出差旅費報告表_(${monthText}_${statusText}).xlsx`);
    } else {
      const memberInfo = employees.find(m => m.name === memberName);
      const ws = createReportSheet(dataToExport, memberName, memberInfo.title, memberInfo.group);
      XLSX.utils.book_append_sheet(wb, ws, "出差旅費報告表");
      XLSX.writeFile(wb, `雲林數辦_出差旅費報告表_${memberName}.xlsx`);
    }
  };

  // 審核同意請假（自動扣抵特休、補休，或累加事/病/喪假）
  const handleApproveLeave = async (reqId) => {
    const req = requests.find((r: any) => r.id === reqId) as any;
    if (!req) return;
    
    const emp = employees.find(e => e.name === req.applicant);
    if (!emp) return;

    let updatedRemainingTe = { ...emp.remainingTe };
    let updatedRemainingBu = { ...emp.remainingBu };
    let updatedTakenShi = { ...emp.takenShi };
    let updatedTakenBing = { ...emp.takenBing };
    let updatedTakenSang = { ...emp.takenSang };

    const totalHours = req.hours;
    if (req.leaveType === '特休') {
      let currTotalHours = emp.remainingTe.d * 8 + emp.remainingTe.h;
      if (currTotalHours < totalHours) {
        showDialog('alert', '錯誤', '核准失敗！該同仁賸餘特休時數不足。');
        return;
      }
      currTotalHours -= totalHours;
      updatedRemainingTe = { d: Math.floor(currTotalHours / 8), h: Math.round((currTotalHours % 8) * 10) / 10 };
    } else if (req.leaveType === '補休') {
      let currTotalHours = emp.remainingBu.d * 8 + emp.remainingBu.h;
      if (currTotalHours < totalHours) {
        showDialog('alert', '錯誤', '核准失敗！該同仁賸餘補休時數不足。');
        return;
      }
      currTotalHours -= totalHours;
      updatedRemainingBu = { d: Math.floor(currTotalHours / 8), h: Math.round((currTotalHours % 8) * 10) / 10 };
    } else if (req.leaveType === '事假') {
      let currTotalHours = emp.takenShi.d * 8 + emp.takenShi.h;
      currTotalHours += totalHours;
      updatedTakenShi = { d: Math.floor(currTotalHours / 8), h: Math.round((currTotalHours % 8) * 10) / 10 };
    } else if (req.leaveType === '病假') {
      let currTotalHours = emp.takenBing.d * 8 + emp.takenBing.h;
      currTotalHours += totalHours;
      updatedTakenBing = { d: Math.floor(currTotalHours / 8), h: Math.round((currTotalHours % 8) * 10) / 10 };
    } else if (req.leaveType === '喪假') {
      let currTotalHours = emp.takenSang.d * 8 + emp.takenSang.h;
      currTotalHours += totalHours;
      updatedTakenSang = { d: Math.floor(currTotalHours / 8), h: Math.round((currTotalHours % 8) * 10) / 10 };
    }

    try {
      await updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'employees', emp.name), {
        remainingTe: updatedRemainingTe,
        remainingBu: updatedRemainingBu,
        takenShi: updatedTakenShi,
        takenBing: updatedTakenBing,
        takenSang: updatedTakenSang
      });
      await updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'leaves', req.id), { status: '核准' });
      showDialog('alert', '核准成功', `已成功核准 ${req.applicant} 的 ${req.leaveType}！其時數已自動更新。`);
    } catch (e) {
      showDialog('alert', '錯誤', '核准更新雲端失敗。');
    }
  };

  // 審核假日加班（自動為名單中的同仁加總補休）
  const handleApproveOvertime = async (ot) => {
    if (!user) return;
    try {
      for (const pName of ot.participants) {
        const emp = employees.find(e => e.name === pName);
        if (emp) {
          let currTotalHours = emp.remainingBu.d * 8 + emp.remainingBu.h;
          currTotalHours += ot.hours;
          const updatedRemainingBu = { d: Math.floor(currTotalHours / 8), h: Math.round((currTotalHours % 8) * 10) / 10 };
          await updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'employees', pName), { remainingBu: updatedRemainingBu });
        }
      }
      await updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'overtimes', ot.id), { status: '已核准' });
      showDialog('alert', '核准成功', `假日加班協商同意！時數已順利加至同仁帳戶。`);
    } catch (e) {
      showDialog('alert', '錯誤', '時數加總儲存錯誤。');
    }
  };

  // 退回請假/加班
  const handleReject = (reqId) => {
    updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'leaves', reqId), { status: '已退回' })
      .then(() => showDialog('alert', '提示', '已退回該請假單。'))
      .catch(() => showDialog('alert', '錯誤', '退回請假單出錯。'));
  };

  const handleRejectOvertime = (otId) => {
    updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'overtimes', otId), { status: '已退回' })
      .then(() => showDialog('alert', '提示', '已退回該加班協商。'))
      .catch(() => showDialog('alert', '錯誤', '退回加班單出錯。'));
  };

  // 批次更改核銷狀態
  const handleBatchStatusUpdate = async (newStatus) => {
    if (selectedRecordIds.length === 0) return showDialog('alert', '提示', '請選擇要變更狀態的差旅紀錄');
    setIsProcessing(true);
    try {
      await Promise.all(selectedRecordIds.map(async (id) => {
        const recordRef = doc(db, 'artifacts', appId, 'public', 'data', 'travel_expenses', id);
        await updateDoc(recordRef, { status: newStatus });
      }));
      showDialog('alert', '成功', `狀態已變更為【${newStatus === 'verified' ? '已核銷' : '待核銷'}】！`);
      setSelectedRecordIds([]);
    } catch (error) {
      showDialog('alert', '錯誤', '變更狀態失敗。');
    } finally {
      setIsProcessing(false);
    }
  };

  // 主管密碼解鎖驗證
  const handleVerifyPassword = (e) => {
    e.preventDefault();
    if (inputPassword === correctPassword) {
      setIsUnlocked(true);
      showDialog('alert', '解鎖成功', '主管端密碼驗證通過！');
    } else {
      showDialog('alert', '密碼錯誤', '認證密碼不正確，請重新輸入！');
    }
  };

  // 歷史資料強制刪除 (需輸入 056341014 安全防護密碼)
  const handleConfirmDelete = async (e) => {
    e.preventDefault();
    if (deletePassword !== correctDeletePassword) {
      showDialog('alert', '安全防衛鎖定', '密碼不正確！安全防刪驗證失敗！');
      return;
    }
    try {
      if (deleteTarget.type === 'leave') {
        await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', 'leaves', deleteTarget.id));
        showDialog('alert', '成功', '請假單已永久移除！');
      } else if (deleteTarget.type === 'overtime') {
        await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', 'overtimes', deleteTarget.id));
        showDialog('alert', '成功', '協商加班單已永久移除！');
      }
      setDeleteTarget(null);
      setDeletePassword('');
    } catch (e) {
      showDialog('alert', '錯誤', '寫入移除錯誤！');
    }
  };

  const handleDeleteTravelClick = (id) => {
    showDialog('prompt', '安全驗證鎖', '請輸入管理安全鎖密碼 (056341014) 以永久刪除此紀錄：', (pwd) => {
      if (pwd !== '056341014') {
        setTimeout(() => showDialog('alert', '錯誤', '驗證密碼不正確，拒絕刪除！'), 300);
        return;
      }
      setTimeout(() => {
        showDialog('confirm', '最終確認', '確定要永久刪除此紀錄嗎？', async () => {
          try {
            await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', 'travel_expenses', id));
            setSelectedRecordIds(prev => prev.filter(selectedId => selectedId !== id));
            showDialog('alert', '成功', '差旅紀錄已刪除！');
          } catch (e) {
            console.error(e);
          }
        });
      }, 300);
    }, 'password');
  };

  // 跨系統核心整合：一鍵將加班名單同步至差旅填報
  const handleCarryOverToTravel = (ot) => {
    let formattedDate = ot.workDate;
    const parts = ot.workDate.split('-');
    if (parts.length === 3) {
      let year = parseInt(parts[0], 10);
      if (year < 1000) year += 1911; 
      formattedDate = `${year}-${parts[1]}-${parts[2]}`;
    }
    
    setFormData(prev => ({
      ...prev,
      date: formattedDate,
      memberNames: ot.participants,
      reason: ot.activityName,
      destination: '', 
      isMultiDay: false,
    }));
    
    setActiveTab('calc'); 
    showDialog('alert', '智慧同步成功', `已將「${ot.activityName}」的 ${ot.participants.length} 位出勤同仁，智慧載入會計差旅費填報核心！`);
  };

  // 智慧同步請假（公假且可支領差旅費）一鍵帶入出差填報
  const handleLeaveCarryOverToTravel = (req: any) => {
    const formattedDate = req.startDate || req.date;
    
    // 智慧型出差時段判定：時數在半天(4小時)以下且在 12:00 前開始者為「上午半天」，其餘為「下午半天」
    let duration = 'full';
    if (req.hours && Number(req.hours) <= 4) {
      const startHour = req.startTime ? parseInt(req.startTime.split(':')[0], 10) : 8;
      if (startHour < 12) {
        duration = 'morning';
      } else {
        duration = 'afternoon';
      }
    }

    setFormData(prev => ({
      ...prev,
      date: formattedDate,
      memberNames: [req.applicant],
      reason: req.reason || '公務洽公',
      startLocation: '越港國小',
      destination: req.location || '',
      isMultiDay: req.startDate !== req.endDate,
      endDate: req.endDate || req.startDate,
      duration: duration,
      isOutCounty: false, // 將由系統自動依據出差地點關鍵字判斷縣內/縣外
      transportMode: 'car', // 預設汽車
      distance: '',
      isRoundTrip: true
    }));

    setActiveTab('calc');
    showDialog(
      'alert', 
      '智慧同步成功', 
      `已為您一鍵同步「${req.applicant}」於民國 ${formatMinguoDateText(formattedDate)} 的公出資料！\n\n時段已依請假時數智慧判定為：【${duration === 'morning' ? '上午半天' : duration === 'afternoon' ? '下午半天' : '整天'}】。`
    );
  };

  const handleExportWord = (ot) => {
    const dateText = formatMinguoDateText(ot.workDate);
    const weekdayText = getWeekdayText(ot.workDate);
    const chunked = chunkParticipants(ot.participants);
    const participantsRows = chunked.map(([p1, p2]) => `
        <tr style="height: 28pt; page-break-inside: avoid;">
          <td style="border: 1px solid black; padding: 6px 12px; font-family: 'DFKai-SB', '標楷體', serif; font-size: 13pt; width: 50%;">姓名：<strong>${p1}</strong></td>
          <td style="border: 1px solid black; padding: 6px 12px; font-family: 'DFKai-SB', '標楷體', serif; font-size: 13pt; width: 50%;">${p2 ? `姓名：<strong>${p2}</strong>` : '姓名：'}</td>
        </tr>
    `).join('');
    const htmlContent = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head><meta charset="utf-8"><title>休息日工作協商同意書</title>
        <style>
          @page { size: A4 portrait; margin: 2.0cm; }
          body { font-family: 'DFKai-SB', '標楷體', serif; line-height: 1.8; color: black; }
          .title { text-align: center; font-size: 24pt; font-weight: bold; letter-spacing: 6px; border-bottom: 2px solid black; padding-bottom: 6px; margin-bottom: 25px; }
          .content { font-size: 14pt; text-align: justify; text-indent: 2em; line-height: 2.0; margin-bottom: 20px; }
          .underlined { border-bottom: 1.5px solid black; padding: 0 4px; font-weight: bold; }
          .law-box { border: 1px solid black; background-color: #fcfcfc; padding: 10px; font-size: 10pt; font-family: sans-serif; margin-bottom: 20px; line-height: 1.5; page-break-inside: avoid; }
          .table-title { font-size: 12pt; font-weight: bold; margin-bottom: 8px; page-break-after: avoid; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 25px; page-break-inside: avoid; }
          .signature-table { width: 100%; border: none; margin-top: 35px; page-break-inside: avoid; }
          .signature-cell { width: 33%; border: none; text-align: center; font-family: sans-serif; font-size: 11pt; }
          .sig-line { border-top: 1.5px solid black; padding-top: 8px; margin: 0 15px; }
        </style>
      </head>
      <body>
        <div class="title">協商同意書</div>
        <div class="content">因 <span class="underlined">${dateText}</span> ( ${weekdayText} ) 上午 ${ot.startTime} 時至下午 ${ot.endTime} 時原訂為勞工休息日，本校 <span class="underlined">（名單如下表）</span> 同意於前開休息日協助 <span class="underlined">${ot.activityName}</span>，並經勞雇雙方協商同意於 ${dateText} 次日起一年內完成補休 (共計 <span class="underlined">${ot.hours}</span> 小時)，特此證明。</div>
        <div class="law-box"><strong>法條依據：</strong><br/>勞動基準法第 32-1 條：雇主依休息日工作後，依勞工意願選擇補休並經雇主同意者，應依勞工工作之時數計算補休時數。</div>
        <div class="table-title">協商同意勞工名單：</div><table>${participantsRows}</table>
        <table class="signature-table"><tr><td class="signature-cell"><div class="sig-line">執行秘書</div></td><td class="signature-cell"><div class="sig-line">中心主任</div></td><td class="signature-cell"><div class="sig-line">校長</div></td></tr></table>
      </body></html>
    `;
    const blob = new Blob([htmlContent], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `休息日工作協商同意書_${ot.workDate.replace(/-/g, '')}.doc`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('Word 協商同意書已導出！');
  };

  const handlePrintPDF = (ot) => {
    const dateText = formatMinguoDateText(ot.workDate);
    const weekdayText = getWeekdayText(ot.workDate);
    const chunked = chunkParticipants(ot.participants);
    const participantsRows = chunked.map(([p1, p2]) => `
      <tr style="height: 30pt; page-break-inside: avoid;">
        <td style="border: 1px solid black; padding: 8px 12px; font-family: 'DFKai-SB', '標楷體', serif; font-size: 13pt; width: 50%;">姓名：<strong>${p1}</strong></td>
        <td style="border: 1px solid black; padding: 8px 12px; font-family: 'DFKai-SB', '標楷體', serif; font-size: 13pt; width: 50%;">${p2 ? `姓名：<strong>${p2}</strong>` : '姓名：'}</td>
      </tr>
    `).join('');

    const printWindow = window.open('', '_blank', 'width=850,height=950');
    if (!printWindow) return showToast('請解鎖彈出式視窗限制以列印！', 'error');
    
    printWindow.document.write(`
      <!DOCTYPE html><html><head><meta charset="utf-8"><title>休息日工作協商同意書_${ot.workDate}</title>
      <style>@page { size: A4 portrait; margin: 1.8cm; } body { font-family: 'DFKai-SB', '標楷體', 'PMingLiU', '新細明體', serif; line-height: 1.8; color: black; margin: 0; padding: 0; } .title { text-align: center; font-size: 24pt; font-weight: bold; letter-spacing: 8px; border-bottom: 2px solid black; padding-bottom: 6px; margin-bottom: 25px; } .content { font-size: 14pt; text-align: justify; text-indent: 2.5em; line-height: 2.1; margin-bottom: 20px; } .underlined { border-bottom: 1.5px solid black; padding: 0 4px; font-weight: bold; } .law-box { border: 1px solid black; background-color: #fafafa; padding: 10px 14px; font-size: 10pt; font-family: sans-serif; margin-bottom: 20px; line-height: 1.5; page-break-inside: avoid; } .table-title { font-size: 12pt; font-weight: bold; margin-bottom: 8px; page-break-after: avoid; } table { width: 100%; border-collapse: collapse; margin-bottom: 30px; page-break-inside: avoid; } .signature-table { width: 100%; border: none; margin-top: 45px; page-break-inside: avoid; } .signature-cell { width: 33%; border: none; text-align: center; font-family: sans-serif; font-size: 11pt; } .sig-line { border-top: 1.5px solid black; padding-top: 10px; margin: 0 15px; }</style>
      </head><body>
      <div class="title">協商同意書</div>
      <div class="content">因 <span class="underlined">${dateText}</span> ( ${weekdayText} ) 上午 ${ot.startTime} 時至下午 ${ot.endTime} 時原訂為勞工休息日，本校 <span class="underlined">（名單如下表）</span> 同意於前開休息日協助 <span class="underlined">${ot.activityName}</span>，並經勞雇雙方協商同意於 ${dateText} 次日起一年內完成補休 (共計 <span class="underlined">${ot.hours}</span> 小時)，特此證明。</div>
      <div class="law-box"><strong>法條依據：</strong><br/><strong>勞動基準法第 32-1 條：</strong><br/>雇主使勞工休息日工作後，依勞工意願選擇補休並經雇主同意者，應依勞工工作之時數計算補休時數。前項之補休，其補休期限由勞雇雙方協商。</div>
      <div class="table-title">協商同意勞工名單：</div><table>${participantsRows}</table>
      <table class="signature-table"><tr><td class="signature-cell"><div class="sig-line">執行秘書</div></td><td class="signature-cell"><div class="sig-line">中心主任</div></td><td class="signature-cell"><div class="sig-line">校長</div></td></tr></table>
      <script>window.onload = function() { setTimeout(function() { window.print(); }, 300); window.onafterprint = function() { window.close(); }; };</script>
      </body></html>
    `);
    printWindow.document.close();
    showToast('已安全喚醒 A4 獨立 PDF 列印頁面！');
  };

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans relative">
      
      {/* Toast 提示 */}
      {notification && (
        <div className="fixed top-16 right-4 z-50 animate-bounce print:hidden">
          <div className={`p-4 rounded-xl shadow-lg border flex items-center space-x-3 text-sm font-semibold ${
            notification.type === 'success' ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-rose-50 text-rose-800 border-rose-200'
          }`}>
            <CheckCircle className="h-5 w-5 text-emerald-600" />
            <span>{notification.message}</span>
          </div>
        </div>
      )}

      {/* 客製化 Dialog */}
      {dialog.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm px-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full overflow-hidden">
            <div className={`px-6 py-4 border-b flex items-center gap-3 bg-indigo-50 text-indigo-900 border-indigo-100`}>
              <MapPin className="w-5 h-5 text-indigo-600" />
              <h3 className="font-bold text-lg">{dialog.title}</h3>
            </div>
            <div className="px-6 py-5 text-gray-700">
              <p className="mb-4 whitespace-pre-line text-sm">{dialog.message}</p>
              {dialog.type === 'prompt' && (
                <input 
                  type={dialog.inputType} 
                  autoFocus
                  value={dialog.inputValue}
                  onChange={(e) => setDialog(p => ({...p, inputValue: e.target.value}))}
                  onKeyDown={(e) => { if(e.key === 'Enter' && dialog.onConfirm) dialog.onConfirm(dialog.inputValue); }}
                  className="w-full p-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm"
                  placeholder="請在此輸入..."
                />
              )}
            </div>
            <div className="px-6 py-4 bg-gray-50 flex justify-end gap-3 border-t">
              {(dialog.type === 'confirm' || dialog.type === 'prompt') && (
                <button onClick={closeDialog} className="px-4 py-2 text-sm text-gray-600 hover:bg-gray-200 rounded-lg transition-colors">取消</button>
              )}
              <button 
                onClick={() => {
                  if (dialog.onConfirm) {
                    if (dialog.type === 'prompt') dialog.onConfirm(dialog.inputValue);
                    else dialog.onConfirm();
                  } else {
                    closeDialog();
                  }
                }} 
                className="px-5 py-2 text-sm text-white rounded-lg transition-colors bg-indigo-600 hover:bg-indigo-700"
              >
                確認
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 頂部控制欄 */}
      <header className="bg-slate-950 text-white p-4 shadow flex flex-col sm:flex-row justify-between items-center space-y-2 sm:space-y-0 print:hidden">
        <div className="flex items-center space-x-3">
          <CalendarDays className="h-8 w-8 text-indigo-400 animate-pulse" />
          <div>
            <h1 className="text-xl font-bold">雲林縣數位學習推動辦公室</h1>
            <p className="text-xs text-slate-400">出缺勤管理與差旅費核銷整合系統 (115學年度專用)</p>
          </div>
        </div>

        {/* 角色權限切換 */}
        <div className="flex items-center space-x-2 bg-slate-900 p-2 rounded-xl text-sm border border-slate-800">
          <span className="text-slate-300 font-medium">切換操作同仁:</span>
          <select 
            value={currentUser} 
            onChange={(e) => {
              setCurrentUser(e.target.value);
              setFormApplicant(e.target.value === '主管端' ? '' : e.target.value);
            }}
            className="bg-slate-800 text-white rounded-lg px-2 py-1 outline-none text-xs border border-slate-700 cursor-pointer font-bold"
          >
            <option value="主管端">主管審核端 (簽核授權)</option>
            {employees.map(e => (
              <option key={e.name} value={e.name}>{e.name} ({e.group})</option>
            ))}
          </select>
          {currentUser === '主管端' && (
            <span className="text-xs ml-1">
              {isUnlocked ? <span className="text-emerald-400 font-bold">🔓 已授權</span> : <span className="text-rose-400 font-bold">🔒 鎖定中</span>}
            </span>
          )}
        </div>
      </header>

      <div className="flex flex-1 flex-col lg:flex-row print:p-0">
        
        {/* 側邊雙軌導覽列 */}
        <aside className="w-full lg:w-64 bg-slate-900 text-slate-200 p-4 shrink-0 flex flex-col print:hidden justify-between border-r border-slate-800">
          <div className="space-y-6">
            <div>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2 px-2">【人事請假差勤】</p>
              <div className="space-y-1">
                <button onClick={() => setActiveTab('attendance_dashboard')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 ${activeTab === 'attendance_dashboard' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800'}`}>
                  <Sparkles className="h-4 w-4" /> <span>辦公室差勤總覽</span>
                </button>
                <button onClick={() => setActiveTab('attendance_apply')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 ${activeTab === 'attendance_apply' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800'}`}>
                  <Plus className="h-4 w-4" /> <span>線上請假/加班</span>
                </button>
                <button onClick={() => setActiveTab('attendance_roster')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 ${activeTab === 'attendance_roster' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800'}`}>
                  <Users className="h-4 w-4" /> <span>同仁特休細算表</span>
                </button>
                <button onClick={() => setActiveTab('attendance_audit')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 relative ${activeTab === 'attendance_audit' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800'}`}>
                  <CheckCircle className="h-4 w-4" /> <span>主管審核中心</span>
                  {(requests.filter((r: any) => r.status === '待審核').length + overtimeRequests.filter((r: any) => r.status === '待審核').length) > 0 && (
                    <span className="absolute right-2 top-2.5 bg-rose-500 text-white text-[10px] w-5.5 h-5.5 rounded-full flex items-center justify-center font-bold">
                      {requests.filter((r: any) => r.status === '待審核').length + overtimeRequests.filter((r: any) => r.status === '待審核').length}
                    </span>
                  )}
                </button>
              </div>
            </div>

            <div>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2 px-2">【會計差旅費核銷】</p>
              <div className="space-y-1">
                <button onClick={() => setActiveTab('calc')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 ${activeTab === 'calc' ? 'bg-emerald-600 text-white' : 'hover:bg-slate-800'}`}>
                  <Calculator className="h-4 w-4" /> <span>差旅費填報核心</span>
                </button>
                <button onClick={() => setActiveTab('history')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 ${activeTab === 'history' ? 'bg-emerald-600 text-white' : 'hover:bg-slate-800'}`}>
                  <History className="h-4 w-4" /> <span>差旅明細清冊</span>
                </button>
                <button onClick={() => setActiveTab('export')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 ${activeTab === 'export' ? 'bg-emerald-600 text-white' : 'hover:bg-slate-800'}`}>
                  <FileSpreadsheet className="h-4 w-4" /> <span>報表下載中心</span>
                </button>
                <button onClick={() => setActiveTab('settings')} className={`w-full text-left px-3 py-2.5 rounded-lg text-sm font-semibold flex items-center space-x-2 ${activeTab === 'settings' ? 'bg-emerald-600 text-white' : 'hover:bg-slate-800'}`}>
                  <Settings className="h-4 w-4" /> <span>系統服務設定</span>
                </button>
              </div>
            </div>
          </div>

          <div className="pt-4 border-t border-slate-800 text-center text-[10px] text-slate-500 font-medium">
            越港國小數位辦公室 · 115學年度
          </div>
        </aside>

        {/* 系統主操作顯示區 */}
        <main className="flex-1 p-6 overflow-y-auto print:p-0">
          
          {/* ================================================== */}
          {/* A. 辦公室差勤總覽 */}
          {/* ================================================== */}
          {activeTab === 'attendance_dashboard' && (
            <div className="space-y-6">
              
              {/* 同仁請假最新狀態 */}
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div className="p-5 border-b border-slate-100 bg-slate-50 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
                  <div>
                    <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                      <Clock className="text-indigo-600 h-5 w-5" />
                      <span>同仁請假及差勤最新狀態</span>
                    </h3>
                    <p className="text-xs text-slate-500">雲端即時連線同步，記錄同仁事假、特休、公假、代理人設定資訊</p>
                  </div>
                </div>

                {/* 歷史請假篩選與搜尋 */}
                <div className="p-4 bg-slate-50 border-b border-slate-100 grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
                  <div className="relative">
                    <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-slate-400" />
                    <input
                      type="text"
                      placeholder="搜尋姓名或請假事由..."
                      value={leaveSearch}
                      onChange={(e) => { setLeaveSearch(e.target.value); setLeavePage(1); }}
                      className="pl-8 w-full p-2 bg-white border border-slate-300 rounded-lg outline-none font-medium"
                    />
                  </div>
                  <div className="flex items-center space-x-2">
                    <span className="text-slate-500 shrink-0 font-bold">年度</span>
                    <select value={leaveFilterYear} onChange={(e) => { setLeaveFilterYear(e.target.value); setLeavePage(1); }} className="w-full p-2 bg-white border border-slate-300 rounded-lg font-semibold">
                      {leaveYears.map(y => <option key={y} value={y}>{y}</option>)}
                    </select>
                  </div>
                  <div className="flex items-center space-x-2">
                    <span className="text-slate-500 shrink-0 font-bold">月份</span>
                    <select value={leaveFilterMonth} onChange={(e) => { setLeaveFilterMonth(e.target.value); setLeavePage(1); }} className="w-full p-2 bg-white border border-slate-300 rounded-lg font-semibold">
                      {leaveMonths.map(m => <option key={m} value={m}>{m}</option>)}
                    </select>
                  </div>
                  <div className="flex items-center space-x-2">
                    <span className="text-slate-500 shrink-0 font-bold">狀態</span>
                    <select value={leaveFilterStatus} onChange={(e) => { setLeaveFilterStatus(e.target.value); setLeavePage(1); }} className="w-full p-2 bg-white border border-slate-300 rounded-lg font-semibold">
                      <option value="全部">全部狀態</option>
                      <option value="待審核">待審核</option>
                      <option value="核准">已核准</option>
                      <option value="已退回">已退回</option>
                    </select>
                  </div>
                </div>

                {/* 請假清冊 */}
                <div className="divide-y divide-slate-100">
                  {paginatedLeaves.length === 0 ? (
                    <div className="p-8 text-center text-slate-400 text-sm font-medium">目前尚無任何符合條件之同仁請假紀錄</div>
                  ) : (
                    paginatedLeaves.map((req: any) => {
                      const isSynced = syncedRequests.includes(req.id);
                      return (
                        <div key={req.id} className="p-4 sm:p-5 hover:bg-slate-50/50 transition duration-150 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                          <div className="space-y-1">
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-slate-800 text-base">{req.applicant}</span>
                              <span className="text-xs bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded-full font-bold">{employees.find(e => e.name === req.applicant)?.group}</span>
                              <span className="text-xs bg-teal-50 text-teal-800 px-2.5 py-0.5 rounded-full font-bold border border-teal-100">
                                {req.leaveType} · {req.hours} 小時
                              </span>
                            </div>
                            <p className="text-xs text-slate-600 font-semibold">
                              日期區間：{req.startDate} 至 {req.endDate} ({req.startTime} 至 {req.endTime}) · 職務代理同仁：<strong className="text-indigo-800">{req.agent}</strong>
                            </p>
                            {req.leaveType === '公假' && req.location && (
                              <div className="text-[11px] flex items-center gap-2 mt-1">
                                <span className="bg-amber-50 text-amber-800 border border-amber-100 px-2 py-0.5 rounded font-bold">起訖：{req.location}</span>
                                {req.isBusinessTrip && <span className="bg-emerald-50 text-emerald-800 border border-emerald-100 px-2 py-0.5 rounded font-bold animate-pulse">✓ 具出差性質可核銷</span>}
                              </div>
                            )}
                            <p className="text-xs italic text-slate-400 mt-1">請假詳細事由：{req.reason || '未填寫'}</p>
                          </div>

                          <div className="flex items-center gap-2 flex-wrap self-end sm:self-center">
                            <span className={`text-xs px-2.5 py-1 rounded-full font-bold ${req.status === '核准' ? 'bg-emerald-100 text-emerald-800' : req.status === '待審核' ? 'bg-amber-100 text-amber-800 animate-pulse' : 'bg-rose-100 text-rose-800'}`}>
                              {req.status}
                            </span>
                            {req.status === '核准' && (
                              <button onClick={() => triggerCalendarSync(req, getGoogleCalendarUrl(req))} className={`text-xs px-3 py-1.5 rounded-lg border font-bold flex items-center space-x-1 ${isSynced ? 'bg-slate-700 text-white border-transparent' : 'bg-white text-indigo-600 border-indigo-200 hover:bg-indigo-50'}`}>
                                <Calendar className="h-3.5 w-3.5" />
                                <span>{isSynced ? '已同步行事曆' : '同步 Google 行事曆'}</span>
                              </button>
                            )}
                            
                            {/* 🔗 跨系統整合：公假且核准且可領旅費者一鍵智慧同步 */}
                            {req.status === '核准' && req.leaveType === '公假' && req.isBusinessTrip && (
                              <button 
                                onClick={() => handleLeaveCarryOverToTravel(req)} 
                                className="text-xs bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 px-3 py-1.5 rounded-lg font-bold flex items-center space-x-1 transition duration-150 shadow-sm"
                              >
                                <ExternalLink className="h-3.5 w-3.5" />
                                <span>🔗 智慧同步差旅</span>
                              </button>
                            )}

                            <button onClick={() => setDeleteTarget({ type: 'leave', id: req.id, name: `${req.applicant} 的 ${req.leaveType}` })} className="p-1.5 text-rose-500 hover:bg-rose-50 border border-rose-100 rounded-lg">
                              <Trash2 className="h-4 w-4" />
                            </button>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>

                {/* 請假分頁 */}
                <div className="p-4 border-t border-slate-100 flex justify-between items-center text-xs text-slate-500 font-medium">
                  <span>共 {filteredLeaves.length} 筆，顯示第 {Math.min(filteredLeaves.length, (leavePage - 1) * recordsPerPage + 1)} - {Math.min(filteredLeaves.length, leavePage * recordsPerPage)} 筆</span>
                  <div className="flex items-center space-x-2">
                    <button disabled={leavePage === 1} onClick={() => setLeavePage(p => p - 1)} className="p-1.5 border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40">
                      <ChevronLeft className="h-4 w-4" />
                    </button>
                    <span>頁次 {leavePage} / {Math.ceil(filteredLeaves.length / recordsPerPage) || 1}</span>
                    <button disabled={leavePage >= Math.ceil(filteredLeaves.length / recordsPerPage)} onClick={() => setLeavePage(p => p + 1)} className="p-1.5 border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40">
                      <ChevronRight className="h-4 w-4" />
                    </button>
                  </div>
                </div>
              </div>

              {/* 假日加班補休協商 */}
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div className="p-5 border-b border-slate-100 bg-slate-50 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                  <div>
                    <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                      <Calendar className="text-amber-600 h-5 w-5" />
                      <span>假日加班與補休協商歷史紀錄</span>
                    </h3>
                    <p className="text-xs text-slate-500">團體加班協商。一經核定，加班時數將自動累積至同仁補休，並提供勞資協商同意書 A4 列印</p>
                  </div>
                  <button 
                    onClick={handleExportOvertimeExcel}
                    className="bg-amber-600 hover:bg-amber-700 text-white text-xs px-4 py-2.5 rounded-lg font-bold flex items-center space-x-2 shadow-sm transition"
                  >
                    <FileSpreadsheet className="h-4 w-4" />
                    <span>📊 匯出補休協商登記簿 (對帳對休用)</span>
                  </button>
                </div>

                {/* 加班協商篩選與搜尋 */}
                <div className="p-4 bg-slate-50 border-b border-slate-100 grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
                  <div className="relative">
                    <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-slate-400" />
                    <input
                      type="text"
                      placeholder="搜尋事由、活動名稱或參與名單..."
                      value={otSearch}
                      onChange={(e) => { setOtSearch(e.target.value); setOtPage(1); }}
                      className="pl-8 w-full p-2 bg-white border border-slate-300 rounded-lg outline-none font-medium"
                    />
                  </div>
                  <div className="flex items-center space-x-2">
                    <span className="text-slate-500 shrink-0 font-bold">年度</span>
                    <select value={otFilterYear} onChange={(e) => { setOtFilterYear(e.target.value); setOtPage(1); }} className="w-full p-2 bg-white border border-slate-300 rounded-lg font-semibold">
                      {otYears.map(y => <option key={y} value={y}>{y}</option>)}
                    </select>
                  </div>
                  <div className="flex items-center space-x-2">
                    <span className="text-slate-500 shrink-0 font-bold">月份</span>
                    <select value={otFilterMonth} onChange={(e) => { setOtFilterMonth(e.target.value); setOtPage(1); }} className="w-full p-2 bg-white border border-slate-300 rounded-lg font-semibold">
                      {otMonths.map(m => <option key={m} value={m}>{m}</option>)}
                    </select>
                  </div>
                  <div className="flex items-center space-x-2">
                    <span className="text-slate-500 shrink-0 font-bold">狀態</span>
                    <select value={otFilterStatus} onChange={(e) => { setOtFilterStatus(e.target.value); setOtPage(1); }} className="w-full p-2 bg-white border border-slate-300 rounded-lg font-semibold">
                      <option value="全部">全部狀態</option>
                      <option value="待審核">待審核</option>
                      <option value="已核准">已核准</option>
                      <option value="已退回">已退回</option>
                    </select>
                  </div>
                </div>

                {/* 加班列表 */}
                <div className="divide-y divide-slate-100">
                  {paginatedOvertimes.length === 0 ? (
                    <div className="p-8 text-center text-slate-400 text-sm font-medium">目前無假日加班補休協商紀錄</div>
                  ) : (
                    paginatedOvertimes.map((ot: any) => {
                      const isOtSynced = syncedRequests.includes(ot.id);
                      return (
                        <div key={ot.id} className="p-4 sm:p-5 hover:bg-slate-50/50 transition duration-150 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                          <div className="space-y-1">
                            <p className="font-bold text-slate-800 text-base flex items-center gap-2">
                              <span>{ot.activityName}</span>
                              <span className="bg-amber-50 text-amber-800 px-2.5 py-0.5 rounded text-xs border border-amber-200 font-bold">
                                補休計 +{ot.hours} 小時
                              </span>
                            </p>
                            <p className="text-xs text-slate-600 font-semibold">
                              日期：民國 {formatMinguoDateText(ot.workDate)} ({getWeekdayText(ot.workDate)}) 🕰️ 時段：{ot.startTime} 至 {ot.endTime}
                            </p>
                            <p className="text-xs text-slate-500">
                              申報出勤名冊：<strong className="text-teal-700">{ot.participants ? ot.participants.join('、') : ''}</strong>
                            </p>
                          </div>

                          <div className="flex items-center gap-2 flex-wrap self-end sm:self-center">
                            <span className={`text-xs px-2.5 py-1 rounded-full font-bold ${ot.status === '已核准' ? 'bg-emerald-100 text-emerald-800' : ot.status === '待審核' ? 'bg-amber-100 text-amber-800 animate-pulse' : 'bg-rose-100 text-rose-800'}`}>
                              {ot.status}
                            </span>
                            
                            {ot.status === '已核准' && (
                              <button 
                                onClick={() => triggerCalendarSync(ot, getGoogleCalendarUrl({
                                  startDate: ot.workDate,
                                  endDate: ot.workDate,
                                  startTime: ot.startTime,
                                  endTime: ot.endTime,
                                  leaveType: '補休加班',
                                  applicant: ot.participants ? ot.participants.join('、') : '',
                                  agent: '無',
                                  reason: ot.activityName
                                }))} 
                                className={`text-xs px-2.5 py-1.5 rounded-lg border font-bold flex items-center space-x-1 ${isOtSynced ? 'bg-slate-700 text-white' : 'bg-white text-indigo-600 border-indigo-200 hover:bg-indigo-50'}`}
                              >
                                <Calendar className="h-3.5 w-3.5" />
                                <span>{isOtSynced ? '已同步' : '同步行事曆'}</span>
                              </button>
                            )}

                            {ot.status === '已核准' && (
                              <button onClick={() => { setSelectedOtRequest(ot); setShowAgreementModal(true); }} className="text-xs text-slate-700 hover:text-white border border-slate-200 bg-slate-50 hover:bg-slate-700 px-3 py-1.5 rounded-lg flex items-center space-x-1 font-bold">
                                <Printer className="h-3.5 w-3.5" />
                                <span>列印同意書</span>
                              </button>
                            )}

                            {ot.status === '已核准' && (
                              <button onClick={() => handleCarryOverToTravel(ot)} className="text-xs bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 px-3 py-1.5 rounded-lg font-bold flex items-center space-x-1 transition duration-150 shadow-sm">
                                <ExternalLink className="h-3.5 w-3.5" />
                                <span>🔗 智慧同步差旅</span>
                              </button>
                            )}

                            <button onClick={() => setDeleteTarget({ type: 'overtime', id: ot.id, name: ot.activityName })} className="p-1.5 text-rose-500 hover:bg-rose-50 border border-rose-100 rounded-lg">
                              <Trash2 className="h-4 w-4" />
                            </button>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>

                {/* 加班分頁 */}
                <div className="p-4 border-t border-slate-100 flex justify-between items-center text-xs text-slate-500 font-medium">
                  <span>共 {filteredOvertimes.length} 筆，顯示第 {Math.min(filteredOvertimes.length, (otPage - 1) * recordsPerPage + 1)} - {Math.min(filteredOvertimes.length, otPage * recordsPerPage)} 筆</span>
                  <div className="flex items-center space-x-2">
                    <button disabled={otPage === 1} onClick={() => setOtPage(p => p - 1)} className="p-1.5 border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40">
                      <ChevronLeft className="h-4 w-4" />
                    </button>
                    <span>頁次 {otPage} / {Math.ceil(filteredOvertimes.length / recordsPerPage) || 1}</span>
                    <button disabled={otPage >= Math.ceil(filteredOvertimes.length / recordsPerPage)} onClick={() => setOtPage(p => p + 1)} className="p-1.5 border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40">
                      <ChevronRight className="h-4 w-4" />
                    </button>
                  </div>
                </div>
              </div>

            </div>
          )}

          {/* ================================================== */}
          {/* B. 線上請假/加班登記表 */}
          {/* ================================================== */}
          {activeTab === 'attendance_apply' && (
            <div className="max-w-3xl mx-auto space-y-6">
              <div className="flex bg-slate-200 p-1 rounded-xl">
                <button type="button" onClick={() => setApplyType('leave')} className={`flex-1 py-2.5 text-center text-sm font-bold rounded-lg transition ${applyType === 'leave' ? 'bg-indigo-600 text-white shadow' : 'text-slate-600 hover:text-slate-900'}`}>
                  🎒 數辦同仁請假登記表
                </button>
                <button type="button" onClick={() => setApplyType('overtime')} className={`flex-1 py-2.5 text-center text-sm font-bold rounded-lg transition ${applyType === 'overtime' ? 'bg-amber-600 text-white shadow' : 'text-slate-600 hover:text-slate-900'}`}>
                  📅 假日加班/研習團體申報
                </button>
              </div>

              {applyType === 'leave' ? (
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden animate-fadeIn">
                  <div className="px-6 py-4 bg-indigo-700 text-white">
                    <h3 className="font-bold text-base">人事線上請假登記表 (含代理與差旅判定)</h3>
                    <p className="text-xs text-indigo-100 mt-1">請填妥假別與代理同仁，公假部分可一鍵連動差旅請領資格</p>
                  </div>
                  <form onSubmit={handleApplyLeave} className="p-6 space-y-4 text-sm">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">請假同仁 <span className="text-rose-500">*</span></label>
                        <select value={formApplicant} onChange={(e) => setFormApplicant(e.target.value)} className="w-full p-2.5 bg-white border border-slate-300 rounded-lg font-semibold">
                          <option value="">-- 請選擇人員 --</option>
                          {employees.map(e => <option key={e.name} value={e.name}>{e.name} ({e.group})</option>)}
                        </select>
                      </div>
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">假別選擇 <span className="text-rose-500">*</span></label>
                        <select value={formLeaveType} onChange={(e) => setFormLeaveType(e.target.value)} className="w-full p-2.5 bg-white border border-slate-300 rounded-lg font-bold text-indigo-700">
                          <option value="特休">特休 (經核准後將自動核扣同仁額度)</option>
                          <option value="補休">補休 (經核准後將自動核扣同仁額度)</option>
                          <option value="事假">事假 (年度事假總累計)</option>
                          <option value="病假">病假 (年度病假總累計)</option>
                          <option value="喪假">喪假 (年度喪假總累計)</option>
                          <option value="公假">公假 (得勾選是否申報出差費)</option>
                        </select>
                      </div>
                    </div>

                    <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">開始日期與時間</label>
                        <div className="flex gap-2">
                          <input type="date" value={formStartDate} onChange={(e) => setFormStartDate(e.target.value)} className="p-2 border rounded-lg w-3/5 font-semibold" required />
                          <input type="time" value={formStartTime} onChange={(e) => setFormStartTime(e.target.value)} className="p-2 border rounded-lg w-2/5 font-semibold" required />
                        </div>
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">結束日期與時間</label>
                        <div className="flex gap-2">
                          <input type="date" value={formEndDate} onChange={(e) => setFormEndDate(e.target.value)} className="p-2 border rounded-lg w-3/5 font-semibold" required />
                          <input type="time" value={formEndTime} onChange={(e) => setFormEndTime(e.target.value)} className="p-2 border rounded-lg w-2/5 font-semibold" required />
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">本次請假時數 (小時)</label>
                        <input type="number" step="0.5" value={formHours} onChange={(e) => setFormHours(parseFloat(e.target.value) || 0)} className="w-32 p-2.5 border border-slate-300 rounded-lg font-bold text-indigo-700 text-center text-lg focus:ring-2 focus:ring-indigo-500" />
                      </div>
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">職務代理同仁 <span className="text-rose-500">*</span></label>
                        <select value={formAgent} onChange={(e) => setFormAgent(e.target.value)} className="w-full p-2.5 bg-white border border-slate-300 rounded-lg font-semibold">
                          <option value="">-- 請選擇同辦公室代理人 --</option>
                          {employees.filter(e => e.name !== formApplicant).map(e => <option key={e.name} value={e.name}>{e.name} ({e.group})</option>)}
                        </select>
                      </div>
                    </div>

                    {formLeaveType === '公假' && (
                      <div className="bg-amber-50/50 p-4 rounded-xl border border-amber-200 space-y-3">
                        <label className="block text-xs font-bold text-amber-950">公出起訖地點 (預設為：越港國小)</label>
                        <input type="text" value={formLocation} onChange={(e) => setFormLocation(e.target.value)} className="w-full p-2.5 border border-amber-300 rounded-lg bg-white font-semibold" />
                        <div className="flex items-center space-x-2">
                          <input type="checkbox" id="isBusinessTrip" checked={formIsBusinessTrip} onChange={(e) => setFormIsBusinessTrip(e.target.checked)} className="rounded text-amber-600 focus:ring-amber-500 h-4 w-4 cursor-pointer" />
                          <label htmlFor="isBusinessTrip" className="text-xs text-amber-900 font-bold select-none cursor-pointer">此公假具備出差性質（核定後准予前往差旅費申報系統填寫）</label>
                        </div>
                      </div>
                    )}

                    <div>
                      <label className="block font-bold text-slate-700 mb-1">請假事由或詳細原因說明</label>
                      <textarea rows={3} value={formReason} onChange={(e) => setFormReason(e.target.value)} className="w-full p-2.5 border border-slate-300 rounded-lg outline-none font-medium focus:ring-2 focus:ring-indigo-500" placeholder="填寫詳細公務事由或請假原因，有利於主管快速審核..." />
                    </div>

                    <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 px-4 rounded-xl flex items-center justify-center space-x-2 shadow-sm transition">
                      <Plus className="h-4 w-4" />
                      <span>送出請假申請單</span>
                    </button>
                  </form>
                </div>
              ) : (
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden animate-fadeIn">
                  <div className="px-6 py-4 bg-amber-600 text-white">
                    <h3 className="font-bold text-base">休息日工作加班補休申報 (團體名冊協商)</h3>
                    <p className="text-xs text-amber-100 mt-1">針對假日辦理全縣研習活動、增能會議，支持行政人員團體協商與大寫補休天數產製</p>
                  </div>
                  <form onSubmit={handleApplyOvertime} className="p-6 space-y-5 text-sm">
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">加班出勤日期</label>
                        <input type="date" value={otDate} onChange={(e) => setOtDate(e.target.value)} className="w-full p-2.5 border rounded-lg font-semibold bg-white" required />
                      </div>
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">開始時間</label>
                        <input type="time" value={otStartTime} onChange={(e) => setOtStartTime(e.target.value)} className="w-full p-2.5 border rounded-lg font-semibold bg-white" required />
                      </div>
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">結束時間</label>
                        <input type="time" value={otEndTime} onChange={(e) => setOtEndTime(e.target.value)} className="w-full p-2.5 border rounded-lg font-semibold bg-white" required />
                      </div>
                    </div>

                    <div>
                      <label className="block font-bold text-slate-700 mb-2">出勤同仁勾選 (可團體複選)</label>
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 p-3 bg-slate-50 border rounded-xl">
                        {employees.map(emp => (
                          <label key={emp.name} className={`flex items-center space-x-2 p-2 border rounded-lg cursor-pointer bg-white transition ${otParticipants.includes(emp.name) ? 'border-amber-500 bg-amber-50 font-bold text-amber-900 shadow-sm' : 'text-slate-600 hover:bg-slate-50'}`}>
                            <input type="checkbox" checked={otParticipants.includes(emp.name)} onChange={() => toggleOtParticipant(emp.name)} className="rounded border-slate-300 text-amber-600 focus:ring-amber-500 h-4 w-4" />
                            <span>{emp.name}</span>
                          </label>
                        ))}
                      </div>
                    </div>

                    <div>
                      <label className="block font-bold text-slate-700 mb-1">研習研討會名稱 / 加班詳細事由</label>
                      <input type="text" value={otActivityName} onChange={(e) => setOtActivityName(e.target.value)} placeholder="例如：辦理 115 年度全縣數位素養宣講增能研習" className="w-full p-2.5 border rounded-lg font-semibold outline-none focus:ring-2 focus:ring-amber-500" required />
                    </div>

                    <button type="submit" className="w-full bg-amber-600 hover:bg-amber-700 text-white font-bold py-3 rounded-xl flex items-center justify-center space-x-2 transition shadow-sm">
                      <Plus className="h-4 w-4" />
                      <span>建立團體協商同意案並送出</span>
                    </button>
                  </form>
                </div>
              )}
            </div>
          )}

          {/* ================================================== */}
          {/* C. 同仁假別細算表 */}
          {/* ================================================== */}
          {activeTab === 'attendance_roster' && (
            <div className="space-y-6">
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden animate-fadeIn">
                <div className="p-5 border-b border-slate-100 bg-slate-50 flex flex-col sm:flex-row justify-between items-center gap-3">
                  <div>
                    <h3 className="font-bold text-lg text-slate-800">數辦同仁出缺勤與特休/補休細算表 (115學年度最新狀態)</h3>
                    <p className="text-xs text-slate-500">名冊已依照會計及勞政主管機關要求之「行政人員 ➔ 資訊人員 ➔ 輔導人員」格式精準排序</p>
                  </div>
                  <button onClick={handleExportCSV} className="bg-teal-700 hover:bg-teal-800 text-white text-xs px-4 py-2.5 rounded-lg font-bold flex items-center space-x-1.5 shadow-sm transition font-sans">
                    <Download className="h-4 w-4" />
                    <span>下載同仁差勤特補休總清冊 CSV</span>
                  </button>
                </div>

                {/* 搜尋與篩選 */}
                <div className="p-4 bg-slate-50 border-b border-slate-100 grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
                  <div className="relative">
                    <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-slate-400" />
                    <input
                      type="text"
                      placeholder="搜尋同仁姓名、職稱、或組別..."
                      value={rosterSearch}
                      onChange={(e) => setRosterSearch(e.target.value)}
                      className="pl-8 w-full p-2 bg-white border border-slate-300 rounded-lg outline-none font-semibold"
                    />
                  </div>
                  <div className="flex items-center space-x-2">
                    <span className="text-slate-500 shrink-0 font-bold">小組篩選</span>
                    <select 
                      value={selectedGroup} 
                      onChange={(e) => setSelectedGroup(e.target.value)} 
                      className="w-full p-2 bg-white border border-slate-300 rounded-lg font-semibold cursor-pointer"
                    >
                      <option value="全部">全部小組</option>
                      <option value="中區小組">中區小組</option>
                      <option value="東區小組">東區小組</option>
                      <option value="北區小組">北區小組</option>
                      <option value="西區小組">西區小組</option>
                      <option value="南區小組">南區小組</option>
                      <option value="行政">行政人員</option>
                    </select>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="min-w-full divide-y divide-slate-200 text-left text-sm font-medium">
                    <thead className="bg-slate-50 text-xs font-bold text-slate-500 uppercase tracking-wider">
                      <tr>
                        <th className="px-4 py-3">同仁基本資訊</th>
                        <th className="px-4 py-3">到職日及年資</th>
                        <th className="px-4 py-3 bg-indigo-50/50 text-indigo-900 font-bold">賸餘特休 (天/時)</th>
                        <th className="px-4 py-3 bg-teal-50/50 text-teal-900 font-bold">賸餘補休 (天/時)</th>
                        <th className="px-4 py-3 text-slate-600">已請事假</th>
                        <th className="px-4 py-3 text-slate-600">已請病假</th>
                        <th className="px-4 py-3 text-slate-600">已請喪假</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 bg-white text-slate-700">
                      {sortedAndFilteredEmployees.map((emp) => {
                        const sInfo = calculateSeniorityInfo(emp.hireDate);
                        const isRetired = emp.name === '新進人員(待定)';
                        return (
                          <tr key={emp.name} className={`hover:bg-slate-50/70 transition ${isRetired ? 'bg-slate-100/50 text-slate-400' : ''}`}>
                            <td className="px-4 py-4">
                              <div className="font-bold text-slate-900 text-base">{emp.name}</div>
                              <div className="text-xs text-slate-500 mt-0.5">{emp.group} · {emp.title}</div>
                            </td>
                            <td className="px-4 py-4">
                              <div className="text-xs text-slate-400">{emp.hireDate ? `到職: ${emp.hireDate}` : '人員待聘中'}</div>
                              <div className="text-sm font-semibold text-slate-700 mt-0.5">{sInfo.text}</div>
                            </td>
                            <td className="px-4 py-4 bg-indigo-50/10 font-bold text-indigo-700 text-base">
                              {emp.remainingTe.d}天 {emp.remainingTe.h}時
                            </td>
                            <td className="px-4 py-4 bg-teal-50/10 font-bold text-teal-800 text-base">
                              {emp.remainingBu.d}天 {emp.remainingBu.h}時
                              <div className="text-[10px] text-teal-500 font-normal mt-0.5">累計剩餘 {emp.remainingBu.d * 8 + emp.remainingBu.h} 小時</div>
                            </td>
                            <td className="px-4 py-4 text-slate-600">{emp.takenShi.d}天 {emp.takenShi.h}時</td>
                            <td className="px-4 py-4 text-slate-600">{emp.takenBing.d}天 {emp.takenBing.h}時</td>
                            <td className="px-4 py-4 text-slate-600">{emp.takenSang.d}天 {emp.takenSang.h}時</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ================================================== */}
          {/* D. 主管審核中心 */}
          {/* ================================================== */}
          {activeTab === 'attendance_audit' && (
            <div className="space-y-6">
              {(!isUnlocked && currentUser === '主管端') ? (
                <div className="max-w-md mx-auto bg-white rounded-xl border border-slate-200 shadow p-8 text-center space-y-4 animate-fadeIn">
                  <Lock className="h-12 w-12 text-slate-400 mx-auto" />
                  <h3 className="font-bold text-slate-900 text-lg">主管審核身份驗證鎖</h3>
                  <p className="text-xs text-slate-500">簽核請假核扣時數與核准補休加總，請輸入數辦密碼解鎖權限。</p>
                  <form onSubmit={handleVerifyPassword} className="space-y-3">
                    <input type="password" placeholder="請輸入審核驗證密碼 (預設：115)" value={inputPassword} onChange={(e) => setInputPassword(e.target.value)} className="w-full p-2.5 border rounded-lg text-center font-bold tracking-widest outline-none focus:ring-1 focus:ring-indigo-500 text-sm" autoFocus />
                    <button type="submit" className="w-full bg-slate-900 hover:bg-black text-white font-bold p-2.5 rounded-lg text-xs tracking-wider">身份解鎖</button>
                  </form>
                </div>
              ) : currentUser !== '主管端' ? (
                <div className="max-w-md mx-auto bg-white rounded-xl border border-slate-200 p-8 text-center space-y-4 shadow-sm animate-fadeIn">
                  <AlertCircle className="h-12 w-12 text-amber-500 mx-auto animate-bounce" />
                  <h3 className="font-bold text-slate-900">目前處於同仁自填申報模式</h3>
                  <p className="text-xs text-slate-500 leading-relaxed">同仁權限僅限於填寫並送出表單。請於右上方帳戶切換為「主管審核端」並解鎖以進行假單審核與核銷操作。</p>
                </div>
              ) : (
                <div className="space-y-6 animate-fadeIn">
                  {/* 加班協商審定 */}
                  <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                    <div className="px-6 py-4 bg-amber-600 text-white flex justify-between items-center">
                      <h4 className="font-bold text-sm sm:text-base">假日加班協商 補休授權核准區</h4>
                      <span className="text-[10px] bg-white/20 px-2 py-0.5 rounded border border-white/30 font-bold">主管解鎖成功 🔓</span>
                    </div>
                    <div className="p-6 divide-y divide-slate-100">
                      {overtimeRequests.filter((r: any) => r.status === '待審核').length === 0 ? (
                        <p className="text-slate-400 text-center py-4 text-xs font-semibold">目前無待審假日加班申報案件</p>
                      ) : (
                        overtimeRequests.filter((r: any) => r.status === '待審核').map((ot: any) => (
                          <div key={ot.id} className="py-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                            <div className="space-y-1">
                              <p className="font-bold text-slate-900 text-base">{ot.activityName} (+{ot.hours}小時)</p>
                              <p className="text-xs text-slate-500 font-semibold">加班日：民國 {formatMinguoDateText(ot.workDate)} ({getWeekdayText(ot.workDate)})</p>
                              <p className="text-xs text-slate-500">申報出勤同仁：<strong>{ot.participants.join('、')}</strong></p>
                            </div>
                            <div className="flex gap-2">
                              <button onClick={() => handleRejectOvertime(ot.id)} className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-xs font-bold border border-rose-200 flex items-center space-x-1">
                                <XCircle className="h-3.5 w-3.5" />
                                <span>退回</span>
                              </button>
                              <button onClick={() => handleApproveOvertime(ot)} className="px-4 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold shadow flex items-center space-x-1 transition">
                                <Check className="h-3.5 w-3.5" />
                                <span>簽核同意補休</span>
                              </button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>

                  {/* 請假審查 */}
                  <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                    <div className="px-6 py-4 bg-indigo-700 text-white">
                      <h4 className="font-bold text-sm sm:text-base">辦公室同仁請假 簽准與核扣區</h4>
                    </div>
                    <div className="p-6 divide-y divide-slate-100">
                      {requests.filter((r: any) => r.status === '待審核').length === 0 ? (
                        <p className="text-slate-400 text-center py-4 text-xs font-semibold">目前無待簽核同仁假單</p>
                      ) : (
                        requests.filter((r: any) => r.status === '待審核').map((req: any) => (
                          <div key={req.id} className="py-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                            <div className="space-y-1">
                              <p className="font-bold text-slate-900 text-base">{req.applicant} · {req.leaveType} ({req.hours}小時)</p>
                              <p className="text-xs text-slate-500 font-semibold">請假日期：{req.startDate} 至 {req.endDate} (代理同仁: {req.agent})</p>
                              <p className="text-xs text-slate-500 italic">事由：{req.reason || '未說明'}</p>
                            </div>
                            <div className="flex gap-2">
                              <button onClick={() => handleReject(req.id)} className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-xs font-bold border border-rose-200 flex items-center space-x-1">
                                <XCircle className="h-3.5 w-3.5" />
                                <span>退回</span>
                              </button>
                              <button onClick={() => handleApproveLeave(req.id)} className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-bold shadow flex items-center space-x-1 transition">
                                <Check className="h-3.5 w-3.5" />
                                <span>批准並自動扣抵時數</span>
                              </button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ================================================== */}
          {/* E. 差旅費填報核心 */}
          {/* ================================================== */}
          {activeTab === 'calc' && (
            <div className="max-w-4xl mx-auto space-y-6">
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden animate-fadeIn">
                <div className="px-6 py-4 bg-slate-950 text-white">
                  <h3 className="font-bold text-base">會計核銷規格 · 差旅費智慧填報模組</h3>
                  <p className="text-xs text-slate-400">支援單程未滿5公里篩抵、一秒定位。填妥後按會計標準產製彙整分頁Excel</p>
                </div>
                
                <form onSubmit={handleCalculateTravel} className="p-6 space-y-5 text-sm">
                  {/* 人員勾選 */}
                  <div>
                    <label className="block font-bold text-slate-700 mb-2">出差同仁選擇 (可複選，已按行政、資訊、輔導同仁排序)</label>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-slate-50 p-3 border border-slate-200 rounded-xl">
                      {employees
                        .sort((a, b) => (ROLE_SORT_ORDER[a.title] || 99) - (ROLE_SORT_ORDER[b.title] || 99))
                        .map(emp => {
                          const isSelected = formData.memberNames.includes(emp.name);
                          return (
                            <button
                              key={emp.name} 
                              type="button" 
                              onClick={() => toggleMemberSelection(emp.name)}
                              className={`flex items-center space-x-2 p-2.5 border rounded-lg cursor-pointer bg-white transition ${isSelected ? 'border-emerald-500 bg-emerald-50 font-bold text-emerald-950 shadow-sm' : 'text-slate-600 hover:bg-slate-50'}`}
                            >
                              {isSelected ? <CheckSquare className="w-4 h-4 text-emerald-600" /> : <Square className="w-4 h-4 text-slate-400" />}
                              <span>{emp.name}</span>
                            </button>
                          );
                        })}
                    </div>
                  </div>

                  {/* 出差天數 */}
                  <div className="bg-indigo-50/50 p-4 rounded-xl border border-indigo-100">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input 
                        type="checkbox" 
                        name="isMultiDay" 
                        checked={formData.isMultiDay || false} 
                        onChange={handleInputChange} 
                        className="w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500 cursor-pointer" 
                      />
                      <span className="font-bold text-indigo-950">這是一次跨天數出差 (兩天含以上多日公出)</span>
                    </label>
                    
                    {formData.isMultiDay && (
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-3 animate-in slide-in-from-top-2 duration-200">
                        <div>
                          <label className="block text-xs font-semibold text-indigo-700 mb-1">出差結束日期</label>
                          <input 
                            type="date" 
                            name="endDate" 
                            min={formData.date}
                            value={formData.endDate || formData.date} 
                            onChange={handleInputChange} 
                            className="w-full p-2 bg-white border border-indigo-200 rounded-lg text-sm outline-none font-bold" 
                          />
                        </div>
                        <div className="flex items-end">
                          <span className="text-xs text-indigo-600 pb-2">
                            系統自動估計：<strong className="text-sm font-bold text-indigo-800">{calculateFees(formData).days}</strong> 天
                            <br />
                            <span className="text-[10px] text-indigo-500 font-normal">(去程日與返程日採請假時段，中間天數全額申發膳雜費)</span>
                          </span>
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">出差首日日期</label>
                      <input type="date" name="date" required value={formData.date} onChange={handleInputChange} className="w-full p-2.5 bg-white border border-gray-300 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 font-bold" />
                    </div>
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">交通工具</label>
                      <select name="transportMode" value={formData.transportMode} onChange={handleInputChange} className="w-full p-2.5 bg-white border border-gray-300 rounded-lg outline-none font-bold text-indigo-700">
                        <option value="car">汽車來回 (3元/單程公里 補償去回程)</option>
                        <option value="scooter">機車來回 (2元/單程公里 補償去回程)</option>
                        <option value="hsr">台灣高鐵 (越港國小基準 雲林站出發)</option>
                      </select>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">出發地點</label>
                      <input type="text" name="startLocation" required value={formData.startLocation} onChange={handleInputChange} className="w-full p-2.5 bg-white border border-gray-300 rounded-lg outline-none font-semibold" />
                    </div>
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">公出目的地及受訪學校 (支援下拉與手填)</label>
                      <div className="flex gap-2">
                        <input 
                          type="text" 
                          name="destination"
                          required 
                          value={formData.destination} 
                          onChange={handleInputChange} 
                          list="destinations-list"
                          placeholder="請手填或點選學校名稱..."
                          className="flex-1 p-2.5 bg-white border border-gray-300 rounded-lg outline-none w-full font-semibold" 
                        />
                        <datalist id="destinations-list">
                          {COMMON_LOCATIONS.map(l => <option key={l.name} value={l.name}>{l.address}</option>)}
                        </datalist>
                      </div>
                    </div>
                  </div>

                  {formData.transportMode !== 'hsr' && (
                    <div className="bg-slate-50 p-4 border rounded-xl space-y-3">
                      <div className="flex justify-between items-center">
                        <span className="font-bold text-slate-700 text-xs">公務里程智慧計算</span>
                        <div className="flex items-center space-x-1">
                          <input type="checkbox" id="isRoundTrip" name="isRoundTrip" checked={formData.isRoundTrip} onChange={handleInputChange} className="rounded" />
                          <label htmlFor="isRoundTrip" className="text-xs text-slate-600 font-bold cursor-pointer">按往返計算 (單程公里數x2)</label>
                        </div>
                      </div>
                      <div className="flex gap-2">
                        <input 
                          ref={distanceInputRef}
                          type="number" 
                          name="distance" 
                          min="0" 
                          step="0.1" 
                          placeholder="請點擊右方自動計算里程..." 
                          value={formData.distance} 
                          onChange={handleInputChange} 
                          required 
                          readOnly={!isDistanceManual} 
                          className={`flex-1 p-2.5 border rounded-lg font-bold text-center ${isDistanceManual ? 'bg-white border-gray-300' : 'bg-gray-100 border-gray-200 text-gray-500'}`} 
                        />
                        {!isDistanceManual && <button type="button" onClick={() => setIsDistanceManual(true)} className="px-3 py-2 bg-white text-gray-700 border border-slate-300 rounded-lg hover:bg-slate-50 text-xs font-bold">手動輸入</button>}
                        <button type="button" onClick={handleAutoFetchDistance} disabled={isProcessing} className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 text-xs font-bold flex items-center space-x-1.5 transition">
                          <MapPin className="h-3.5 w-3.5" />
                          <span>{isProcessing ? '測距中...' : '📍 定位導航測距'}</span>
                        </button>
                      </div>
                    </div>
                  )}

                  {/* 智慧型高鐵票價 */}
                  {formData.transportMode === 'hsr' && (
                    <div className="space-y-4 bg-indigo-50/40 p-4 rounded-xl border border-indigo-100 animate-fadeIn">
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1">去程到達高鐵站 (雲林至...)</label>
                          <select name="hsrStation" value={formData.hsrStation} onChange={handleInputChange} className="w-full p-2.5 bg-white border border-gray-300 rounded-lg text-sm font-bold text-indigo-950 cursor-pointer">
                            {Object.keys(HSR_FARES_FROM_YUNLIN).map(station => (<option key={station} value={station}>{station}站</option>))}
                          </select>
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1">回程起點高鐵站 (...至雲林)</label>
                          <select 
                            name="hsrStationReturn" 
                            value={formData.isHsrSameStation !== false ? formData.hsrStation : (formData.hsrStationReturn || formData.hsrStation)} 
                            onChange={handleInputChange} 
                            disabled={formData.isHsrSameStation !== false} 
                            className="w-full p-2.5 bg-white border border-gray-300 rounded-lg text-sm font-bold text-indigo-950 disabled:opacity-50 disabled:bg-gray-100 cursor-pointer"
                          >
                            {Object.keys(HSR_FARES_FROM_YUNLIN).map(station => (<option key={station} value={station}>{station}站</option>))}
                          </select>
                        </div>
                      </div>
                      
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 border-t border-indigo-100/60 pt-3">
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1">去程票種選擇</label>
                          <div className="flex gap-4 p-1">
                            <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700">
                              <input 
                                type="radio" 
                                name="hsrTicketTypeOutbound" 
                                value="standard" 
                                checked={(formData.hsrTicketTypeOutbound || formData.hsrTicketType || 'standard') === 'standard'} 
                                onChange={handleInputChange}
                                className="text-indigo-600 focus:ring-indigo-500 h-4 w-4" 
                              />
                              對號全票 (實報實銷)
                            </label>
                            <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700">
                              <input 
                                type="radio" 
                                name="hsrTicketTypeOutbound" 
                                value="non-reserved" 
                                checked={(formData.hsrTicketTypeOutbound || formData.hsrTicketType || 'standard') === 'non-reserved'} 
                                onChange={handleInputChange}
                                className="text-indigo-600 focus:ring-indigo-500 h-4 w-4" 
                              />
                              自由座
                            </label>
                          </div>
                        </div>
                        
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1">回程票種選擇</label>
                          {formData.isHsrSameTicketType !== false ? (
                            <div className="text-xs text-gray-500 py-1 pl-1 font-semibold">採用去程相同票種核銷</div>
                          ) : (
                            <div className="flex gap-4 p-1 animate-fadeIn">
                              <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700">
                                <input 
                                  type="radio" 
                                  name="hsrTicketTypeReturn" 
                                  value="standard" 
                                  checked={(formData.hsrTicketTypeReturn || 'standard') === 'standard'} 
                                  onChange={handleInputChange} 
                                  className="text-indigo-600 focus:ring-indigo-500 h-4 w-4" 
                                />
                                對號全票
                              </label>
                              <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700">
                                <input 
                                  type="radio" 
                                  name="hsrTicketTypeReturn" 
                                  value="non-reserved" 
                                  checked={(formData.hsrTicketTypeReturn || 'standard') === 'non-reserved'} 
                                  onChange={handleInputChange} 
                                  className="text-indigo-600 focus:ring-indigo-500 h-4 w-4" 
                                />
                                自由座
                              </label>
                            </div>
                          )}
                        </div>
                      </div>

                      <div className="flex flex-wrap gap-4 pt-2 border-t border-indigo-100/60 items-center justify-between text-xs text-indigo-900">
                        <label className="flex items-center gap-1.5 cursor-pointer font-bold select-none">
                          <input 
                            type="checkbox" 
                            name="isHsrSameStation" 
                            checked={formData.isHsrSameStation !== false} 
                            onChange={handleInputChange}
                            className="rounded border-gray-300 focus:ring-indigo-500 h-4 w-4 text-indigo-600" 
                          />
                          去回程起訖高鐵站完全相同
                        </label>

                        <label className="flex items-center gap-1.5 cursor-pointer font-bold select-none">
                          <input 
                            type="checkbox" 
                            name="isHsrSameTicketType" 
                            checked={formData.isHsrSameTicketType !== false} 
                            onChange={handleInputChange}
                            className="rounded border-gray-300 focus:ring-indigo-500 h-4 w-4 text-indigo-600" 
                          />
                          去回程申報票種相同
                        </label>
                      </div>
                    </div>
                  )}

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 border-t border-slate-100 pt-4">
                    <div>
                      <label className="block font-bold text-slate-700 mb-2">{formData.isMultiDay ? '首末兩日公出假別時段' : '出差請假時段'}</label>
                      <div className="flex gap-4">
                        <label className="flex items-center gap-1.5 cursor-pointer font-semibold text-slate-700"><input type="radio" name="duration" value="morning" checked={formData.duration === 'morning'} onChange={handleInputChange} className="text-indigo-600 focus:ring-indigo-500 h-4 w-4" />上午 (申領半日膳雜費)</label>
                        <label className="flex items-center gap-1.5 cursor-pointer font-semibold text-slate-700"><input type="radio" name="duration" value="afternoon" checked={formData.duration === 'afternoon'} onChange={handleInputChange} className="text-indigo-600 focus:ring-indigo-500 h-4 w-4" />下午 (申領半日膳雜費)</label>
                        <label className="flex items-center gap-1.5 cursor-pointer font-semibold text-slate-700"><input type="radio" name="duration" value="full" checked={formData.duration === 'full'} onChange={handleInputChange} className="text-indigo-600 focus:ring-indigo-500 h-4 w-4" />整天 (申領整日膳雜費)</label>
                      </div>
                    </div>

                    <div>
                      <label className="block font-bold text-slate-700 mb-1">{formData.isMultiDay ? '累計多日住宿費實報實銷總額 (若無留空)' : '住宿費申領金額 (若無留空)'}</label>
                      <input 
                        type="number" 
                        name="accommodationFee" 
                        min="0" 
                        placeholder="例如：1600" 
                        value={formData.accommodationFee} 
                        onChange={handleInputChange} 
                        className="w-full p-2.5 bg-white border border-gray-300 rounded-lg outline-none font-bold focus:ring-2 focus:ring-indigo-500" 
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-bold text-slate-700 mb-2">膳雜費申領狀態</label>
                      <label className="flex items-center gap-2 cursor-pointer mt-1 select-none">
                        <input 
                          type="checkbox" 
                          name="claimMealFee" 
                          checked={formData.claimMealFee} 
                          onChange={handleInputChange} 
                          className="w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500 cursor-pointer" 
                        />
                        <span className={formData.claimMealFee ? 'font-bold text-indigo-700' : 'text-gray-500 font-semibold'}>
                          {formData.claimMealFee ? '申領膳雜費補助 (依國內出差法規發放)' : '免申領膳雜費 (出差學校已供膳)'}
                        </span>
                      </label>
                    </div>

                    <div>
                      <label className="block font-bold text-slate-700 mb-2">出差範圍</label>
                      <label className="flex items-center gap-2 cursor-pointer mt-1 select-none">
                        <input type="checkbox" name="isOutCounty" checked={formData.isOutCounty} onChange={handleInputChange} className="w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500 cursor-pointer" />
                        <span className={formData.isOutCounty ? 'font-bold text-indigo-700 animate-pulse' : 'font-semibold text-slate-600'}>
                          {formData.isOutCounty ? '雲林縣外 (跨縣市，雜費標準 400元/天)' : '雲林縣內出差 (縣內，雜費標準 200元/天)'}
                        </span>
                      </label>
                    </div>
                  </div>

                  {/* 費用看板 */}
                  <div className="bg-indigo-50/50 border border-indigo-100 rounded-xl p-5 mt-4">
                    <h4 className="text-sm font-bold text-indigo-900 mb-3 flex items-center gap-2"><Calculator className="w-5 h-5 text-indigo-700" /> 會計複檢 · 單人費用預估</h4>
                    <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 text-xs font-bold">
                      <div className="p-3 bg-white rounded-lg border border-indigo-100">
                        <p className="text-indigo-700 mb-1">交通費 (車/高鐵來回)</p>
                        <p className="text-lg text-gray-800">NT$ {transportFee}</p>
                      </div>
                      <div className="p-3 bg-white rounded-lg border border-indigo-100">
                        <p className="text-indigo-700 mb-1">住宿費 (實報實銷)</p>
                        <p className="text-lg text-gray-800">NT$ {currentAccomFee}</p>
                      </div>
                      <div className="p-3 bg-white rounded-lg border border-indigo-100">
                        <p className="text-indigo-700 mb-1">膳雜費 (天數x標準額)</p>
                        <p className="text-lg text-gray-800">NT$ {mealFee}</p>
                      </div>
                      <div className="p-3 bg-indigo-600 rounded-lg text-white shadow-sm flex flex-col justify-center">
                        <p className="text-indigo-100 mb-0.5">預計請領金額 (單人)</p>
                        <p className="text-xl font-extrabold">NT$ {totalFee}</p>
                      </div>
                    </div>
                  </div>

                  <button type="submit" disabled={isProcessing} className="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg shadow-sm focus:outline-none focus:ring-4 focus:ring-emerald-500/50 transition-all mt-4 text-base tracking-wider">
                    {isProcessing ? '正在建立雲端數據檔...' : `建立儲存 ${formData.memberNames.length} 筆公務差旅請領單`}
                  </button>
                </form>
              </div>
            </div>
          )}

          {/* ================================================== */}
          {/* F. 差旅請領明細清冊 */}
          {/* ================================================== */}
          {activeTab === 'history' && (
            <div className="space-y-6">
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden animate-fadeIn">
                <div className="p-5 bg-slate-50 border-b flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                  <div>
                    <h3 className="font-bold text-slate-800">公務差旅請領明細與核銷狀態管控清冊</h3>
                    <p className="text-xs text-slate-500">此清冊支持會計「月報核對」，選定待報明細後，可一鍵產生 Word 請示單</p>
                  </div>
                  <div className="flex flex-wrap gap-2 w-full sm:w-auto">
                    {selectedRecordIds.length > 0 && (
                      <>
                        <button onClick={() => handleBatchStatusUpdate('verified')} className="flex items-center justify-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-xs font-bold transition shadow-sm focus:outline-none">
                          <CheckCircle className="w-4 h-4" /> 標記為已核銷
                        </button>
                        <button onClick={() => handleBatchStatusUpdate('pending')} className="flex items-center justify-center gap-1.5 bg-amber-600 hover:bg-amber-700 text-white px-3 py-2 rounded-lg text-xs font-bold transition shadow-sm focus:outline-none">
                          <Clock className="w-4 h-4" /> 恢復為待核銷
                        </button>
                      </>
                    )}
                    <button onClick={handleExportTravelWord} className="bg-indigo-600 hover:bg-indigo-700 text-white text-xs px-4 py-2.5 rounded-lg font-bold flex items-center space-x-1.5 shadow-sm transition">
                      <FileDown className="h-4 w-4" />
                      <span>產生勾選 Word 請示單</span>
                    </button>
                  </div>
                </div>

                {/* 篩選與月核 */}
                <div className="p-4 bg-slate-50 border-b border-slate-100 grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                  <select value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)} className="p-2.5 bg-white border border-slate-300 rounded-lg outline-none font-bold text-slate-700 cursor-pointer">
                    <option value="all">所有月份明細</option>
                    {availableMonths.map(m => (
                      <option key={m} value={m}>{m.replace('-', '年 ')}月</option>
                    ))}
                  </select>
                  <select value={filterName} onChange={(e) => setFilterName(e.target.value)} className="p-2.5 bg-white border border-slate-300 rounded-lg outline-none font-bold text-slate-700 cursor-pointer">
                    <option value="all">全體數辦同仁</option>
                    {employees.map(m => <option key={m.name} value={m.name}>{m.name}</option>)}
                  </select>
                  <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="p-2.5 bg-white border border-slate-300 rounded-lg outline-none font-bold text-slate-700 cursor-pointer">
                    <option value="pending">待核銷報帳明細</option>
                    <option value="verified">已核銷存檔明細</option>
                    <option value="all">顯示歷史全部明細</option>
                  </select>
                </div>

                <div className="overflow-x-auto">
                  <table className="min-w-full divide-y divide-slate-200 text-left text-sm font-medium text-slate-700">
                    <thead className="bg-slate-50 text-xs font-bold text-slate-500 uppercase tracking-wider">
                      <tr>
                        <th className="px-4 py-3 w-10">
                          <input 
                            type="checkbox" 
                            onChange={(e) => {
                              if (e.target.checked) {
                                const newSelections = new Set([...selectedRecordIds, ...filteredRecords.map(r => r.id)]);
                                setSelectedRecordIds(Array.from(newSelections));
                              } else {
                                const filteredIds = filteredRecords.map(r => r.id);
                                setSelectedRecordIds(selectedRecordIds.filter(id => !filteredIds.includes(id)));
                              }
                            }} 
                            checked={filteredRecords.length > 0 && filteredRecords.every(r => selectedRecordIds.includes(r.id))} 
                            className="rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer h-4 w-4" 
                          />
                        </th>
                        <th className="px-4 py-3">出差同仁</th>
                        <th className="px-4 py-3">日期及天數</th>
                        <th className="px-4 py-3">受訪學校 / 公務事由</th>
                        <th className="px-4 py-3">里程公里 / 工具</th>
                        <th className="px-4 py-3">申領小計金額</th>
                        <th className="px-4 py-3 text-right">管理</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 bg-white">
                      {currentRecords.map((rec) => {
                        const rStatus = rec.status || 'pending';
                        return (
                          <tr key={rec.id} className="hover:bg-slate-50/70 transition">
                            <td className="px-4 py-4">
                              <input type="checkbox" checked={selectedRecordIds.includes(rec.id)} onChange={() => setSelectedRecordIds(prev => prev.includes(rec.id) ? prev.filter(id => id !== rec.id) : [...prev, rec.id])} className="rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer h-4 w-4" />
                            </td>
                            <td className="px-4 py-4">
                              <span className="font-bold text-slate-900">{rec.name}</span>
                              <div className="text-[10px] text-slate-400 font-semibold">{rec.group} · {rec.title}</div>
                            </td>
                            <td className="px-4 py-4">
                              {rec.isMultiDay && rec.endDate ? (
                                <span className="text-xs font-bold text-indigo-800">{rec.date.substring(5)} ~ {rec.endDate.substring(5)} ({rec.days}天)</span>
                              ) : (
                                <span className="font-semibold text-slate-800">{rec.date}</span>
                              )}
                              <div className="mt-1">
                                {rStatus === 'verified' ? (
                                  <span className="inline-flex px-1.5 py-0.2 rounded-full text-[9px] bg-green-50 text-green-700 border border-green-200 font-bold">已核銷</span>
                                ) : (
                                  <span className="inline-flex px-1.5 py-0.2 rounded-full text-[9px] bg-amber-50 text-amber-800 border border-amber-200 font-bold">待核銷</span>
                                )}
                              </div>
                            </td>
                            <td className="px-4 py-4">
                              <span className="font-bold text-slate-900">{rec.destination}</span>
                              <div className="text-xs text-slate-400 mt-0.5">事由: {rec.reason || '辦事'}</div>
                            </td>
                            <td className="px-4 py-4 font-bold text-slate-700">
                              {rec.transportMode === 'hsr' ? (
                                <span className="text-xs font-bold text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded">高鐵 (雲林-{rec.hsrStation})</span>
                              ) : (
                                <span>{rec.totalDistance} km <span className="text-xs font-normal text-slate-400">({rec.transportMode === 'car' ? '汽' : '機'})</span></span>
                              )}
                            </td>
                            <td className="px-4 py-4 font-bold text-emerald-700">
                              NT$ {rec.totalFee}
                              <div className="text-[10px] font-medium text-slate-400 mt-0.5">膳:${rec.mealFee} / 住:${rec.accommodationFee || 0} / 交:${rec.transportFee}</div>
                            </td>
                            <td className="px-4 py-4 text-right">
                              <div className="flex justify-end gap-2">
                                {rec.transportMode !== 'hsr' && (
                                  <a 
                                    href={`https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(BASE_ADDRESS)}&destination=${encodeURIComponent(rec.destination)}&travelmode=driving${rec.transportMode === 'scooter' ? '&avoid=highways,tolls' : ''}`} 
                                    target="_blank" 
                                    rel="noreferrer" 
                                    className="px-2.5 py-1 text-green-700 border border-green-200 bg-green-50 rounded-lg hover:bg-green-100 text-xs font-bold flex items-center"
                                  >
                                    路徑
                                  </a>
                                )}
                                <button onClick={() => handleDeleteTravelClick(rec.id)} className="p-1.5 text-rose-500 hover:bg-rose-50 rounded-lg border border-transparent transition">
                                  <Trash2 className="h-4 w-4" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                {/* 分頁 */}
                {filteredRecords.length > 0 && (
                  <div className="p-4 border-t border-slate-100 flex justify-between items-center text-xs text-slate-500 bg-slate-50 font-medium">
                    <span>顯示第 {((currentTravelPage - 1) * travelRecordsPerPage) + 1} - {Math.min(currentTravelPage * travelRecordsPerPage, filteredRecords.length)} 筆，共 {filteredRecords.length} 筆</span>
                    <div className="flex items-center space-x-2">
                      <button disabled={currentTravelPage === 1} onClick={() => setCurrentTravelPage(p => Math.max(1, p - 1))} className="p-1.5 border rounded bg-white hover:bg-slate-50 shadow-sm transition"><ChevronLeft className="h-4 w-4" /></button>
                      <span className="font-bold">{currentTravelPage} / {totalTravelPages}</span>
                      <button disabled={currentTravelPage === totalTravelPages} onClick={() => setCurrentTravelPage(p => Math.min(totalTravelPages, p + 1))} className="p-1.5 border rounded bg-white hover:bg-slate-50 shadow-sm transition"><ChevronRight className="h-4 w-4" /></button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ================================================== */}
          {/* G. 報表下載中心 */}
          {/* ================================================== */}
          {activeTab === 'export' && (
            <div className="max-w-2xl mx-auto py-4 animate-fadeIn">
              <div className="text-center mb-8">
                <FileSpreadsheet className="w-14 h-14 text-emerald-600 mx-auto mb-3" />
                <h2 className="text-2xl font-bold text-gray-800">匯出會計專用國內出差旅費報告表</h2>
                <p className="text-gray-500 text-sm mt-1">
                  格式經會計室與審計法規精密比對校對。支援全體彙整表、中文大寫金額渲染與一鍵導出。
                </p>
              </div>

              {/* 報銷條件篩選 */}
              <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm space-y-4">
                <h3 className="text-base font-bold text-indigo-950 border-b pb-2 flex items-center gap-2">
                  <Filter className="w-5 h-5 text-indigo-700" /> 1. 設定下載之篩選月份
                </h3>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">報核月份 (按月編報)</label>
                    <select value={exportMonth} onChange={(e) => setExportMonth(e.target.value)} className="w-full p-2.5 bg-gray-50 border border-gray-300 rounded-lg outline-none text-xs font-bold focus:ring-2 focus:ring-indigo-500">
                      <option value="all">全部歷史月份</option>
                      {availableMonths.map(m => (
                        <option key={m} value={m}>{m.replace('-', '年 ')}月</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">報銷審核狀態</label>
                    <select value={exportStatus} onChange={(e) => setExportStatus(e.target.value)} className="w-full p-2.5 bg-gray-50 border border-gray-300 rounded-lg outline-none text-xs font-bold focus:ring-2 focus:ring-indigo-500">
                      <option value="pending">僅包含「待核銷」請領明細</option>
                      <option value="verified">僅包含「已核銷」封存明細</option>
                      <option value="all">不限狀態 (全部導出對帳)</option>
                    </select>
                  </div>
                </div>
              </div>

              <div className="bg-gray-50 p-6 rounded-xl border border-gray-200 space-y-4 mt-6 shadow-sm">
                <h3 className="text-base font-bold text-slate-800">2. 產生報告表並下載</h3>
                {(() => {
                  const count = records.filter(r => {
                    const rStatus = r.status || 'pending';
                    const matchMonth = exportMonth === 'all' || (r.date && r.date.startsWith(exportMonth));
                    const matchStatus = exportStatus === 'all' || rStatus === exportStatus;
                    return matchMonth && matchStatus;
                  }).length;
                  return (
                    <div className="text-xs bg-indigo-50 border border-indigo-100 text-indigo-900 p-3 rounded-lg flex justify-between items-center font-bold">
                      <span>此區間篩選出之申報明細共有 {count} 筆</span>
                      <span className="bg-indigo-200 px-2.5 py-0.5 rounded text-[10px]">{exportMonth === 'all' ? '不限月' : exportMonth} | {exportStatus === 'pending' ? '待核銷' : exportStatus === 'verified' ? '已核銷' : '全部'}</span>
                    </div>
                  );
                })()}

                <button onClick={() => exportExcel('all')} className="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg shadow font-bold flex items-center justify-center space-x-2 transition text-base tracking-wider">
                  <Download className="w-5 h-5" />
                  <span>下載全體出差旅費報告表 (彙整總表 + 各人明細分頁)</span>
                </button>

                <div className="border-t border-gray-200 pt-4">
                  <h4 className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wider">單獨下載同仁個人出差報告：</h4>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                    {employees.map(member => {
                      const mCount = records.filter(r => {
                        const rStatus = r.status || 'pending';
                        return r.name === member.name && (exportMonth === 'all' || r.date.startsWith(exportMonth)) && (exportStatus === 'all' || rStatus === exportStatus);
                      }).length;
                      return (
                        <button key={member.name} onClick={() => exportExcel(member.name)} disabled={mCount === 0} className={`px-2 py-2.5 text-xs rounded-lg border transition flex flex-col items-center justify-center focus:outline-none ${mCount > 0 ? 'bg-white text-slate-800 hover:border-emerald-600 hover:text-emerald-800 font-bold focus:ring-2 focus:ring-emerald-500' : 'bg-slate-100 text-slate-400 cursor-not-allowed border-transparent'}`}>
                          <span>{member.name}</span>
                          <span className="text-[9px] mt-0.5 bg-slate-100 px-2 py-0.2 rounded-full text-slate-500">{mCount} 筆</span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ================================================== */}
          {/* H. 系統 API 與設定 */}
          {/* ================================================== */}
          {activeTab === 'settings' && (
            <div className="max-w-xl mx-auto bg-white p-8 rounded-xl border border-slate-200 shadow-sm space-y-4 animate-fadeIn">
              <div className="text-center space-y-2">
                <Settings className="h-12 w-12 text-slate-400 mx-auto" />
                <h3 className="font-bold text-slate-900 text-lg">系統 API 與維護設定</h3>
                <p className="text-xs text-slate-500 leading-relaxed">設定 <strong>Google Cloud Routes API Key</strong> 以進行地理導航測距。若留空，系統將自動啟動 OpenStreetMap/OSRM 計算。密碼與刪除鎖定密碼為數辦專用。</p>
              </div>
              <div className="space-y-3">
                <input 
                  type="password" 
                  value={gmapsKey} 
                  onChange={(e) => setGmapsKey(e.target.value)} 
                  placeholder="請輸入 AIzaSy 開頭的 API 密鑰" 
                  className="w-full p-2.5 bg-slate-50 border rounded-lg text-sm text-center outline-none tracking-widest font-mono focus:ring-2 focus:ring-indigo-500 transition-shadow" 
                />
                <button 
                  onClick={() => {
                    localStorage.setItem('gmaps_api_key', gmapsKey);
                    showDialog('alert', '設定成功', 'Google API 金鑰已儲存並成功套用！');
                  }} 
                  className="w-full py-2.5 bg-slate-900 hover:bg-black text-white text-xs font-bold rounded-lg transition"
                >
                  儲存並套用金鑰
                </button>
              </div>
            </div>
          )}

        </main>
      </div>

      {/* ==================== Modal 1. 休息日加班團體同意書 A4 預覽 ==================== */}
      {showAgreementModal && selectedOtRequest && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl w-full max-w-3xl shadow-2xl overflow-hidden border border-slate-200 my-8 modal-container-fixed print:my-0 print:border-none">
            <div className="px-6 py-4 bg-slate-900 text-white flex justify-between items-center print:hidden">
              <span className="font-bold flex items-center space-x-1.5 text-sm sm:text-base">
                <Printer className="h-5 w-5 text-teal-400" />
                <span>休息日工作協商同意書 (團體格式列印預覽)</span>
              </span>
              <button type="button" onClick={() => { setShowAgreementModal(false); setSelectedOtRequest(null); }} className="text-slate-400 hover:text-white transition text-lg">✕</button>
            </div>

            <div className="p-8 sm:p-10 bg-white text-black font-serif leading-relaxed overflow-y-auto max-h-[70vh] print:p-0 print:max-h-none print:overflow-visible">
              <div className="max-w-2xl mx-auto space-y-5">
                <h2 className="text-2xl sm:text-3xl font-bold text-center tracking-widest border-b-2 border-black pb-2">協商同意書</h2>
                <p className="text-justify text-base sm:text-lg indent-8 leading-loose pt-2">
                  因 <span className="font-bold border-b border-black px-1">{formatMinguoDateText(selectedOtRequest.workDate)}</span> ( {getWeekdayText(selectedOtRequest.workDate)} ) 上午 {selectedOtRequest.startTime} 時至下午 {selectedOtRequest.endTime} 時原訂為勞工休息日，本校 <u> （名單如下表） </u> 同意於前開休息日協助 <span className="font-bold border-b border-black px-1">{selectedOtRequest.activityName}</span>，並經勞雇雙方協商同意於 {formatMinguoDateText(selectedOtRequest.workDate)} 次日起一年內完成補休 (共計 <span className="font-bold border-b border-black px-1">{selectedOtRequest.hours}</span> 小時)，特此證明。
                </p>
                <div className="bg-slate-50 p-4 border border-black text-xs sm:text-sm font-sans space-y-1 rounded">
                  <p className="font-bold">法條依據：</p>
                  <p className="text-slate-700 text-justify">
                    <span className="font-bold">勞動基準法第 32-1 條：</span>
                    雇主使勞工休息日工作後，依勞工意願選擇補休並經雇主同意者，應依勞工工作之時數計算補休時數。前項之補休，其補休期限由勞雇雙方協商。
                  </p>
                </div>
                <div className="space-y-2 pt-2">
                  <p className="font-bold text-sm">協商同意勞工名單：</p>
                  <div className="border border-black overflow-hidden font-sans text-sm sm:text-base">
                    <table className="w-full border-collapse">
                      <tbody>
                        {chunkParticipants(selectedOtRequest.participants).map(([p1, p2], idx) => (
                          <tr key={idx} className="border-b border-black last:border-b-0 h-[40px]">
                            <td className="border-r border-black p-2.5 w-1/2">
                              <span>姓名：<strong className="text-black">{p1}</strong></span>
                            </td>
                            <td className="p-2.5 w-1/2">
                              {p2 ? <span>姓名：<strong className="text-black">{p2}</strong></span> : <span className="text-slate-300">姓名：</span>}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-4 pt-10 text-center text-sm font-sans">
                  <div><p className="border-t border-black pt-2 mx-2">執行秘書</p></div>
                  <div><p className="border-t border-black pt-2 mx-2">中心主任</p></div>
                  <div><p className="border-t border-black pt-2 mx-2">校長</p></div>
                </div>
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex flex-wrap justify-end gap-2 print:hidden">
              <button type="button" onClick={() => { setShowAgreementModal(false); setSelectedOtRequest(null); }} className="px-4 py-2 text-sm text-slate-500 hover:text-slate-700 font-semibold">關閉預覽</button>
              <button type="button" onClick={() => handleExportWord(selectedOtRequest)} className="px-4 py-2 text-sm text-teal-800 bg-teal-50 hover:bg-teal-100 border border-teal-200 rounded-xl font-bold flex items-center space-x-1.5 transition duration-150">
                <Download className="h-4 w-4" />
                <span>匯出 Word 同意書 (.doc)</span>
              </button>
              <button type="button" onClick={() => handlePrintPDF(selectedOtRequest)} className="px-5 py-2 text-sm text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl shadow font-bold flex items-center space-x-1.5 transition duration-150">
                <Printer className="h-4 w-4" />
                <span>列印 / 另存為 PDF</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== Modal 2. 安全防刪除安全驗證 ==================== */}
      {deleteTarget && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl overflow-hidden border border-slate-200 p-6 space-y-4">
            <div className="flex items-center space-x-2 text-rose-600">
              <AlertCircle className="h-6 w-6 animate-pulse" />
              <h3 className="text-lg font-bold">永久安全刪除確認？</h3>
            </div>
            <p className="text-xs text-slate-500 leading-relaxed">
              您正在刪除：<strong className="text-slate-800 font-semibold">{deleteTarget.name}</strong> 的雲端數據。此操作無法撤銷。
            </p>
            <form onSubmit={handleConfirmDelete} className="space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">請輸入防誤刪安全保護密碼 (056341014)</label>
                <input type="password" placeholder="請輸入安全保護密碼" value={deletePassword} onChange={(e) => setDeletePassword(e.target.value)} className="w-full rounded-xl border-slate-300 text-center text-sm font-bold tracking-widest focus:border-rose-500 focus:ring-rose-500 py-2.5 outline-none border" autoFocus required />
              </div>
              <div className="flex justify-end space-x-3 pt-2 text-xs font-bold">
                <button type="button" onClick={() => { setDeleteTarget(null); setDeletePassword(''); }} className="px-4 py-2 text-slate-500 hover:text-slate-700">取消</button>
                <button type="submit" className="px-6 py-2 text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-sm">確認刪除</button>
              </div>
            </form>
          </div>
        </div>
      )}

      <footer className="bg-slate-950 text-slate-500 border-t border-slate-900 text-xs py-4 text-center print:hidden font-medium">
        <p>© 115學年度 雲林縣土庫鎮越港國民小學 數位學習推動辦公室 版權所有</p>
      </footer>
    </div>
  );
}
