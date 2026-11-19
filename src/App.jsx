import React, { useState, useMemo, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  Calendar as CalendarIcon,
  Users,
  Bell,
  Plus,
  Download,
  RefreshCw,
  Filter,
  Search,
  ChevronLeft,
  ChevronRight,
  MessageCircle,
  CheckCircle2,
  Clock,
  Phone,
  MapPin,
  ExternalLink
} from 'lucide-react';
import { PANCHANG_DATA, JAIN_MONTHS_META } from './panchangData';
import { INITIAL_ENTRIES } from './initialEntries';

// Your Live Published Google Sheet CSV URL
// Instant Live Google Sheet CSV URL (No 5-minute Publish-to-Web cache delay)
const GOOGLE_SHEET_CSV_URL =
  'https://docs.google.com/spreadsheets/d/15kuAgpcs4NNF9fE2WnUo3Odjx43BKu_5N2pSxbNpsmw/gviz/tq?tqx=out:csv&gid=755030390&headers=4';
  const WEEKDAYS = [
  { key: 0, short: 'SUN', guj: 'રવિ' },
  { key: 1, short: 'MON', guj: 'સોમ' },
  { key: 2, short: 'TUE', guj: 'મંગળ' },
  { key: 3, short: 'WED', guj: 'બુધ' },
  { key: 4, short: 'THU', guj: 'ગુરુ' },
  { key: 5, short: 'FRI', guj: 'શુક્ર' },
  { key: 6, short: 'SAT', guj: 'શનિ' },
];

const GUJARATI_TITHI_NAMES = {
  1: 'એકમ',
  2: 'બીજ',
  3: 'ત્રીજ',
  4: 'ચોથ',
  5: 'પાંચમ',
  6: 'છઠ્ઠ',
  7: 'સાતમ',
  8: 'આઠમ',
  9: 'નોમ',
  10: 'દશમ',
  11: 'અગિયારસ',
  12: 'બારસ',
  13: 'તેરસ',
  14: 'ચૌદશ',
  15: 'પૂનમ',
  30: 'અમાસ'
};

function addDays(isoDate, days) {
  if (!isoDate) return '';
  const [y, m, d] = isoDate.split('-').map(Number);
  if (!y || !m || !d) return '';
  const dateObj = new Date(Date.UTC(y, m - 1, d + days));
  return dateObj.toISOString().split('T')[0];
}

function formatDisplayDate(isoDate) {
  if (!isoDate || !isoDate.includes('-')) return isoDate || '-';
  const [y, m, d] = isoDate.split('-');
  return `${d}/${m}/${y}`;
}

// Accepts YYYY-MM-DD, DD-MM-YYYY, or DD/MM/YYYY and returns YYYY-MM-DD
function normalizeToIsoDate(raw) {
  if (!raw) return '';
  const cleaned = String(raw).replace(/['"]/g, '').trim();
  if (!cleaned) return '';

  const isoMatch = cleaned.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2].padStart(2, '0')}-${isoMatch[3].padStart(2, '0')}`;
  }

  const slashMatch = cleaned.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (slashMatch) {
    const first = Number(slashMatch[1]);
    const second = Number(slashMatch[2]);
    const year = slashMatch[3];
    const month = first > 12 ? second : first;
    const day = first > 12 ? first : second;
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  return '';
}

export default function App() {
  const [activeTab, setActiveTab] = useState('calendar'); // 'calendar' | 'directory' | 'alerts' | 'add'
  const [entries, setEntries] = useState(INITIAL_ENTRIES);
  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSyncedAt, setLastSyncedAt] = useState('');
  const [syncError, setSyncError] = useState('');

  // Track which alerts have been marked sent locally
  const [sentIds, setSentIds] = useState(() => {
    const saved = localStorage.getItem('jain_sent_alerts_v1');
    return saved ? JSON.parse(saved) : {};
  });

  useEffect(() => {
    localStorage.setItem('jain_sent_alerts_v1', JSON.stringify(sentIds));
  }, [sentIds]);

  // Fetch & Parse Live Google Sheet CSV
  const fetchLiveGoogleSheet = useCallback(async () => {
    setIsSyncing(true);
    setSyncError('');
    try {
      const response = await fetch(`${GOOGLE_SHEET_CSV_URL}&t=${Date.now()}`);
      if (!response.ok) throw new Error('Failed to fetch Google Sheet');
      const csvText = await response.text();

      // Use SheetJS to parse CSV cleanly (handles Gujarati text, quotes, commas inside addresses)
      const wb = XLSX.read(csvText, { type: 'string' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rowsArray = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      // Automatically find the Gujarati header row (usually Row 4, index 3)
      const headerRowIndex = rowsArray.findIndex((row) =>
        row.some((cell) => String(cell).includes('આરાધક પૂજ્યશ્રીનું નામ'))
      );

      if (headerRowIndex === -1) {
        throw new Error('Header row not found in Google Sheet');
      }

      const headers = rowsArray[headerRowIndex].map((h) => String(h).trim());
      const dataRows = rowsArray.slice(headerRowIndex + 1);

      const getCol = (row, colNameFragment, fallbackIndex = -1) => {
        const idx = headers.findIndex((h) => h.includes(colNameFragment));
        const val = idx !== -1 ? String(row[idx] || '').trim() : '';
        if (val) return val;
        return fallbackIndex !== -1 ? String(row[fallbackIndex] || '').trim() : '';
      };

      const parsedEntries = dataRows
        .filter((row) => getCol(row, 'આરાધક પૂજ્યશ્રીનું નામ', 1) !== '')
        .map((row, idx) => {
          const rawStart =
            getCol(row, 'ફાળવેલ શરૂઆત તારીખ', 15) ||
            getCol(row, 'ફાળવેલ તારીખ', 15) ||
            String(row[15] || row[row.length - 1] || '').trim();
          const isoStart = normalizeToIsoDate(rawStart);

          return {
            id: idx + 1,
            serialNo: getCol(row, 'અનુક્રમ નંબર') || String(idx + 1),
            aaradhakName: getCol(row, 'આરાધક પૂજ્યશ્રીનું નામ'),
            guruBhagwantName: getCol(row, 'પૂ. ગુરુભગવંતશ્રીનું નામ'),
            dikshaParyay: getCol(row, 'દિક્ષા પર્યાય'),
            samuday: getCol(row, 'સમુદાય'),
            convenientTithi: getCol(row, 'આપણને અનુકૂળ તિથી'),
            dateCol: getCol(row, 'તારીખ'),
            yearCol: getCol(row, 'સાલ'),
            moreThanOneAshtam: getCol(row, 'એકથી વધારે અષ્ટમ'),
            requestedDateText: getCol(row, 'આપના અનુકૂળ દિવસો'),
            ourDaysText: getCol(row, 'અથવા અમારા દિવસો'),
            mobileNumber: getCol(row, 'મોબાઈલ નંબર'),
            address: getCol(row, 'સરનામું'),
            entryDate: getCol(row, 'પ્રતિસાદ નોંધણી તારીખ'),
            remarks: getCol(row, 'સ્થિતિ'),
            allottedStartDate: isoStart,
            allottedDaysCount: 3
          };
        });

      if (parsedEntries.length > 0) {
        setEntries(parsedEntries);
      }
      setLastSyncedAt(new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }));
    } catch (err) {
      console.error(err);
      setSyncError('Google Sheet સિંક કરવામાં ભૂલ આવી. ઇન્ટરનેટ તપાસો.');
    } finally {
      setIsSyncing(false);
    }
  }, []);

  // Sync automatically when page loads
  useEffect(() => {
    fetchLiveGoogleSheet();
  }, [fetchLiveGoogleSheet]);

  // Calendar State
  const [selectedJainMonth, setSelectedJainMonth] = useState('Posh');
  const [selectedDateDetail, setSelectedDateDetail] = useState('2027-01-27');

  // Filter State
  const [selectedGuruFilter, setSelectedGuruFilter] = useState('ALL');
  const [searchQuery, setSearchQuery] = useState('');

  // Unique Guru Bhagwant Names for One-Click Filter Pills
  const uniqueGurus = useMemo(() => {
    const set = new Set();
    entries.forEach((e) => {
      if (e.guruBhagwantName && e.guruBhagwantName.trim()) {
        set.add(e.guruBhagwantName.trim());
      }
    });
    return Array.from(set);
  }, [entries]);

  // Filtered Directory Entries
  const filteredEntries = useMemo(() => {
    return entries.filter((entry) => {
      const matchesGuru =
        selectedGuruFilter === 'ALL' || entry.guruBhagwantName === selectedGuruFilter;
      const q = searchQuery.toLowerCase();
      const matchesSearch =
        !q ||
        (entry.aaradhakName || '').toLowerCase().includes(q) ||
        (entry.guruBhagwantName || '').toLowerCase().includes(q) ||
        (entry.samuday || '').toLowerCase().includes(q) ||
        (entry.address || '').toLowerCase().includes(q) ||
        (entry.mobileNumber || '').toLowerCase().includes(q);
      return matchesGuru && matchesSearch;
    });
  }, [entries, selectedGuruFilter, searchQuery]);

  // Map every ISO Date -> Allotted Maharaj Saheb(s) for their 3-day slot
  const dateAllocationMap = useMemo(() => {
    const map = {};
    entries.forEach((entry) => {
      if (!entry.allottedStartDate) return;
      const count = Number(entry.allottedDaysCount) || 3;
      for (let i = 0; i < count; i++) {
        const d = addDays(entry.allottedStartDate, i);
        if (!d) continue;
        if (!map[d]) map[d] = [];
        map[d].push({
          ...entry,
          dayNumber: i + 1,
          totalDays: count
        });
      }
    });
    return map;
  }, [entries]);

  // Panchang lookup by ISO date
  const panchangByDate = useMemo(() => {
    const map = {};
    PANCHANG_DATA.forEach((item) => {
      map[item.date] = item;
    });
    return map;
  }, []);

  // Build Calendar Grid for Selected Jain Month
  const currentMonthDays = useMemo(() => {
    return PANCHANG_DATA.filter((d) => d.jainMonth === selectedJainMonth);
  }, [selectedJainMonth]);

  const currentMonthMeta = useMemo(() => {
    return JAIN_MONTHS_META.find((m) => m.id === selectedJainMonth) || JAIN_MONTHS_META[0];
  }, [selectedJainMonth]);

  const calendarMatrix = useMemo(() => {
    if (!currentMonthDays.length) return [];
    const firstDate = new Date(currentMonthDays[0].date + 'T00:00:00');
    const firstDow = firstDate.getDay();

    const cells = [];
    for (let i = 0; i < firstDow; i++) cells.push(null);
    currentMonthDays.forEach((dayObj) => cells.push(dayObj));
    while (cells.length % 7 !== 0) cells.push(null);

    const numWeeks = cells.length / 7;
    const rows = Array.from({ length: 7 }, () => Array(numWeeks).fill(null));
    for (let w = 0; w < numWeeks; w++) {
      for (let dow = 0; dow < 7; dow++) {
        rows[dow][w] = cells[w * 7 + dow];
      }
    }
    return rows;
  }, [currentMonthDays]);

  // 3-Day Advance WhatsApp Alert Queue
  const alertList = useMemo(() => {
    return entries
      .filter((e) => e.allottedStartDate)
      .map((e) => {
        const reminderDate = addDays(e.allottedStartDate, -3);
        const endDate = addDays(e.allottedStartDate, (Number(e.allottedDaysCount) || 3) - 1);
        const pStart = panchangByDate[e.allottedStartDate];
        const tithiLabel = pStart
          ? `${pStart.jainMonth} ${pStart.paksha === 'Sud' ? 'સુદ' : 'વદ'} ${GUJARATI_TITHI_NAMES[pStart.tithi] || pStart.tithi}`
          : '';
        return {
          ...e,
          reminderDate,
          endDate,
          tithiLabel
        };
      })
      .sort((a, b) => a.reminderDate.localeCompare(b.reminderDate));
  }, [entries, panchangByDate]);

  // Generate Pre-filled WhatsApp Link
  const buildWhatsAppUrl = (item) => {
    const rawPhone = (item.mobileNumber || '').split('/')[0].replace(/\D/g, '');
    const phone = rawPhone.length === 10 ? `91${rawPhone}` : rawPhone;
    const msg =
      `🙏 પ્રણામ / સાદર જય જિનેન્દ્ર,\n\n` +
      `આરાધક પૂજ્યશ્રી: *${item.aaradhakName}*\n` +
      `પૂ. ગુરુભગવંતશ્રી: *${item.guruBhagwantName}*\n\n` +
      `આપની અષ્ટમ તપ આરાધનાના નિર્ધારિત ૩ દિવસ તા. *${formatDisplayDate(item.allottedStartDate)}* થી *${formatDisplayDate(item.endDate)}* (${item.tithiLabel}) ના રોજ શરૂ થનાર છે (આજથી ૩ દિવસ બાદ).\n\n` +
      `કૃપા કરીને નોંધ લેશો. 🙏`;
    return `https://wa.me/${phone}?text=${encodeURIComponent(msg)}`;
  };

  // Temporary inline date preview (and quick save reminder)
  const handleUpdateAllottedDate = (id, newStartDate) => {
    setEntries((prev) =>
      prev.map((item) => (item.id === id ? { ...item, allottedStartDate: newStartDate } : item))
    );
  };

  const toggleReminderSent = (key) => {
    setSentIds((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  // Export Current Data to Gujarati Excel (.xlsx)
  const handleExportExcel = () => {
    const rows = entries.map((e, idx) => ({
      'અનુક્રમ નંબર': idx + 1,
      'આરાધક પૂજ્યશ્રીનું નામ': e.aaradhakName,
      'પૂ. ગુરુભગવંતશ્રીનું નામ': e.guruBhagwantName,
      'દિક્ષા પર્યાય': e.dikshaParyay,
      'સમુદાય / ગચ્છનું નામ': e.samuday,
      'આપણને અનુકૂળ તિથી': e.convenientTithi,
      'તારીખ': e.dateCol,
      'સાલ': e.yearCol,
      'એકથી વધારે અષ્ટમ તપ કરવાની ભાવના ખરી? (હા/ના)': e.moreThanOneAshtam,
      'આપના અનુકૂળ દિવસો': e.requestedDateText,
      'અથવા અમારા દિવસો': e.ourDaysText,
      'સંપર્ક સૂત્રનો મોબાઈલ નંબર': e.mobileNumber,
      'પત્ર વ્યવહારનું સરનામું': e.address,
      'પ્રતિસાદ નોંધણી તારીખ': e.entryDate,
      'સ્થિતિ / નોંધ': e.remarks,
      'ફાળવેલ શરૂઆત તારીખ': e.allottedStartDate
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'આરાધક આવેલ પ્રતિસાદ વિગત');
    XLSX.writeFile(wb, 'Aaradhak_Response_Live_Backup.xlsx');
  };

  const monthIndex = JAIN_MONTHS_META.findIndex((m) => m.id === selectedJainMonth);

  return (
    <div className="min-h-screen bg-stone-100 pb-16">
      {/* Top Header */}
      <header className="bg-maroon-800 text-white shadow-lg sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-4 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-amber-300 text-xl">卐</span>
              <h1 className="text-lg sm:text-xl font-bold tracking-tight">
                આરાધક પૂજ્યશ્રી વિગતો & જૈન પંચાંગ (VS 2083)
              </h1>
            </div>
            <div className="flex flex-wrap items-center gap-2 text-xs text-maroon-100 mt-0.5">
              <span>શ્વેતાંબર મૂર્તિપૂજક ૧-તિથિ પંચાંગ</span>
              <span>•</span>
              <span className="inline-flex items-center gap-1 text-emerald-300 font-medium">
                ● Live Google Sheet Connected {lastSyncedAt && `(Synced ${lastSyncedAt})`}
              </span>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={fetchLiveGoogleSheet}
              disabled={isSyncing}
              className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-2 rounded-lg shadow transition disabled:opacity-60"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
              <span>{isSyncing ? 'Syncing...' : 'Sync Google Sheet'}</span>
            </button>

            <button
              onClick={handleExportExcel}
              className="inline-flex items-center gap-1.5 bg-amber-500 hover:bg-amber-600 text-stone-950 text-xs font-semibold px-3 py-2 rounded-lg shadow transition"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Download Excel</span>
            </button>
          </div>
        </div>

        {syncError && (
          <div className="bg-rose-900 text-rose-100 text-xs px-4 py-1.5 text-center">
            {syncError}
          </div>
        )}

        {/* Responsive Navigation Tabs */}
        <nav className="bg-maroon-900/80 border-t border-maroon-700">
          <div className="max-w-7xl mx-auto px-4 flex overflow-x-auto gap-1 py-1.5">
            <button
              onClick={() => setActiveTab('calendar')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium whitespace-nowrap transition ${
                activeTab === 'calendar'
                  ? 'bg-white text-maroon-900 shadow'
                  : 'text-maroon-100 hover:bg-maroon-800'
              }`}
            >
              <CalendarIcon className="w-4 h-4" />
              <span>જૈન તિથિ કેલેન્ડર (Calendar)</span>
            </button>

            <button
              onClick={() => setActiveTab('directory')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium whitespace-nowrap transition ${
                activeTab === 'directory'
                  ? 'bg-white text-maroon-900 shadow'
                  : 'text-maroon-100 hover:bg-maroon-800'
              }`}
            >
              <Users className="w-4 h-4" />
              <span>આરાધક યાદી & ગુરુ ફિલ્ટર ({entries.length})</span>
            </button>

            <button
              onClick={() => setActiveTab('alerts')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium whitespace-nowrap transition ${
                activeTab === 'alerts'
                  ? 'bg-white text-maroon-900 shadow'
                  : 'text-maroon-100 hover:bg-maroon-800'
              }`}
            >
              <Bell className="w-4 h-4" />
              <span>૩-દિવસ પહેલા WhatsApp એલર્ટ ({alertList.length})</span>
            </button>

            <button
              onClick={() => setActiveTab('add')}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium whitespace-nowrap transition ${
                activeTab === 'add'
                  ? 'bg-amber-400 text-stone-950 font-semibold shadow'
                  : 'text-amber-200 hover:bg-maroon-800'
              }`}
            >
              <Plus className="w-4 h-4" />
              <span>નવી વિગત ઉમેરો (Google Sheet)</span>
            </button>
          </div>
        </nav>
      </header>

      <main className="max-w-7xl mx-auto px-4 pt-6">
        {/* ================= TAB 1: JAIN TITHI & ALLOCATION CALENDAR ================= */}
        {activeTab === 'calendar' && (
          <div className="space-y-6">
            {/* Month Switcher Bar */}
            <div className="bg-white rounded-2xl shadow-sm border border-stone-200 p-4 flex flex-col md:flex-row items-center justify-between gap-4">
              <div className="flex items-center gap-3">
                <button
                  disabled={monthIndex <= 0}
                  onClick={() => setSelectedJainMonth(JAIN_MONTHS_META[monthIndex - 1].id)}
                  className="p-2 rounded-full bg-maroon-50 text-maroon-800 hover:bg-maroon-100 disabled:opacity-30"
                >
                  <ChevronLeft className="w-5 h-5" />
                </button>
                <div className="text-center min-w-[210px]">
                  <h2 className="text-xl font-bold text-maroon-900">
                    {currentMonthMeta.guj} ({currentMonthMeta.eng})
                  </h2>
                  <p className="text-xs text-stone-500">{currentMonthMeta.range}</p>
                </div>
                <button
                  disabled={monthIndex >= JAIN_MONTHS_META.length - 1}
                  onClick={() => setSelectedJainMonth(JAIN_MONTHS_META[monthIndex + 1].id)}
                  className="p-2 rounded-full bg-maroon-50 text-maroon-800 hover:bg-maroon-100 disabled:opacity-30"
                >
                  <ChevronRight className="w-5 h-5" />
                </button>
              </div>

              {/* Quick Month Pills */}
              <div className="flex flex-wrap justify-center gap-1.5">
                {JAIN_MONTHS_META.map((m) => (
                  <button
                    key={m.id}
                    onClick={() => setSelectedJainMonth(m.id)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition ${
                      selectedJainMonth === m.id
                        ? 'bg-maroon-800 text-white shadow'
                        : 'bg-stone-100 text-stone-700 hover:bg-stone-200'
                    }`}
                  >
                    {m.guj.split(' ')[0]}
                  </button>
                ))}
              </div>
            </div>

            {/* Color Legend */}
            <div className="flex flex-wrap items-center justify-between gap-3 px-1 text-xs">
              <div className="flex flex-wrap items-center gap-3">
                <span className="inline-flex items-center gap-1.5 bg-emerald-100 text-emerald-900 px-3 py-1 rounded-full font-medium">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-600"></span> સુદ પક્ષ (Sud)
                </span>
                <span className="inline-flex items-center gap-1.5 bg-rose-100 text-rose-900 px-3 py-1 rounded-full font-medium">
                  <span className="w-2.5 h-2.5 rounded-full bg-rose-600"></span> વદ પક્ષ (Vad)
                </span>
                <span className="inline-flex items-center gap-1.5 bg-amber-200 text-amber-950 px-3 py-1 rounded-full font-medium">
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-600"></span> પર્વ / અષ્ટમી / ચતુર્દશી
                </span>
              </div>
              <span className="text-stone-500">
                કોઈપણ તારીખ પર ક્લિક કરીને ફાળવેલ મહારાજ સાહેબની વિગત જુઓ
              </span>
            </div>

            {/* Calendar Grid */}
            <div className="bg-white rounded-2xl shadow-sm border border-stone-200 overflow-x-auto">
              <table className="w-full border-collapse min-w-[680px]">
                <tbody>
                  {WEEKDAYS.map((dayObj, rowIdx) => (
                    <tr key={dayObj.key} className="border-b border-stone-200 last:border-none">
                      <td className="w-20 bg-stone-50 border-r border-stone-200 p-2 text-center">
                        <div className="text-xs font-bold text-stone-700">{dayObj.short}</div>
                        <div className="text-[11px] text-stone-500">{dayObj.guj}</div>
                      </td>

                      {calendarMatrix[rowIdx]?.map((cell, colIdx) => {
                        if (!cell) {
                          return (
                            <td
                              key={colIdx}
                              className="border-r border-stone-200 last:border-none bg-stone-50/60 p-2 text-center align-middle"
                            >
                              {rowIdx === 0 && colIdx === 0 && (
                                <div className="text-maroon-800 font-bold text-lg">卐</div>
                              )}
                            </td>
                          );
                        }

                        const allocs = dateAllocationMap[cell.date] || [];
                        const isSelected = selectedDateDetail === cell.date;
                        const [, mStr, dStr] = cell.date.split('-');
                        const gujTithi = GUJARATI_TITHI_NAMES[cell.tithi] || '';

                        let bgClass =
                          cell.paksha === 'Sud'
                            ? 'bg-emerald-50/90 text-emerald-900'
                            : 'bg-rose-50/90 text-rose-900';
                        if (cell.isImportant) {
                          bgClass = 'bg-amber-200/90 text-amber-950';
                        }

                        return (
                          <td
                            key={cell.date}
                            onClick={() => setSelectedDateDetail(cell.date)}
                            className={`border-r border-stone-200 last:border-none p-2 align-top cursor-pointer transition relative min-w-[130px] h-24 ${bgClass} ${
                              isSelected ? 'ring-2 ring-inset ring-maroon-800' : 'hover:brightness-95'
                            }`}
                          >
                            <div className="flex items-start justify-between">
                              <span className="text-xl font-extrabold leading-none">
                                {cell.tithi === 30 ? '૩૦' : cell.tithi}
                              </span>
                              <span className="text-[11px] font-semibold bg-white/75 text-stone-700 px-1.5 py-0.5 rounded">
                                {dStr}/{mStr}
                              </span>
                            </div>

                            <div className="text-[11px] font-semibold opacity-85 mt-0.5">
                              {cell.paksha === 'Sud' ? 'સુદ' : 'વદ'} {gujTithi}
                            </div>

                            {allocs.length > 0 && (
                              <div className="mt-1.5 space-y-1">
                                {allocs.map((a) => (
                                  <div
                                    key={a.id}
                                    className="bg-maroon-800 text-white text-[10px] font-semibold px-1.5 py-0.5 rounded shadow-sm truncate"
                                    title={`${a.aaradhakName} (Day ${a.dayNumber}/${a.totalDays})`}
                                  >
                                    • {a.aaradhakName} (D{a.dayNumber})
                                  </div>
                                ))}
                              </div>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Selected Day Details Panel */}
            {selectedDateDetail && (
              <div className="bg-white rounded-2xl shadow-sm border border-stone-200 p-5">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-3 border-b border-stone-100">
                  <div>
                    <h3 className="text-base font-bold text-maroon-900">
                      તારીખ વિગત: {formatDisplayDate(selectedDateDetail)}
                      {panchangByDate[selectedDateDetail] && (
                        <span className="ml-2 text-sm font-semibold text-amber-700">
                          ({panchangByDate[selectedDateDetail].jainMonth}{' '}
                          {panchangByDate[selectedDateDetail].paksha === 'Sud' ? 'સુદ' : 'વદ'}{' '}
                          {GUJARATI_TITHI_NAMES[panchangByDate[selectedDateDetail].tithi]} - VS 2083)
                        </span>
                      )}
                    </h3>
                  </div>
                </div>

                {(dateAllocationMap[selectedDateDetail] || []).length === 0 ? (
                  <p className="text-sm text-stone-500 py-4">
                    આ તારીખે કોઈ આરાધક પૂજ્યશ્રીના ૩ દિવસ ફાળવેલ નથી. (તમારી Google Sheet માં Column P માં{' '}
                    <code className="bg-stone-100 px-1.5 py-0.5 rounded font-mono text-xs">
                      {selectedDateDetail}
                    </code>{' '}
                    લખવાથી અહીં નામ આપમેળે આવી જશે.)
                  </p>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-4">
                    {dateAllocationMap[selectedDateDetail].map((item) => (
                      <div
                        key={item.id}
                        className="p-4 rounded-xl bg-maroon-50/70 border border-maroon-200 flex flex-col justify-between gap-2"
                      >
                        <div>
                          <div className="flex items-center justify-between">
                            <span className="text-xs font-bold bg-maroon-800 text-white px-2 py-0.5 rounded">
                              દિવસ {item.dayNumber} / {item.totalDays}
                            </span>
                            <span className="text-xs text-stone-600">{item.samuday}</span>
                          </div>
                          <h4 className="font-bold text-stone-900 mt-2">{item.aaradhakName}</h4>
                          <p className="text-xs text-stone-600">
                            ગુરુભગવંત: {item.guruBhagwantName}
                          </p>
                          <p className="text-xs text-stone-600 mt-1">
                            સંપર્ક: <span className="font-semibold">{item.mobileNumber}</span>
                          </p>
                        </div>
                        <a
                          href={buildWhatsAppUrl(item)}
                          target="_blank"
                          rel="noreferrer"
                          className="inline-flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold py-2 px-3 rounded-lg mt-2"
                        >
                          <MessageCircle className="w-4 h-4" /> WhatsApp પર રિમાઇન્ડર મોકલો
                        </a>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ================= TAB 2: DIRECTORY & GURU BHAGWANT FILTER ================= */}
        {activeTab === 'directory' && (
          <div className="space-y-5">
            <div className="bg-white rounded-2xl shadow-sm border border-stone-200 p-4 space-y-4">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                <div className="flex items-center gap-2 text-maroon-900 font-bold text-sm">
                  <Filter className="w-4 h-4" />
                  <span>પૂ. ગુરુભગવંતશ્રીનું નામ મુજબ ફિલ્ટર કરો (One-Click Filter):</span>
                </div>

                <div className="relative w-full md:w-72">
                  <Search className="w-4 h-4 text-stone-400 absolute left-3 top-2.5" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="નામ, સમુદાય, મોબાઈલ અથવા શહેર શોધો..."
                    className="w-full pl-9 pr-3 py-1.5 text-sm border border-stone-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-maroon-700"
                  />
                </div>
              </div>

              <div className="flex flex-wrap gap-2">
                <button
                  onClick={() => setSelectedGuruFilter('ALL')}
                  className={`px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                    selectedGuruFilter === 'ALL'
                      ? 'bg-maroon-800 text-white shadow'
                      : 'bg-stone-100 text-stone-700 hover:bg-stone-200'
                  }`}
                >
                  બધા ગુરુભગવંત ({entries.length})
                </button>
                {uniqueGurus.map((guru) => {
                  const count = entries.filter((e) => e.guruBhagwantName === guru).length;
                  return (
                    <button
                      key={guru}
                      onClick={() => setSelectedGuruFilter(guru)}
                      className={`px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                        selectedGuruFilter === guru
                          ? 'bg-maroon-800 text-white shadow'
                          : 'bg-amber-50 text-maroon-900 border border-amber-200 hover:bg-amber-100'
                      }`}
                    >
                      {guru} ({count})
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {filteredEntries.map((item, idx) => {
                const reminderDate = addDays(item.allottedStartDate, -3);
                return (
                  <div
                    key={item.id}
                    className="bg-white rounded-2xl shadow-sm border border-stone-200 p-5 flex flex-col justify-between gap-3"
                  >
                    <div>
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <span className="text-xs font-bold text-maroon-800 bg-maroon-50 px-2 py-0.5 rounded">
                            #{item.serialNo || idx + 1} • {item.samuday || 'સમુદાય'}
                          </span>
                          <h3 className="text-base font-bold text-stone-900 mt-1.5">
                            {item.aaradhakName}
                          </h3>
                          <p className="text-xs text-stone-600">
                            પૂ. ગુરુભગવંત: <span className="font-semibold">{item.guruBhagwantName}</span>
                            {item.dikshaParyay && ` • દિક્ષા પર્યાય: ${item.dikshaParyay}`}
                          </p>
                        </div>
                      </div>

                      <div className="mt-3 bg-stone-50 rounded-xl p-3 text-xs space-y-1.5 border border-stone-100">
                        <div>
                          <span className="text-stone-500">આપના અનુકૂળ દિવસો:</span>{' '}
                          <span className="font-semibold text-stone-800">
                            {item.requestedDateText || '-'}
                          </span>
                        </div>
                        {item.ourDaysText && (
                          <div>
                            <span className="text-stone-500">અથવા અમારા દિવસો:</span>{' '}
                            <span className="text-stone-700">{item.ourDaysText}</span>
                          </div>
                        )}
                        <div className="flex items-center gap-1.5 text-stone-700">
                          <Phone className="w-3.5 h-3.5 text-maroon-700" />
                          <span className="font-semibold">{item.mobileNumber || '-'}</span>
                        </div>
                        <div className="flex items-start gap-1.5 text-stone-600">
                          <MapPin className="w-3.5 h-3.5 text-maroon-700 shrink-0 mt-0.5" />
                          <span>{item.address || '-'}</span>
                        </div>
                      </div>
                    </div>

                    <div className="pt-2 border-t border-stone-100 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                      <div>
                        <label className="block text-[11px] font-semibold text-stone-500">
                          ફાળવેલ તારીખ (Preview / Google Sheet Col P):
                        </label>
                        <input
                          type="date"
                          value={item.allottedStartDate || ''}
                          onChange={(e) => handleUpdateAllottedDate(item.id, e.target.value)}
                          className="mt-0.5 text-xs font-semibold border border-stone-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-maroon-700"
                        />
                      </div>

                      {item.allottedStartDate && (
                        <div className="text-right">
                          <div className="text-[11px] text-amber-800 font-semibold">
                            ૩-દિવસ પહેલા એલર્ટ: {formatDisplayDate(reminderDate)}
                          </div>
                          <a
                            href={buildWhatsAppUrl(item)}
                            target="_blank"
                            rel="noreferrer"
                            className="mt-1 inline-flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg"
                          >
                            <MessageCircle className="w-3.5 h-3.5" /> WhatsApp Alert
                          </a>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ================= TAB 3: 3-DAY ADVANCE WHATSAPP ALERTS ================= */}
        {activeTab === 'alerts' && (
          <div className="space-y-4">
            <div className="bg-amber-50 border border-amber-200 rounded-2xl p-4">
              <h2 className="text-base font-bold text-amber-950">
                ૩-દિવસ પહેલા WhatsApp નોટિફિકેશન શેડ્યૂલ (3-Day Advance Reminder Queue)
              </h2>
              <p className="text-xs text-amber-800 mt-1">
                Google Sheet માં Column P (ફાળવેલ શરૂઆત તારીખ) ભરેલ હોય તેવા તમામ મહારાજ સાહેબની ૩-દિવસ પહેલાની એલર્ટ તારીખ અહીં આપમેળે દેખાશે.
              </p>
            </div>

            <div className="bg-white rounded-2xl shadow-sm border border-stone-200 overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-stone-50 border-b border-stone-200 text-xs font-bold text-stone-700">
                      <th className="p-3.5">એલર્ટ તારીખ (૩ દિવસ પહેલા)</th>
                      <th className="p-3.5">ફાળવેલ ૩ દિવસ & તિથિ</th>
                      <th className="p-3.5">આરાધક પૂજ્યશ્રી</th>
                      <th className="p-3.5">મોબાઈલ નંબર</th>
                      <th className="p-3.5">સ્થિતિ (Status)</th>
                      <th className="p-3.5 text-right">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-stone-100 text-sm">
                    {alertList.map((item) => {
                      const alertKey = `${item.aaradhakName}_${item.allottedStartDate}`;
                      const isSent = !!sentIds[alertKey];
                      return (
                        <tr key={alertKey} className="hover:bg-stone-50/80">
                          <td className="p-3.5 font-bold text-maroon-800">
                            {formatDisplayDate(item.reminderDate)}
                          </td>
                          <td className="p-3.5">
                            <div className="font-semibold text-stone-900">
                              {formatDisplayDate(item.allottedStartDate)} -{' '}
                              {formatDisplayDate(item.endDate)}
                            </div>
                            <div className="text-xs text-amber-700 font-medium">{item.tithiLabel}</div>
                          </td>
                          <td className="p-3.5">
                            <div className="font-bold text-stone-900">{item.aaradhakName}</div>
                            <div className="text-xs text-stone-500">{item.guruBhagwantName}</div>
                          </td>
                          <td className="p-3.5 font-mono text-xs">{item.mobileNumber}</td>
                          <td className="p-3.5">
                            <button
                              onClick={() => toggleReminderSent(alertKey)}
                              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold ${
                                isSent
                                  ? 'bg-emerald-100 text-emerald-800'
                                  : 'bg-amber-100 text-amber-800'
                              }`}
                            >
                              {isSent ? (
                                <>
                                  <CheckCircle2 className="w-3.5 h-3.5" /> મોકલાઈ ગયેલ
                                </>
                              ) : (
                                <>
                                  <Clock className="w-3.5 h-3.5" /> બાકી (Pending)
                                </>
                              )}
                            </button>
                          </td>
                          <td className="p-3.5 text-right">
                            <a
                              href={buildWhatsAppUrl(item)}
                              target="_blank"
                              rel="noreferrer"
                              onClick={() => {
                                if (!isSent) toggleReminderSent(alertKey);
                              }}
                              className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-2 rounded-lg shadow-sm"
                            >
                              <MessageCircle className="w-4 h-4" /> Send WhatsApp
                            </a>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ================= TAB 4: ADD NEW ENTRIES VIA LIVE GOOGLE SHEET ================= */}
        {activeTab === 'add' && (
          <div className="bg-white rounded-2xl shadow-sm border border-stone-200 p-6 max-w-2xl mx-auto space-y-4">
            <h2 className="text-lg font-bold text-maroon-900 border-b border-stone-100 pb-3">
              નવી વિગતો કેવી રીતે ઉમેરવી? (Live Google Sheet Sync)
            </h2>
            <p className="text-sm text-stone-700 leading-relaxed">
              આ વેબસાઈટ સીધી તમારી <strong>Google Sheet</strong> સાથે જોડાયેલી છે. જ્યારે પણ નવા ૨૦–૩૦ મહારાજ સાહેબની વિગતો આવે ત્યારે:
            </p>
            <ol className="list-decimal list-inside text-sm text-stone-700 space-y-2 bg-stone-50 p-4 rounded-xl border border-stone-200">
              <li>
                તમારી Google Sheet ખોલો અને નીચેની ખાલી હરોળ (Rows) માં નવી વિગતો લખો.
              </li>
              <li>
                <strong>Column P (ફાળવેલ શરૂઆત તારીખ)</strong> માં <code className="bg-white px-1.5 py-0.5 rounded border text-xs">YYYY-MM-DD</code> (દા.ત. <code className="bg-white px-1.5 py-0.5 rounded border text-xs">2027-02-14</code>) લખો.
              </li>
              <li>
                અહીં ઉપર આપેલ <strong>"Sync Google Sheet"</strong> બટન દબાવો (અથવા પેજ રિફ્રેશ કરો) — બધા જ નામ કેલેન્ડર, ગુરુ ફિલ્ટર અને WhatsApp એલર્ટમાં આપમેળે આવી જશે!
              </li>
            </ol>

            <div className="flex flex-wrap gap-3 pt-2">
              <a
                href="https://drive.google.com"
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-2 bg-maroon-800 hover:bg-maroon-900 text-white text-sm font-semibold px-4 py-2.5 rounded-lg shadow"
              >
                <ExternalLink className="w-4 h-4" /> Google Drive / Sheet ખોલો
              </a>
              <button
                onClick={() => {
                  fetchLiveGoogleSheet();
                  setActiveTab('calendar');
                }}
                className="inline-flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold px-4 py-2.5 rounded-lg shadow"
              >
                <RefreshCw className="w-4 h-4" /> હમણાં Sync કરીને કેલેન્ડર જુઓ
              </button>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}