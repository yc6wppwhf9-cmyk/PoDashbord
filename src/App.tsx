import { useState, useRef, useEffect, useMemo } from 'react';
import * as xlsx from 'xlsx';
import { Calendar, AlertCircle, Clock, UploadCloud, FileSpreadsheet, X, RefreshCw, Search } from 'lucide-react';
import { isToday, isWithinInterval, addDays, startOfDay, differenceInDays } from 'date-fns';

interface PODetail {
  poNo: string;
  date: string;
  supplier: string;
  status: string;
  dueDate: string;
  orderQty: number;
  receivedQty: number;
  creator: string;
  totalDays: number | string;
  dateRaw: Date | null;
  dueDateRaw: Date | null;
  remark: string;
}

interface POMetrics {
  remarksPOs: PODetail[];
  dueWithin7Days: PODetail[];
  dueToday: PODetail[];
  overduePOs: PODetail[];
  openPOs: PODetail[];
}

type FilterType = keyof POMetrics | null;

function App() {
  const [metrics, setMetrics] = useState<POMetrics | null>(null);
  const [loading, setLoading] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  const [activeFilter, setActiveFilter] = useState<FilterType>(null);
  const [poRawDataMap, setPoRawDataMap] = useState<Map<string, any[]>>(new Map());
  const [searchQuery, setSearchQuery] = useState('');
  const [columnFilters, setColumnFilters] = useState<{ [key: string]: string }>({});
  const [dateRange, setDateRange] = useState<{ start: string; end: string }>({ start: '', end: '' });
  const [remarksMap, setRemarksMap] = useState<{ [key: string]: string }>({});

  // Attempt to auto-sync on first load
  useEffect(() => {
    handleSyncEmail();
    fetchRemarks();
  }, []);

  const fetchRemarks = async () => {
    try {
      const isDev = import.meta.env.DEV;
      const baseUrl = isDev ? 'http://localhost:3000' : '';
      const response = await fetch(`${baseUrl}/api/remarks`);
      if (response.ok) {
        const data = await response.json();
        setRemarksMap(data);
      }
    } catch (err) {
      console.error("Failed to fetch remarks:", err);
    }
  };

  const handleRemarkChange = async (poNo: string, remark: string) => {
    setRemarksMap(prev => ({ ...prev, [poNo]: remark }));
    
    try {
      const isDev = import.meta.env.DEV;
      const baseUrl = isDev ? 'http://localhost:3000' : '';
      await fetch(`${baseUrl}/api/remarks`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ poNo, remark })
      });
    } catch (err) {
      console.error("Failed to save remark:", err);
    }
  };

  const processArrayBuffer = (arrayBuffer: ArrayBuffer) => {
    const workbook = xlsx.read(arrayBuffer, { type: 'buffer', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    const rawData = xlsx.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
    let headerRowIndex = 0;
    
    // Robust header detection
    for (let i = 0; i < Math.min(20, rawData.length); i++) {
      if (!rawData[i]) continue;
      const rowHasPONo = rawData[i].some(cell => typeof cell === 'string' && (cell.includes('PO No') || cell.includes('PO Number')));
      if (rowHasPONo) {
        headerRowIndex = i;
        break;
      }
    }

    const data = xlsx.utils.sheet_to_json(worksheet, { range: headerRowIndex }) as any[];

    if (data.length === 0) {
      throw new Error('No data found in the Excel file.');
    }

    const poSet = new Set<string>();
    const due7DaysSet = new Set<string>();
    const dueTodaySet = new Set<string>();
    const overduePOSet = new Set<string>();
    const remarksPOSet = new Set<string>();
    const openPOSet = new Set<string>();
    
    const poDetailsMap = new Map<string, PODetail>();
    const rawMap = new Map<string, any[]>();

    const today = startOfDay(new Date());
    const sevenDaysFromNow = addDays(today, 7);

    data.forEach(row => {
      const poNo = row['PO No.'] || row['PO No'] || row['Purchase Order'] || row['PO Number'];
      if (!poNo) return;

      if (!rawMap.has(poNo)) {
        rawMap.set(poNo, []);
      }
      rawMap.get(poNo)!.push(row);

      // Check Due Dates - Prioritize 'Valid Till' as requested
      const rawDueDate = row['Valid Till'] || row['Due Date'] || row['Delivery Date'] || row['SCHEDULE_DATE'] || row['Shedule Date'];
      let dueDateObj: Date | null = null;
      if (rawDueDate) {
        if (rawDueDate instanceof Date) {
          // Use UTC components to prevent timezone shifts
          dueDateObj = new Date(rawDueDate.getUTCFullYear(), rawDueDate.getUTCMonth(), rawDueDate.getUTCDate());
        } else {
          dueDateObj = startOfDay(new Date(rawDueDate));
        }
        if (isNaN(dueDateObj.getTime())) {
           dueDateObj = null;
        }
      }

      if (!poDetailsMap.has(poNo)) {
         const date = row['Order Date'] || row['PO Date'] || row['PO Creation Date'] || '';
         const supplier = row['Party Name'] || row['Supplier'] || row['Party Code'] || '';
         const status = row['PO Status'] || '';
         const creator = row['Created By'] || row['Creator'] || '-';
         
         let dateObj: Date | null = null;
         if (date) {
           if (date instanceof Date) {
             dateObj = new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
           } else {
             dateObj = startOfDay(new Date(date));
           }
           if (isNaN(dateObj.getTime())) dateObj = null;
         }
         
         let totalDays: number | string = '-';
         if (dateObj && dueDateObj) {
           totalDays = differenceInDays(dueDateObj, dateObj);
         }
         
         poDetailsMap.set(poNo, {
             poNo: String(poNo),
             date: dateObj ? dateObj.toLocaleDateString() : '-',
             supplier: String(supplier),
             status: String(status),
             dueDate: dueDateObj ? dueDateObj.toLocaleDateString() : '-',
             orderQty: 0,
             receivedQty: 0,
             creator: String(creator),
             totalDays,
             dateRaw: dateObj,
             dueDateRaw: dueDateObj,
             remark: remarksMap[String(poNo)] || ''
         });
      }

      const poDetail = poDetailsMap.get(poNo)!;
      if (poDetail.remark) {
        remarksPOSet.add(String(poNo));
      }
      poDetail.orderQty += Number(row['Order Qty'] || row['PO Original Qty'] || row['ORDER_QUATITY'] || 0);
      poDetail.receivedQty += Number(row['PO Received Qty'] || row['Received Qty'] || 0);

      poSet.add(poNo);

      const status = row['PO Status'];
      const balanceQty = row['Balance Qty'] || row['PO Pending Qty'] || 0;
      
      let isOpen = false;
      if (balanceQty > 0) isOpen = true;
      if (status && typeof status === 'string' && status.toLowerCase().includes('open')) isOpen = true;
      if (status && typeof status === 'string' && !status.toLowerCase().includes('total received/cancelled') && !status.toLowerCase().includes('closed')) isOpen = true; 
      
      if (status && typeof status === 'string' && status.toLowerCase().includes('cancelled')) isOpen = false;
      if (status && typeof status === 'string' && status.toLowerCase().includes('total received')) isOpen = false;
      
      if (balanceQty > 0) isOpen = true;

      if (isOpen) {
        openPOSet.add(poNo);
        if (dueDateObj && dueDateObj < today) {
          overduePOSet.add(poNo);
        }
      }

      if (dueDateObj) {
        if (isToday(dueDateObj)) {
          dueTodaySet.add(poNo);
        }
        
        if (isWithinInterval(dueDateObj, { start: today, end: sevenDaysFromNow })) {
          due7DaysSet.add(poNo);
        }
      }
    });

    setMetrics({
      remarksPOs: Array.from(remarksPOSet).map(id => poDetailsMap.get(id)!),
      dueWithin7Days: Array.from(due7DaysSet).map(id => poDetailsMap.get(id)!),
      dueToday: Array.from(dueTodaySet).map(id => poDetailsMap.get(id)!),
      overduePOs: Array.from(overduePOSet).map(id => poDetailsMap.get(id)!),
      openPOs: Array.from(openPOSet).map(id => poDetailsMap.get(id)!)
    });
    setPoRawDataMap(rawMap);
  };

  const processExcelData = async (file: File) => {
    setLoading(true);
    setError(null);
    setFileName(file.name);
    setActiveFilter(null);

    try {
      const arrayBuffer = await file.arrayBuffer();
      processArrayBuffer(arrayBuffer);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'An error occurred while processing the file.');
      setMetrics(null);
    } finally {
      setLoading(false);
    }
  };

  const handleSyncEmail = async (forceRefresh = false) => {
    setSyncing(true);
    setLoading(true);
    setError(null);
    setActiveFilter(null);
    setFileName('Synced from Email');

    try {
      // In development, the backend runs on port 3000. In production, it's the same host.
      const isDev = import.meta.env.DEV;
      const baseUrl = isDev ? 'http://localhost:3000' : '';
      const url = `${baseUrl}/api/po-data${forceRefresh ? '?forceRefresh=true' : ''}`;
      
      const response = await fetch(url);
      
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || 'Failed to sync with email backend. Please check Render environment variables.');
      }

      const arrayBuffer = await response.arrayBuffer();
      processArrayBuffer(arrayBuffer);
    } catch (err: any) {
      console.error(err);
      setError(err.message);
      setFileName(null);
    } finally {
      setLoading(false);
      setSyncing(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processExcelData(file);
    }
  };

  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  const handleMetricClick = (filter: FilterType) => {
    setActiveFilter(activeFilter === filter ? null : filter);
  };

  const activeData = activeFilter && metrics ? metrics[activeFilter] : [];

  const filteredData = useMemo(() => {
    let data = activeData;

    if (searchQuery) {
      const lowerQuery = searchQuery.toLowerCase();
      data = data.filter(po => 
        Object.values(po).some(val => 
          String(val).toLowerCase().includes(lowerQuery)
        )
      );
    }

    Object.entries(columnFilters).forEach(([key, filterValue]) => {
      if (filterValue) {
        const lowerFilter = filterValue.toLowerCase();
        data = data.filter(po => 
          String(po[key as keyof PODetail] || '').toLowerCase().includes(lowerFilter)
        );
      }
    });

    if (dateRange.start || dateRange.end) {
      data = data.filter(po => {
        if (!po.dateRaw) return false;
        const poDate = startOfDay(po.dateRaw);
        
        if (dateRange.start && dateRange.end) {
          return isWithinInterval(poDate, {
            start: startOfDay(new Date(dateRange.start)),
            end: startOfDay(new Date(dateRange.end))
          });
        } else if (dateRange.start) {
          return poDate >= startOfDay(new Date(dateRange.start));
        } else if (dateRange.end) {
          return poDate <= startOfDay(new Date(dateRange.end));
        }
        return true;
      });
    }

    return data;
  }, [activeData, searchQuery, columnFilters]);

  const getMetricTitle = (filter: FilterType) => {
    switch (filter) {
        case 'remarksPOs': return 'POs with Remarks';
        case 'openPOs': return 'Open Purchase Orders';
        case 'dueToday': return 'POs Due Today';
        case 'dueWithin7Days': return 'POs Due in 7 Days';
        case 'overduePOs': return 'Overdue Purchase Orders';
        default: return 'Analytics Overview';
    }
  };

  const openPOPreview = (poNo: string) => {
    const rows = poRawDataMap.get(poNo);
    if (!rows || rows.length === 0) return;
    
    // Generate the HTML table using SheetJS
    const worksheet = xlsx.utils.json_to_sheet(rows);
    const htmlTable = xlsx.utils.sheet_to_html(worksheet);
    
    const html = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>PO Detail - ${poNo}</title>
          <style>
            body { 
              margin: 0; 
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; 
              background: #f9fbfd; 
              color: #3c4043;
            }
            .top-bar {
              background: #fff;
              height: 64px;
              display: flex;
              align-items: center;
              padding: 0 20px;
              border-bottom: 1px solid #e0e0e0;
              position: sticky;
              top: 0;
              z-index: 100;
            }
            .logo {
              width: 32px;
              height: 32px;
              background: #0f9d58;
              border-radius: 4px;
              display: flex;
              align-items: center;
              justify-content: center;
              margin-right: 12px;
            }
            .logo svg { fill: white; width: 20px; height: 20px; }
            .title-info { flex-grow: 1; }
            .filename { font-size: 18px; font-weight: 400; color: #202124; margin: 0; }
            .file-meta { font-size: 12px; color: #5f6368; margin-top: 2px; }
            
            .content {
              padding: 20px;
              overflow: auto;
              height: calc(100vh - 104px);
            }
            
            table {
              border-collapse: collapse;
              background: white;
              box-shadow: 0 1px 3px rgba(60,64,67,0.3);
              min-width: 100%;
            }
            
            th, td {
              border: 1px solid #e0e0e0;
              padding: 8px 12px;
              font-size: 13px;
              height: 24px;
            }
            
            th {
              background: #f8f9fa;
              color: #5f6368;
              font-weight: 500;
              text-align: center;
              position: sticky;
              top: 0;
              z-index: 10;
            }
            
            tr:first-child td {
              background: #f8f9fa;
              color: #5f6368;
              font-weight: bold;
              text-align: center;
            }
            
            td {
              color: #3c4043;
            }
            
            tr:hover td {
              background-color: #f1f3f4;
            }
            
            .excel-grid-header {
              background: #f8f9fa;
              width: 40px;
              text-align: center;
              border-right: 2px solid #e0e0e0;
              color: #5f6368;
              font-weight: bold;
            }
          </style>
        </head>
        <body>
          <div class="top-bar">
            <div class="logo">
              <svg viewBox="0 0 24 24"><path d="M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M15.8,20H14L12,16.6L10,20H8.2L11.1,15.5L8.2,11H10L12,14.4L14,11H15.8L12.9,15.5L15.8,20M13,9V3.5L18.5,9H13Z" /></svg>
            </div>
            <div class="title-info">
              <h1 class="filename">Purchase Order Details: ${poNo}</h1>
              <div class="file-meta">Viewing spreadsheet preview</div>
            </div>
          </div>
          <div class="content">
            ${htmlTable}
          </div>
        </body>
      </html>
    `;
    
    const newWindow = window.open('', '_blank');
    if (newWindow) {
      newWindow.document.write(html);
      newWindow.document.close();
    }
  };

  return (
    <div className="dashboard-container">
      <div className="header" style={{ marginBottom: '2rem' }}>
        <div>
          <h1>PO Dashboard</h1>
          <p>Sync automatically with your email or upload manually</p>
        </div>
      </div>

      <div style={{ marginBottom: '3rem', display: 'flex', gap: '1rem', alignItems: 'center', flexWrap: 'wrap' }}>
        <button 
          onClick={() => handleSyncEmail(true)}
          disabled={syncing}
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '0.5rem',
            background: 'var(--primary-color)',
            color: 'white',
            border: 'none',
            padding: '0.75rem 1.5rem',
            borderRadius: '0.5rem',
            fontSize: '1rem',
            fontWeight: '600',
            cursor: syncing ? 'not-allowed' : 'pointer',
            transition: 'background 0.2s',
            opacity: syncing ? 0.7 : 1
          }}
        >
          <RefreshCw size={20} className={syncing ? 'spinning' : ''} />
          {syncing ? 'Syncing...' : 'Sync Latest Email'}
        </button>

        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
          <span style={{ color: 'var(--text-secondary)', padding: '0 0.5rem' }}>OR</span>
          <input 
            type="file" 
            accept=".xlsx, .xls, .csv" 
            style={{ display: 'none' }} 
            ref={fileInputRef}
            onChange={handleFileChange}
          />
          <button 
            onClick={handleUploadClick}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '0.5rem',
              background: 'transparent',
              color: 'var(--accent-color)',
              border: '1px solid var(--accent-color)',
              padding: '0.75rem 1.5rem',
              borderRadius: '0.5rem',
              fontSize: '1rem',
              fontWeight: '600',
              cursor: 'pointer',
              transition: 'background 0.2s',
            }}
          >
            <UploadCloud size={20} />
            Upload File Manually
          </button>
        </div>

        {fileName && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--text-secondary)', marginLeft: 'auto' }}>
            <FileSpreadsheet size={18} />
            <span>{fileName}</span>
          </div>
        )}
      </div>

      <style>{`
        @keyframes spin { 100% { transform: rotate(360deg); } }
        .spinning { animation: spin 1s linear infinite; }
      `}</style>

      {loading && !syncing && (
        <div className="loader-container">
          <div className="spinner"></div>
          <p>Analyzing PO Register...</p>
        </div>
      )}

      {error && (
        <div style={{ color: 'var(--danger-color)', padding: '1.5rem', background: 'var(--glass-bg)', borderRadius: '1rem', border: '1px solid rgba(239, 68, 68, 0.3)', marginBottom: '2rem' }}>
          {error}
        </div>
      )}

      {!loading && !error && metrics && (
        <>
          <div className="metrics-grid">
            <div 
              className={`metric-card remarks ${activeFilter === 'remarksPOs' ? 'active' : ''}`}
              onClick={() => handleMetricClick('remarksPOs')}
            >
              <div className="metric-header">
                <span className="metric-title">With Remarks</span>
                <div className="metric-icon">
                  <FileSpreadsheet size={20} />
                </div>
              </div>
              <div className="metric-value">{metrics.remarksPOs.length}</div>
            </div>

            <div 
              className={`metric-card open ${activeFilter === 'openPOs' ? 'active' : ''}`}
              onClick={() => handleMetricClick('openPOs')}
            >
              <div className="metric-header">
                <span className="metric-title">Open POs</span>
                <div className="metric-icon">
                  <AlertCircle size={20} />
                </div>
              </div>
              <div className="metric-value">{metrics.openPOs.length}</div>
            </div>

            <div 
              className={`metric-card overdue ${activeFilter === 'overduePOs' ? 'active' : ''}`}
              onClick={() => handleMetricClick('overduePOs')}
            >
              <div className="metric-header">
                <span className="metric-title">Overdue POs</span>
                <div className="metric-icon">
                  <Clock size={20} />
                </div>
              </div>
              <div className="metric-value">{metrics.overduePOs.length}</div>
            </div>

            <div 
              className={`metric-card due-today ${activeFilter === 'dueToday' ? 'active' : ''}`}
              onClick={() => handleMetricClick('dueToday')}
            >
              <div className="metric-header">
                <span className="metric-title">Due Today</span>
                <div className="metric-icon">
                  <Clock size={20} />
                </div>
              </div>
              <div className="metric-value">{metrics.dueToday.length}</div>
            </div>

            <div 
              className={`metric-card due-seven ${activeFilter === 'dueWithin7Days' ? 'active' : ''}`}
              onClick={() => handleMetricClick('dueWithin7Days')}
            >
              <div className="metric-header">
                <span className="metric-title">Due in 7 Days</span>
                <div className="metric-icon">
                  <Calendar size={20} />
                </div>
              </div>
              <div className="metric-value">{metrics.dueWithin7Days.length}</div>
            </div>
          </div>

          <div className="po-list">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem', paddingBottom: '1rem', borderBottom: '1px solid var(--border-color)', flexWrap: 'wrap', gap: '1rem' }}>
              <h2 style={{ margin: 0, padding: 0, border: 'none' }}>
                {getMetricTitle(activeFilter)}
              </h2>
              {activeFilter && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', flexWrap: 'wrap' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', background: 'rgba(255, 255, 255, 0.03)', padding: '0.25rem 0.75rem', borderRadius: '0.5rem', border: '1px solid var(--border-color)' }}>
                    <span style={{ fontSize: '0.875rem', color: 'var(--text-secondary)' }}>Range:</span>
                    <input 
                      type="date" 
                      value={dateRange.start} 
                      onChange={e => setDateRange(prev => ({ ...prev, start: e.target.value }))}
                      style={{ background: 'transparent', border: 'none', color: 'var(--text-primary)', fontSize: '0.875rem', cursor: 'pointer' }}
                    />
                    <span style={{ color: 'var(--text-secondary)' }}>to</span>
                    <input 
                      type="date" 
                      value={dateRange.end} 
                      onChange={e => setDateRange(prev => ({ ...prev, end: e.target.value }))}
                      style={{ background: 'transparent', border: 'none', color: 'var(--text-primary)', fontSize: '0.875rem', cursor: 'pointer' }}
                    />
                  </div>
                  <div style={{ position: 'relative' }}>
                    <Search size={18} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-secondary)' }} />
                    <input
                      type="text"
                      placeholder="Global Search..."
                      value={searchQuery}
                      onChange={e => setSearchQuery(e.target.value)}
                      style={{
                        padding: '0.5rem 1rem 0.5rem 2.2rem',
                        borderRadius: '0.5rem',
                        border: '1px solid var(--border-color)',
                        background: 'var(--bg-color)',
                        color: 'var(--text-primary)',
                        minWidth: '250px'
                      }}
                    />
                  </div>
                  <button 
                    onClick={() => {
                      setActiveFilter(null);
                      setSearchQuery('');
                      setColumnFilters({});
                      setDateRange({ start: '', end: '' });
                    }}
                    style={{
                      background: 'transparent',
                      border: '1px solid var(--border-color)',
                      color: 'var(--text-secondary)',
                      borderRadius: '0.5rem',
                      padding: '0.5rem 1rem',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '0.5rem',
                      cursor: 'pointer',
                      transition: 'all 0.2s'
                    }}
                    onMouseOver={(e) => {
                      e.currentTarget.style.color = 'var(--text-primary)';
                      e.currentTarget.style.background = 'rgba(255, 255, 255, 0.05)';
                    }}
                    onMouseOut={(e) => {
                      e.currentTarget.style.color = 'var(--text-secondary)';
                      e.currentTarget.style.background = 'transparent';
                    }}
                  >
                    <X size={16} /> Clear Filters
                  </button>
                </div>
              )}
            </div>

            {!activeFilter ? (
              <p style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: '2rem 0' }}>
                Click on any metric card above to view the detailed list of Purchase Orders.
              </p>
            ) : activeData.length === 0 ? (
               <p style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: '2rem 0' }}>
                No records found for this category.
              </p>
            ) : (
              <div style={{ overflowX: 'auto' }}>
                <table>
                  <thead>
                    <tr>
                      <th>
                        <div>PO Number</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.poNo || ''} onChange={e => setColumnFilters(prev => ({ ...prev, poNo: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px' }} />
                      </th>
                      <th>
                        <div>Date</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.date || ''} onChange={e => setColumnFilters(prev => ({ ...prev, date: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px' }} />
                      </th>
                      <th>
                        <div>Supplier</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.supplier || ''} onChange={e => setColumnFilters(prev => ({ ...prev, supplier: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px' }} />
                      </th>
                      <th>
                        <div>Creator</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.creator || ''} onChange={e => setColumnFilters(prev => ({ ...prev, creator: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px' }} />
                      </th>
                      <th>
                        <div>Due Date</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.dueDate || ''} onChange={e => setColumnFilters(prev => ({ ...prev, dueDate: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px' }} />
                      </th>
                      <th style={{ textAlign: 'center' }}>
                        <div>Total Days</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.totalDays || ''} onChange={e => setColumnFilters(prev => ({ ...prev, totalDays: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px', textAlign: 'center' }} />
                      </th>
                      <th>
                        <div>Status</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.status || ''} onChange={e => setColumnFilters(prev => ({ ...prev, status: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px' }} />
                      </th>
                      <th style={{ textAlign: 'right' }}>
                        <div>Order Qty</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.orderQty || ''} onChange={e => setColumnFilters(prev => ({ ...prev, orderQty: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px', textAlign: 'right' }} />
                      </th>
                      <th style={{ textAlign: 'right' }}>
                        <div>Received Qty</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.receivedQty || ''} onChange={e => setColumnFilters(prev => ({ ...prev, receivedQty: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px', textAlign: 'right' }} />
                      </th>
                      <th>
                        <div>Remark</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.remark || ''} onChange={e => setColumnFilters(prev => ({ ...prev, remark: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px' }} />
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.length === 0 ? (
                      <tr>
                        <td colSpan={9} style={{ textAlign: 'center', padding: '2rem 0', color: 'var(--text-secondary)' }}>
                          No records match your search or filters.
                        </td>
                      </tr>
                    ) : filteredData.slice(0, 100).map((po, index) => (
                      <tr key={`${po.poNo}-${index}`}>
                        <td 
                          onClick={() => openPOPreview(po.poNo)}
                          style={{ 
                            fontWeight: 500, 
                            color: 'var(--accent-color)', 
                            cursor: 'pointer',
                            textDecoration: 'underline'
                          }}
                          title="Click to view PO details in spreadsheet view"
                        >
                          {po.poNo}
                        </td>
                        <td>{po.date}</td>
                        <td>{po.supplier}</td>
                        <td>{po.creator}</td>
                        <td>{po.dueDate}</td>
                        <td style={{ textAlign: 'center' }}>
                          {po.totalDays !== '-' ? (
                            <span style={{ fontWeight: 600, color: Number(po.totalDays) < 0 ? 'var(--danger-color)' : (Number(po.totalDays) <= 7 ? 'var(--warning-color)' : 'var(--text-primary)') }}>
                              {po.totalDays}
                            </span>
                          ) : '-'}
                        </td>
                        <td>
                          <span className={`badge ${po.status.toLowerCase().includes('open') ? 'open' : po.status.toLowerCase().includes('cancelled') ? 'danger' : 'closed'}`}>
                            {po.status || 'Unknown'}
                          </span>
                        </td>
                        <td style={{ textAlign: 'right' }}>{po.orderQty}</td>
                        <td style={{ textAlign: 'right' }}>{po.receivedQty}</td>
                        <td>
                          <input 
                            type="text" 
                            value={remarksMap[po.poNo] || ''} 
                            onChange={(e) => handleRemarkChange(po.poNo, e.target.value)}
                            placeholder="Add remark..."
                            style={{
                              width: '100%',
                              padding: '4px 8px',
                              borderRadius: '4px',
                              border: '1px solid var(--border-color)',
                              background: 'rgba(255, 255, 255, 0.05)',
                              color: 'var(--text-primary)',
                              fontSize: '12px'
                            }}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {filteredData.length > 100 && (
                  <p style={{ color: 'var(--text-secondary)', textAlign: 'center', marginTop: '1.5rem', fontSize: '0.875rem' }}>
                    Showing top 100 results out of {filteredData.length}.
                  </p>
                )}
              </div>
            )}
          </div>
        </>
      )}

      {!loading && !error && !metrics && !syncing && (
        <div className="po-list" style={{ textAlign: 'center', padding: '4rem 2rem' }}>
          <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '1rem', color: 'var(--text-secondary)' }}>
            <FileSpreadsheet size={48} opacity={0.5} />
          </div>
          <h2>No Data Found</h2>
          <p style={{ color: 'var(--text-secondary)' }}>
            Configure your Render Environment Variables to sync automatically, or upload your PO Register manually.
          </p>
        </div>
      )}
    </div>
  );
}

export default App;
