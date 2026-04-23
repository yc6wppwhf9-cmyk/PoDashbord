import { useState, useRef, useEffect, useMemo } from 'react';
import * as xlsx from 'xlsx';
import { ShoppingCart, Calendar, AlertCircle, Clock, UploadCloud, FileSpreadsheet, X, RefreshCw, Search } from 'lucide-react';
import { isToday, isWithinInterval, addDays, startOfDay, differenceInDays } from 'date-fns';

interface PODetail {
  poNo: string;
  date: string;
  supplier: string;
  status: string;
  dueDate: string;
  qty: number;
  creator: string;
  totalDays: number | string;
}

interface POMetrics {
  totalPOs: PODetail[];
  dueWithin7Days: PODetail[];
  dueToday: PODetail[];
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

  // Attempt to auto-sync on first load
  useEffect(() => {
    handleSyncEmail();
  }, []);

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

      // Check Due Dates
      const rawDueDate = row['Due Date'] || row['Delivery Date'] || row['SCHEDULE_DATE'] || row['Shedule Date'] || row['Valid Till'];
      let dueDateObj: Date | null = null;
      if (rawDueDate) {
        if (rawDueDate instanceof Date) {
          dueDateObj = startOfDay(rawDueDate);
        } else {
          dueDateObj = startOfDay(new Date(rawDueDate));
        }
        if (isNaN(dueDateObj.getTime())) {
           dueDateObj = null;
        }
      }

      if (!poDetailsMap.has(poNo)) {
         const qty = row['ORDER_QUATITY'] || row['Order Qty'] || row['PO Qty'] || row['PO Original Qty'] || 0;
         const date = row['PO Creation Date'] || row['PO Date'] || row['Order Date'] || '';
         const supplier = row['Party Name'] || row['Supplier'] || row['Party Code'] || '';
         const status = row['PO Status'] || '';
         const creator = row['Created By'] || row['Creator'] || '-';
         
         let dateObj: Date | null = null;
         if (date) {
           dateObj = date instanceof Date ? startOfDay(date) : startOfDay(new Date(date));
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
             qty: Number(qty) || 0,
             creator: String(creator),
             totalDays
         });
      }

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
      totalPOs: Array.from(poSet).map(id => poDetailsMap.get(id)!),
      dueWithin7Days: Array.from(due7DaysSet).map(id => poDetailsMap.get(id)!),
      dueToday: Array.from(dueTodaySet).map(id => poDetailsMap.get(id)!),
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

    return data;
  }, [activeData, searchQuery, columnFilters]);

  const getMetricTitle = (filter: FilterType) => {
    switch (filter) {
        case 'totalPOs': return 'Total Purchase Orders';
        case 'openPOs': return 'Open Purchase Orders';
        case 'dueToday': return 'POs Due Today';
        case 'dueWithin7Days': return 'POs Due in 7 Days';
        default: return 'Analytics Overview';
    }
  };

  const downloadPOAsExcel = (poNo: string) => {
    const rows = poRawDataMap.get(poNo);
    if (!rows || rows.length === 0) return;
    
    const worksheet = xlsx.utils.json_to_sheet(rows);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'PO_Details');
    
    xlsx.writeFile(workbook, `PO_${poNo}.xlsx`);
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
              className={`metric-card total ${activeFilter === 'totalPOs' ? 'active' : ''}`}
              onClick={() => handleMetricClick('totalPOs')}
            >
              <div className="metric-header">
                <span className="metric-title">Total POs</span>
                <div className="metric-icon">
                  <ShoppingCart size={20} />
                </div>
              </div>
              <div className="metric-value">{metrics.totalPOs.length}</div>
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
          </div>

          <div className="po-list">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem', paddingBottom: '1rem', borderBottom: '1px solid var(--border-color)', flexWrap: 'wrap', gap: '1rem' }}>
              <h2 style={{ margin: 0, padding: 0, border: 'none' }}>
                {getMetricTitle(activeFilter)}
              </h2>
              {activeFilter && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', flexWrap: 'wrap' }}>
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
                        <div>Qty</div>
                        <input type="text" placeholder="Filter..." value={columnFilters.qty || ''} onChange={e => setColumnFilters(prev => ({ ...prev, qty: e.target.value }))} style={{ width: '100%', padding: '4px', marginTop: '4px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '12px', textAlign: 'right' }} />
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.length === 0 ? (
                      <tr>
                        <td colSpan={8} style={{ textAlign: 'center', padding: '2rem 0', color: 'var(--text-secondary)' }}>
                          No records match your search or filters.
                        </td>
                      </tr>
                    ) : filteredData.slice(0, 100).map((po, index) => (
                      <tr key={`${po.poNo}-${index}`}>
                        <td 
                          onClick={() => downloadPOAsExcel(po.poNo)}
                          style={{ 
                            fontWeight: 500, 
                            color: 'var(--accent-color)', 
                            cursor: 'pointer',
                            textDecoration: 'underline'
                          }}
                          title="Click to download PO details as Excel"
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
                        <td style={{ textAlign: 'right' }}>{po.qty}</td>
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
