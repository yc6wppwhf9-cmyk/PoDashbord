import { useState, useRef, useEffect } from 'react';
import * as xlsx from 'xlsx';
import { ShoppingCart, Calendar, AlertCircle, Clock, UploadCloud, FileSpreadsheet, X, RefreshCw } from 'lucide-react';
import { isToday, isWithinInterval, addDays, startOfDay } from 'date-fns';

interface PODetail {
  poNo: string;
  date: string;
  supplier: string;
  status: string;
  dueDate: string;
  qty: number;
  creator: string;
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
         
         poDetailsMap.set(poNo, {
             poNo: String(poNo),
             date: date instanceof Date ? date.toLocaleDateString() : (date ? new Date(date).toLocaleDateString() : '-'),
             supplier: String(supplier),
             status: String(status),
             dueDate: dueDateObj ? dueDateObj.toLocaleDateString() : '-',
             qty: Number(qty) || 0,
             creator: String(creator)
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

  const getMetricTitle = (filter: FilterType) => {
    switch (filter) {
        case 'totalPOs': return 'Total Purchase Orders';
        case 'openPOs': return 'Open Purchase Orders';
        case 'dueToday': return 'POs Due Today';
        case 'dueWithin7Days': return 'POs Due in 7 Days';
        default: return 'Analytics Overview';
    }
  };

  const openPOInNewTab = (poNo: string) => {
    const rows = poRawDataMap.get(poNo);
    if (!rows || rows.length === 0) return;
    
    const worksheet = xlsx.utils.json_to_sheet(rows);
    const htmlTable = xlsx.utils.sheet_to_html(worksheet);
    
    const html = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>PO Details - ${poNo}</title>
          <style>
            body { font-family: 'Inter', -apple-system, sans-serif; padding: 30px; background: #f8fafc; color: #0f172a; }
            .header { border-bottom: 2px solid #e2e8f0; padding-bottom: 15px; margin-bottom: 25px; }
            h2 { margin: 0; font-size: 24px; font-weight: 600; }
            .po-badge { background: #3b82f6; color: white; padding: 4px 12px; border-radius: 6px; font-size: 20px; }
            table { border-collapse: collapse; width: 100%; background: white; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 8px; overflow: hidden; }
            th, td { border: 1px solid #e2e8f0; padding: 12px 16px; text-align: left; font-size: 14px; }
            th { background-color: #f1f5f9; font-weight: 600; color: #475569; white-space: nowrap; }
            tr:hover td { background-color: #f8fafc; }
          </style>
        </head>
        <body>
          <div class="header">
            <h2>Purchase Order Details: <span class="po-badge">${poNo}</span></h2>
          </div>
          <div style="overflow-x: auto;">
            ${htmlTable}
          </div>
        </body>
      </html>
    `;
    
    const newWindow = window.open('', '_blank');
    if (newWindow) {
      newWindow.document.write(html);
      newWindow.document.close();
    } else {
      alert('Please allow popups to view the PO details in a new tab.');
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
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem', paddingBottom: '1rem', borderBottom: '1px solid var(--border-color)' }}>
              <h2 style={{ margin: 0, padding: 0, border: 'none' }}>
                {getMetricTitle(activeFilter)}
              </h2>
              {activeFilter && (
                <button 
                  onClick={() => setActiveFilter(null)}
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
                  <X size={16} /> Clear Filter
                </button>
              )}
            </div>

            {!activeFilter ? (
              <p style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: '2rem 0' }}>
                Click on any metric card above to view the detailed list of Purchase Orders.
              </p>
            ) : activeData.length === 0 ? (
               <p style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: '2rem 0' }}>
                No records found for this filter.
              </p>
            ) : (
              <div style={{ overflowX: 'auto' }}>
                <table>
                  <thead>
                    <tr>
                      <th>PO Number</th>
                      <th>Date</th>
                      <th>Supplier</th>
                      <th>Creator</th>
                      <th>Due Date</th>
                      <th>Status</th>
                      <th style={{ textAlign: 'right' }}>Qty</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activeData.slice(0, 100).map((po, index) => (
                      <tr key={`${po.poNo}-${index}`}>
                        <td 
                          onClick={() => openPOInNewTab(po.poNo)}
                          style={{ 
                            fontWeight: 500, 
                            color: 'var(--accent-color)', 
                            cursor: 'pointer',
                            textDecoration: 'underline'
                          }}
                          title="Click to view PO details in a new tab"
                        >
                          {po.poNo}
                        </td>
                        <td>{po.date}</td>
                        <td>{po.supplier}</td>
                        <td>{po.creator}</td>
                        <td>{po.dueDate}</td>
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
                {activeData.length > 100 && (
                  <p style={{ color: 'var(--text-secondary)', textAlign: 'center', marginTop: '1.5rem', fontSize: '0.875rem' }}>
                    Showing top 100 results out of {activeData.length}.
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
