import React, { useEffect, useState, useCallback } from 'react';
import { fetchModuleCatalog, fetchAndCleanUnit, ModuleUnit } from '../services/scraper';
import generatePDF from '../services/pdfEngine';
import { Download } from 'lucide-react';
import './popup.css';

type TreeNode = ModuleUnit & { children?: TreeNode[] };

export const Popup: React.FC = () => {
    const [currentUrl, setCurrentUrl] = useState<string | null>(null);
    const [moduleTitle, setModuleTitle] = useState<string | null>(null);
    const [units, setUnits] = useState<ModuleUnit[]>([]);
    const [selected, setSelected] = useState<Record<string, boolean>>({});
    const [loading, setLoading] = useState(false);
    const [progress, setProgress] = useState<{ step: number; total: number } | null>(null);
    const [debugOpen, setDebugOpen] = useState(false);
    const [debugMeta, setDebugMeta] = useState<any>(null);

    useEffect(() => {
        // Get active tab URL
        const getUrl = async () => {
            try {
                const tabs = await new Promise<chrome.tabs.Tab[]>((resolve) =>
                    chrome.tabs.query({ active: true, lastFocusedWindow: true }, resolve)
                );
                const tab = tabs[0];
                const url = tab?.url ?? null;
                setCurrentUrl(url);
            } catch (e) {
                console.error('Failed to query tab', e);
            }
        };
        getUrl();
    }, []);

    useEffect(() => {
        if (!currentUrl) return;
        let mounted = true;
        (async () => {
            setLoading(true);
            const path = currentUrl;
            const meta = await fetchModuleCatalog(path);
            setDebugMeta(meta);
            if (!mounted) return;
            if (meta) {
                setModuleTitle(meta.title ?? null);
                const metaUnits = meta.units || [];
                setUnits(metaUnits);
                const sel: Record<string, boolean> = {};
                metaUnits.forEach((u) => (sel[u.url] = true));
                setSelected(sel);
            }
            setLoading(false);
        })();
        return () => {
            mounted = false;
        };
    }, [currentUrl]);

    const toggle = useCallback((url: string) => {
        setSelected((s) => ({ ...s, [url]: !s[url] }));
    }, []);

    const handleGenerate = async () => {
        const chosen = units.filter((u) => selected[u.url]);
        if (chosen.length === 0) {
            alert('Please select at least one unit');
            return;
        }
        setLoading(true);
        try {
            const cleaned: { title: string; url: string; html: string }[] = [];
            for (let i = 0; i < chosen.length; i++) {
                const u = chosen[i];
                const c = await fetchAndCleanUnit(u.url);
                if (c) cleaned.push({ title: c.title, url: c.url, html: c.html });
            }

            await generatePDF(cleaned, `${moduleTitle || 'learnflow'}.pdf`, (p) => setProgress(p));
        } catch (err) {
            console.error('PDF generation failed', err);
            alert('PDF generation failed: ' + (err as any)?.message || err);
        } finally {
            setLoading(false);
            setProgress(null);
        }
    };

    return (
        <div className="popup">
            <h2 className="header">MS LearnFlow</h2>
            <div className="panel">
                <div className="module">
                    <strong style={{ color: 'var(--primary)' }}>Module:</strong>{' '}
                    {loading && moduleTitle === null ? (
                        <span>Loading module…</span>
                    ) : moduleTitle ? (
                        <span>{moduleTitle}</span>
                    ) : (
                        <span style={{ color: 'var(--muted)' }}>Module title not found</span>
                    )}
                </div>

                <div className="units">
                    {loading && units.length === 0 ? (
                        <div>Loading units…</div>
                    ) : !loading && units.length === 0 ? (
                        <div style={{ color: 'var(--muted)' }}>No units found for this module.</div>
                    ) : (
                        units.map((u) => (
                            <label key={u.url} className="unit">
                                <input className="checkbox" type="checkbox" checked={!!selected[u.url]} onChange={() => toggle(u.url)} />
                                <span>{u.title}</span>
                            </label>
                        ))
                    )}
                </div>

                <div className="controls">
                    <button onClick={handleGenerate} disabled={loading} className="btn btn-primary">
                        <Download size={16} />&nbsp;Generate PDF
                    </button>
                    {loading && <div style={{ marginLeft: 8, color: 'var(--muted)' }}>Working…</div>}
                </div>

                {progress && (
                    <div className="progress">
                        <div style={{ fontSize: 13, color: 'var(--muted)', marginBottom: 6 }}>Page {progress.step} / {progress.total}</div>
                        <div className="progress__bar">
                            <div className="progress__inner" style={{ width: `${(progress.step / progress.total) * 100}%` }} />
                        </div>
                    </div>
                )}

                <div className="tip">Tip: Open a Microsoft Learn module page and click the extension to load units.</div>

                <div className="debugToggle">
                    <button onClick={() => setDebugOpen((s) => !s)} className="btn" style={{ fontSize: 12 }}>
                        {debugOpen ? 'Hide debug' : 'Show debug'}
                    </button>
                    {debugOpen && (
                        <pre className="debugPre" style={{ marginTop: 8 }}>
                            {JSON.stringify(debugMeta, null, 2)}
                        </pre>
                    )}
                </div>
            </div>
        </div>
    );
};

export default Popup;
